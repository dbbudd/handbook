# Feature Request: Expose Section Teachers/TAs via the hkis-schoology MCP

## Summary

There is currently no way to retrieve the teacher(s) or teaching assistant(s) of a Schoology section through the hkis-schoology MCP server. Teachers and TAs exist in the underlying Schoology data as section members (enrollments) with the role flag `admin === 1`, but the only member-listing tool deliberately discards these rows. As a result, a common and basic administrative question — "who teaches course X?" — cannot be answered through the MCP.

## The Gap

In Schoology, both students and staff are modeled as section members ("enrollments"). The distinction is the `admin` field on each enrollment:

- `admin === 0` → student
- `admin === 1` → teacher / TA / co-teacher (admin-role member)

The MCP exposes `list_section_students` for listing section members, but no tool returns the admin-role members. There is also no tool to fetch a single enrollment by its enrollment ID. Consequently, there is no path through the MCP to identify who teaches a given section.

## Root Cause

`list_section_students` calls the Schoology REST endpoint:

```
GET /v1/sections/{id}/enrollments
```

This endpoint **does** return admin-role members alongside students. However, the tool filters out every row where `admin === 1` before returning its result. The teacher/TA rows are fetched from Schoology and then dropped on the floor. No other tool surfaces them.

## Evidence

For section `7899896088` (course "AP COMPUTER SCIENCE PRINCIPLES", section "7(A-B)"):

- `list_section_students` returned **14 rows**, all with `admin: 0`.
- The returned enrollment IDs were consecutive **except** for three gaps: enrollment IDs `3359254167`, `3359254169`, and `3359254176` were absent from the student list.

Those three missing IDs are the filtered-out admin-role members (the teacher and likely a TA/co-teacher). The gaps in the otherwise-consecutive ID sequence are direct proof that the rows exist at the endpoint but are stripped by the tool before the result is returned.

## Proposed Fix

The underlying endpoint already returns the data; the tool simply needs to stop discarding it. Either of the following resolves the gap:

1. **(Preferred, minimal)** Add an `includeAdmins` (or `includeTeachers`) boolean parameter to `list_section_students`, defaulting to `false` to preserve current behavior. When `true`, admin-role members are included with their `admin` flag intact.
2. **(Alternative)** Add a dedicated `list_section_enrollments` tool that returns *all* members of a section with their `admin` role flag preserved, leaving `list_section_students` unchanged.

Either approach should preserve each row's `admin` flag so callers can distinguish staff from students.

## Related Limitations

These are lower priority but worth addressing alongside the primary fix:

1. **No `school_uid` → numeric-UID resolver.** Tools require a numeric UID, but a username such as `gnolan` cannot be resolved to one. Schoology's `/v1/users?school_uids=gnolan` query would do this, but no tool exposes raw user lookup. (`/v1/users/gnolan/sections` returns 404 because the path segment must be a numeric UID, not a username.)
2. **PII redaction is student-aware only.** Names are returned as opaque IDs (e.g. `Student_xxxx`), and `lookup_anonymised` only de-anonymizes *students* in the operator's roster. Even if a teacher enrollment row were returned, the staff name would not resolve. The redaction layer should be aware of staff vs. student, or the de-anonymizer should handle staff members for admin-scoped callers.

## Impact

Without this capability, the MCP cannot answer "who teaches course X?" — a very common administrative question. The teacher/TA data already flows back from Schoology on every `list_section_students` call; it is being thrown away rather than surfaced.
