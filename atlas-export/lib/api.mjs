// Shared client for the Atlas (OnAtlas) API.
//
// Spec: https://api.onatlas.com/docs/ (Swagger source: /docs/ra_api.yaml)
//  - All endpoints are GET-only, served from https://api.onatlas.com/api
//  - Auth:    Authorization: Bearer {token}
//  - Every response is an envelope: { meta, page?, data }
//  - Paginated responses carry page.next_page_url — follow until null.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const BASE = 'https://api.onatlas.com/api';

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const DELAY_MS = 200; // politeness gap between paginated/sequential requests
const MAX_ATTEMPTS = 5;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Minimal .env loader so the scripts work on any Node 20+ without flags.
function loadDotEnv() {
  const file = path.join(ROOT, '.env');
  if (!fs.existsSync(file)) return;
  for (const line of fs.readFileSync(file, 'utf8').split('\n')) {
    const m = line.match(/^\s*([A-Za-z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m && !(m[1] in process.env)) {
      process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
    }
  }
}

export function requireToken() {
  loadDotEnv();
  const token = process.env.ATLAS_API_TOKEN;
  if (!token) {
    console.error(
      'Missing ATLAS_API_TOKEN.\n' +
        'Copy .env.example to .env and paste your Atlas API access key, e.g.\n\n' +
        '  cp .env.example .env\n'
    );
    process.exit(1);
  }
  return token;
}

async function fetchJson(url, token) {
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    let res;
    try {
      res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      });
    } catch (err) {
      // Network blip — back off and retry.
      const wait = 1000 * 2 ** (attempt - 1);
      console.warn(`  network error (${err.message}) — retrying in ${wait / 1000}s`);
      await sleep(wait);
      continue;
    }
    if (res.status === 401 || res.status === 403) {
      throw new Error(`HTTP ${res.status} — the API rejected the token. Check ATLAS_API_TOKEN in .env.`);
    }
    if (res.status === 429 || res.status >= 500) {
      const wait = 1000 * 2 ** (attempt - 1);
      console.warn(`  HTTP ${res.status} from ${url} — retrying in ${wait / 1000}s`);
      await sleep(wait);
      continue;
    }
    if (!res.ok) {
      throw new Error(`HTTP ${res.status} ${res.statusText} from ${url}`);
    }
    return res.json();
  }
  throw new Error(`Gave up after ${MAX_ATTEMPTS} attempts: ${url}`);
}

function buildUrl(pathname, params = {}) {
  const url = new URL(BASE + pathname);
  for (const [k, v] of Object.entries(params)) {
    if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
  }
  return url.toString();
}

// One request; returns the full response envelope ({ meta, page?, data }).
export async function apiGet(pathname, params, token) {
  return fetchJson(buildUrl(pathname, params), token);
}

// Follows pagination and returns the concatenated `data` across all pages.
// Single-object responses come back as a one-element array.
export async function apiGetAll(pathname, params, token) {
  const out = [];
  let next = buildUrl(pathname, params);
  while (next) {
    const body = await fetchJson(next, token);
    const data = body && 'data' in body ? body.data : body;
    if (Array.isArray(data)) out.push(...data);
    else if (data !== null && data !== undefined) out.push(data);
    next = body?.page?.next_page_url || null;
    if (next) await sleep(DELAY_MS);
  }
  return out;
}

export { sleep };
