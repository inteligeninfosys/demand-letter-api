import 'dotenv/config';
import fs from 'fs';
import https from 'https';
import { createRemoteJWKSet, jwtVerify } from 'jose';

// ───────────────────────────────────────────────────────────────────────────────
// ENV
const ISSUER   = (process.env.KEYCLOAK_ISSUER || '').replace(/\/$/, ''); // e.g. https://keycloak.stima-sacco.local/realms/stima-realm
const AUDIENCE = (process.env.KEYCLOAK_AUDIENCE || '').trim();           // optional: if blank we won't enforce aud
const JWKS_URL = process.env.KEYCLOAK_JWKS_URL || `${ISSUER}/protocol/openid-connect/certs`;

//const INSECURE = (process.env.OIDC_TLS_INSECURE || 'false') === 'true';  // true => skip TLS verify (non-prod only)
const INSECURE = true;
const CA_FILE  = process.env.OIDC_EXTRA_CA_CERTS || process.env.NODE_EXTRA_CA_CERTS || ''; // path to PEM chain
const CLIENT_ID = process.env.OIDC_CLIENT_ID || process.env.KEYCLOAK_CLIENT_ID || '';      // optional (for azp fallback)

// ───────────────────────────────────────────────────────────────────────────────
// HTTPS agent for jose JWKS fetch (handles self-signed or private CA)
const agent = new https.Agent({
  rejectUnauthorized: !INSECURE,
  ca: (CA_FILE && fs.existsSync(CA_FILE)) ? fs.readFileSync(CA_FILE) : undefined,
});

// jose remote JWKS with timeouts & caching
const JWKS = createRemoteJWKSet(new URL(JWKS_URL), {
  agent,
  timeoutDuration: 5000,     // 5s network timeout
  cooldownDuration: 60_000,  // backoff between failed fetches
  cacheMaxAge: 10 * 60_000,  // cache JWKS 10 minutes
});

// Small helper
function extractBearer(req) {
  const h = req.headers.authorization || req.headers.Authorization;
  if (!h || typeof h !== 'string') return null;
  const m = h.match(/^Bearer\s+(.+)$/i);
  return m ? m[1] : null;
}

// Optional relaxed audience check if KEYCLOAK_AUDIENCE is not set
function audienceOK(payload) {
  if (AUDIENCE) {
    const aud = payload.aud;
    if (typeof aud === 'string') return aud === AUDIENCE;
    if (Array.isArray(aud)) return aud.includes(AUDIENCE);
    return false;
  }
  // No AUD enforced: accept if token was issued to our client (azp) when provided
  if (CLIENT_ID) return payload.azp === CLIENT_ID
    || !!payload?.resource_access?.[CLIENT_ID];
  return true;
}

export async function authenticate(req, res, next) {
  try {
    const token = extractBearer(req);
    if (!token) return res.status(401).json({ error: 'missing_bearer_token' });

    // Debug peek (non-fatal)
    try {
      const [, p] = token.split('.');
      const payload = JSON.parse(Buffer.from(p, 'base64').toString('utf8'));
      // Comment out after you’re confident:
      // console.log('JWT payload peek:', { iss: payload.iss, aud: payload.aud, azp: payload.azp, exp: payload.exp });
    } catch {}

    // Build verify options
    const verifyOpts = {
      issuer: ISSUER,
      clockTolerance: '5s',
      ...(AUDIENCE ? { audience: AUDIENCE } : {}), // enforce aud only if configured
    };

    const { payload /*, protectedHeader*/ } = await jwtVerify(token, JWKS, verifyOpts);

    // If AUDIENCE not enforced above, do a relaxed check against CLIENT_ID/azp
    if (!audienceOK(payload)) {
      return res.status(401).json({ error: 'unexpected_audience', aud: payload.aud, azp: payload.azp });
    }

    req.user = {
      id: payload.sub,
      username: payload.preferred_username || payload.preferredUsername || payload.email || payload.sub,
      name: payload.name || null,
      email: payload.email || null,
      roles: payload.realm_access?.roles || [],
      resource_access: payload.resource_access || {},
    };

    return next();
  } catch (err) {
    const msg = err?.message || String(err);

    // Network/TLS/DNS issues fetching JWKS
    if (/fetch failed|getaddrinfo|ENOTFOUND|ECONNREFUSED|ECONNRESET|CERT|self[- ]signed|UNABLE_TO_VERIFY/i.test(msg)) {
      console.error('[auth] JWKS fetch error:', msg, 'issuer:', ISSUER, 'jwks:', JWKS_URL, 'insecure:', INSECURE, 'cafile:', CA_FILE || '(none)');
      return res.status(503).json({ error: 'jwks_unreachable', detail: msg });
    }

    console.error('Auth failed:', {
      code: err?.code,
      claim: err?.claim,
      reason: err?.reason,
      message: msg,
    });
    return res.status(401).json({ error: 'invalid_token' });
  }
}
