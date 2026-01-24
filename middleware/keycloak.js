// keycloak.js
import { createRemoteJWKSet, jwtVerify } from 'jose';

const issuer = process.env.KEYCLOAK_ISSUER;
const audience = process.env.KEYCLOAK_AUDIENCE;

if (!issuer) {
  console.warn('[Keycloak] KEYCLOAK_ISSUER is not set');
}
if (!audience) {
  console.warn('[Keycloak] KEYCLOAK_AUDIENCE is not set');
}

// Remote JWKS from Keycloak
// e.g. https://keycloak.mybank.com/realms/ecollect/protocol/openid-connect/certs
const jwksUri = issuer
  ? new URL(`${issuer.replace(/\/$/, '')}/protocol/openid-connect/certs`)
  : null;

const jwks = jwksUri ? createRemoteJWKSet(jwksUri) : null;

/**
 * Verify a Bearer token against Keycloak.
 * Returns the decoded payload if valid, throws if invalid.
 */
export async function verifyToken(token) {
  if (!jwks) {
    throw new Error('JWKS not configured – check KEYCLOAK_ISSUER');
  }

  const { payload } = await jwtVerify(token, jwks, {
    issuer,
    audience, // clientId
  });

  return payload; // this is the Keycloak token payload
}

/**
 * Express middleware to protect routes with Keycloak.
 * - Expects Authorization: Bearer <token>
 * - On success: sets req.user = decoded token payload
 * - On failure: returns 401/403
 */
export async function keycloakMiddleware(req, res, next) {
  try {
    const auth = req.headers['authorization'] || '';
    if (!auth.startsWith('Bearer ')) {
      return res.status(401).json({ message: 'Missing Bearer token' });
    }

    const token = auth.substring('Bearer '.length).trim();
    if (!token) {
      return res.status(401).json({ message: 'Empty Bearer token' });
    }

    const payload = await verifyToken(token);

    // Attach verified payload for downstream middleware/routes
    req.user = payload;

    next();
  } catch (err) {
    console.error('[Keycloak] token verification failed:', err.message || err);
    return res.status(401).json({ message: 'Invalid or expired token' });
  }
}
