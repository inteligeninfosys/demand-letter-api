// authAuditMiddleware.js
import jwt from 'jsonwebtoken';
import sql from 'mssql';
import { pool } from '../db.js';

function getIp(req) {
  const xff = req.headers['x-forwarded-for'] || '';
  const first = xff.split(',')[0].trim();
  const raw = first || req.socket.remoteAddress || '';
  return raw.replace('::ffff:', '');
}

export async function authAuditMiddleware(req, res, next) {
  try {
    const auth = req.headers['authorization'];
    if (!auth || !auth.startsWith('Bearer ')) {
      return next();
    }

    const tokenStr = auth.substring('Bearer '.length);

    // Token is already validated by your Keycloak middleware; we just decode
    const decoded = jwt.decode(tokenStr);
    if (!decoded) {
      return next();
    }

    const sessionId = decoded.session_state || decoded.sub;

    if (!sessionId) {
      return next();
    }

    const username =
      decoded.preferred_username ||
      decoded.email ||
      req.headers['x-user'] ||
      'unknown';

    const userId = decoded.sub;
    const realm = decoded.iss || null;
    const clientId = decoded.azp || null;
    const ip = getIp(req);
    const userAgent = req.headers['user-agent'] || null;

    let tokenExp = null;
    if (decoded.exp) {
      tokenExp = new Date(decoded.exp * 1000);
    }


    const request = pool.request();

    request
      .input('SessionId', sql.NVarChar(64), sessionId)
      .input('UserId', sql.NVarChar(64), userId)
      .input('Username', sql.NVarChar(150), username)
      .input('Realm', sql.NVarChar(200), realm)
      .input('ClientId', sql.NVarChar(200), clientId)
      .input('IpAddress', sql.NVarChar(64), ip)
      .input('UserAgent', sql.NVarChar(512), userAgent)
      .input('TokenExp', sql.DateTime2, tokenExp)
      // store token string if you want, or null
      .input('RawToken', sql.NVarChar(sql.MAX), null);

    await request.query(`
      IF EXISTS (SELECT 1 FROM dbo.AuthSessionAudit WHERE SessionId = @SessionId)
      BEGIN
        UPDATE dbo.AuthSessionAudit
           SET LastSeenAt = SYSUTCDATETIME(),
               IpAddress = @IpAddress,
               UserAgent = @UserAgent
         WHERE SessionId = @SessionId;
      END
      ELSE
      BEGIN
        INSERT INTO dbo.AuthSessionAudit (
          SessionId, UserId, Username, Realm, ClientId,
          IpAddress, UserAgent, LoginAt, LastSeenAt, TokenExp, RawToken
        )
        VALUES (
          @SessionId, @UserId, @Username, @Realm, @ClientId,
          @IpAddress, @UserAgent, SYSUTCDATETIME(), SYSUTCDATETIME(),
          @TokenExp, @RawToken
        );
      END
    `);

    // Optional: pass audit info downstream
    req.authAudit = {
      username,
      userId,
      sessionId,
      ip,
    };

    next();
  } catch (err) {
    console.error('authAuditMiddleware error:', err);
    next(); // do not block the request on audit failure
  }
}
