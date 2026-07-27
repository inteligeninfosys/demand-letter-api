import { getLogger } from './logger.js';
import { v4 as uuidv4 } from 'uuid';

export function requestLoggingMiddleware(options = {}) {
  const logger = getLogger(options);

  return function (req, res, next) {
    const requestId =
      req.headers['x-request-id'] ||
      req.headers['x-requestid'] ||
      uuidv4();

    const user =
      req.headers['x-user'] ||
      req.headers['x_username'] ||
      undefined;

    const context = {
      requestId,
      user,
      path: req.originalUrl || req.url,
      method: req.method,
    };

    req.log = logger.withContext(context);
    req.requestId = requestId;
    req.user = user;

    req.log.info('Incoming request', {
      meta: {
        ip: req.ip,
        userAgent: req.headers['user-agent'],
      },
    });

    const start = process.hrtime.bigint();

    res.on('finish', () => {
      const end = process.hrtime.bigint();
      const durationMs = Number(end - start) / 1e6;

      req.log.info('Request completed', {
        meta: {
          statusCode: res.statusCode,
          durationMs,
          contentLength: res.getHeader('content-length') || undefined,
        },
      });
    });

    next();
  };
}
