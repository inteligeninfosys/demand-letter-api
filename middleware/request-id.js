import crypto from 'crypto';

function requestIdMiddleware(req, res, next) {
  const incoming = req.header('X-Request-Id');
  const requestId = incoming || crypto.randomUUID();

  // attach to request object
  req.requestId = requestId;

  // returned to caller as well
  res.setHeader('X-Request-Id', requestId);

  // timing
  const start = process.hrtime.bigint();

  res.on('finish', () => {
    const durationNs = Number(process.hrtime.bigint() - start);
    const durationMs = durationNs / 1e6;

    // simple console log (will be picked up by Fluentbit/Filebeat)
    console.log(JSON.stringify({
      level: 'info',
      msg: 'request_completed',
      service: process.env.SERVICE_NAME || 'recoveries-api',
      requestId,
      method: req.method,
      path: req.originalUrl || req.url,
      statusCode: res.statusCode,
      durationMs,
      user: req.user?.username || null, // if you fill it from Keycloak
      timestamp: new Date().toISOString()
    }));
  });

  next();
}

export default requestIdMiddleware;
