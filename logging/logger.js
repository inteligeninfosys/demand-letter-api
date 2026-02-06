import amqp from 'amqplib';

const LEVELS = ['debug', 'info', 'warn', 'error'];

function safeJson(x) {
  const seen = new WeakSet();
  try {
    return JSON.stringify(x, (k, v) => {
      if (typeof v === "object" && v !== null) {
        if (seen.has(v)) return "[Circular]";
        seen.add(v);
      }
      if (v instanceof Error) {
        return { name: v.name, message: v.message, stack: v.stack };
      }
      return v;
    });
  } catch (e) {
    // last resort
    return String(x);
  }
}

class Logger {
  constructor(config = {}) {
    this.serviceName = config.serviceName || process.env.SERVICE_NAME || 'demands-service';
    this.env = config.env || process.env.NODE_ENV || 'development';
    this.rabbitmqUrl = config.rabbitmqUrl || process.env.RABBITMQ_URL || 'amqp://guest:guest@localhost:5672';
    this.appVersion = config.appVersion || process.env.APP_VERSION || '0.0.0';
    this.queue = config.queue || process.env.LOG_QUEUE || 'ecollect.logs.queue';

    this.conn = null;
    this.channel = null;
    this.ready = false;
    this.buffer = [];

    this._connectRabbitmq().catch(err => {
      console.error('[logger] Failed to connect to RabbitMQ', err);
    });

    LEVELS.forEach(level => {
      this[level] = (msg, meta) => this.log(level, msg, meta);
    });
  }

  async _connectRabbitmq() {
    this.conn = await amqp.connect(this.rabbitmqUrl);
    this.channel = await this.conn.createChannel();

    await this.channel.assertQueue(this.queue, {
      durable: true,
    });

    this.ready = true;

    if (this.buffer.length > 0) {
      for (const entry of this.buffer) {
        this._publish(entry);
      }
      this.buffer = [];
    }

    this.conn.on('close', () => {
      this.ready = false;
      console.error('[logger] RabbitMQ connection closed');
    });

    this.conn.on('error', (err) => {
      this.ready = false;
      console.error('[logger] RabbitMQ connection error', err);
    });
  }

  log(level, message, meta = {}) {
    try {
      const now = new Date();
      const logEntry = {
        '@timestamp': now.toISOString(),
        level,
        message,
        service: this.serviceName,
        env: this.env,
        version: this.appVersion,
        request_id: meta.requestId || meta.request_id || undefined,
        user: meta.user || meta.username || undefined,
        path: meta.path,
        method: meta.method,
        meta: meta.meta || meta.meta === null ? meta.meta : {
          ...meta,
          requestId: undefined,
          request_id: undefined,
          user: undefined,
          username: undefined,
        },
      };

      this._logToConsole(logEntry);

      if (this.ready && this.channel) {
        this._publish(logEntry);
      } else {
        this.buffer.push(logEntry);
      }
    } catch (err) {
      console.error('[logger] Failed to log message', err);
    }
  }

  _logToConsole(entry) {
    const line = JSON.stringify(entry);
    if (entry.level === 'error') {
      process.stderr.write(line + '\n');
    } else {
      process.stdout.write(line + '\n');
    }
  }

  _publish(entry) {
    try {
      // 🔥 Send to queue instead of exchange
      this.channel.sendToQueue(this.queue, Buffer.from(JSON.stringify(entry)), {
        contentType: 'application/json',
        persistent: true,
      });
    } catch (err) {
      console.error('[logger] Failed to publish log', err);
    }
  }

  withContext(ctx = {}) {
    const base = {
      requestId: ctx.requestId,
      request_id: ctx.request_id,
      user: ctx.user,
      username: ctx.username,
      path: ctx.path,
      method: ctx.method,
    };

    const child = {};
    LEVELS.forEach(level => {
      child[level] = (msg, meta = {}) =>
        this.log(level, msg, { ...base, ...meta });
    });

    child.log = (level, msg, meta = {}) =>
      this.log(level, msg, { ...base, ...meta });

    return child;
  }

  patchConsole() {
    const origLog = console.log;
    const origError = console.error;
    const origWarn = console.warn;
    const origDebug = console.debug;

    console.log = (...args) => {
      //const msg = args.map(a => (typeof a === 'string' ? a : JSON.stringify(a))).join(' ');
      const msg = args.map(a => (typeof a === 'string' ? a : safeJson(a))).join(' ');
      this.info(msg);
      origLog.apply(console, args);
    };

    console.error = (...args) => {
      const msg = args.map(a => (typeof a === 'string' ? a : safeJson(a))).join(' ');
      this.error(msg);
      origError.apply(console, args);
    };

    console.warn = (...args) => {
      const msg = args.map(a => (typeof a === 'string' ? a : safeJson(a))).join(' ');
      this.warn(msg);
      origWarn.apply(console, args);
    };

    console.debug = (...args) => {
      const msg = args.map(a => (typeof a === 'string' ? a : safeJson(a))).join(' ');
      this.debug(msg);
      origDebug.apply(console, args);
    };
  }
}

let defaultLogger;

/**
 * Get singleton logger instance.
 */
export function getLogger(config) {
  if (!defaultLogger) {
    defaultLogger = new Logger(config);
    defaultLogger.patchConsole();
  }
  return defaultLogger;
}

export { Logger };
