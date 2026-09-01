// server.js
import 'dotenv/config';
import express from "express";
import fs from "fs/promises";
import cors from 'cors';
import path from "path";
import { fileURLToPath } from "url";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc.js";
import { execFile } from "child_process";
import { promisify } from "util";
import multer from "multer";
import { createRequire } from "module";
import os from "os";
import crypto from "crypto";
import nodemailer from "nodemailer";
import sql from "mssql";
import { getLogger } from './logging/logger.js';
import { requestLoggingMiddleware } from './logging/express-middleware.js';
import requestIdMiddleware from "./middleware/request-id.js";

import { authenticate } from './auth.js';

dayjs.extend(utc);
const require = createRequire(import.meta.url);


import { S3Client, PutObjectCommand, GetObjectCommand } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";


const __dirname = path.dirname(fileURLToPath(import.meta.url));


const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: "4mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(requestIdMiddleware);
// init shared logger once
getLogger({
  serviceName: 'demands-api',
  rabbitmqUrl: process.env.RABBITMQ_URL,
});
// attach logging middleware here so every request has req.log
app.use(requestLoggingMiddleware());

const TEMPLATES_DIR = path.resolve(process.env.TEMPLATE_DIR || path.join(__dirname, "templates"));
const upload = multer({ storage: multer.memoryStorage() });

/* ---------- Helpers ---------- */

const execFileAsync = promisify(execFile);
const safeCode = (s) => String(s || "").trim().replace(/[^a-zA-Z0-9_\-]/g, "");
async function exists(p) { try { await fs.access(p); return true; } catch { return false; } }

async function listTemplates() {
  if (!(await exists(TEMPLATES_DIR))) return [];
  const entries = await fs.readdir(TEMPLATES_DIR, { withFileTypes: true });
  const out = [];
  for (const d of entries) {
    if (!d.isDirectory()) continue;
    const code = d.name;
    const dir = path.join(TEMPLATES_DIR, code);
    const files = (await fs.readdir(dir).catch(() => [])) || [];
    const versions = files.filter((f) => /\.docx$/i.test(f)).sort();
    out.push({
      code,
      hasCurrent: versions.includes("current.docx"),
      versions,
      meta: await readMeta(code).catch(() => null),
    });
  }
  return out;
}

async function readMeta(code) {
  const p = path.join(TEMPLATES_DIR, code, "meta.json");
  if (!(await exists(p))) return null;
  return JSON.parse(await fs.readFile(p, "utf8"));
}

// Robust resolver
async function resolveTemplatePath(template_code, template_version = null) {
  if (!template_code) throw new Error("template_code is required");

  // 1) sanitize path segments to prevent directory traversal
  const code = String(template_code).trim();                // e.g. 'DL1'
  const verIn = (template_version ?? "current").toString().trim();
  if (!/^[a-zA-Z0-9_-]+$/.test(code)) {
    throw new Error("Invalid template_code");
  }
  if (!/^[a-zA-Z0-9_.-]+$/.test(verIn) || verIn.includes("..")) {
    throw new Error("Invalid template_version");
  }

  // If caller passed a full filename like "current.docx" or "DL1_v2.docx", use it as-is
  const isFileName = /\.(docx)$/i.test(verIn);
  const fileName = isFileName ? verIn : `${verIn}.docx`;    // "current" -> "current.docx"

  // 2) candidate roots (absolute!)
  const roots = [
    TEMPLATES_DIR,                                    // env TEMPLATE_DIR or ./templates
    "/app/templates",                               // default in your image
    "/data/templates",                              // optional external mount
  ].filter((value, index, all) => value && all.indexOf(value) === index);

  // 3) build candidates
  const tried = [];
  for (const root of roots) {
    tried.push(
      path.join(root, code, fileName),                // /app/templates/DL1/current.docx
      path.join(root, `${code}.docx`)                 // /app/templates/DL1.docx (fallback)
    );
  }

  // 4) return first existing
  for (const p of tried) {
    if (await exists(p)) return p;
  }

  // 5) diagnostics
  const root = roots[0];
  let listing = "(missing)";
  try {
    const dirs = await fs.readdir(root, { withFileTypes: true });
    listing = dirs.map(d => (d.isDirectory() ? `${d.name}/` : d.name)).join(", ");
  } catch { /* ignore */ }

  const msg =
    `Template not found for code='${code}', version='${verIn}'.\n` +
    `CWD=${process.cwd()}, __dirname=${__dirname}, TEMPLATE_DIR=${process.env.TEMPLATE_DIR || "(unset)"}\n` +
    `Tried:\n${tried.map(t => ` - ${t}`).join("\n")}\n` +
    `Listing of first root (${root}): ${listing}`;
  throw new Error(msg);
}

async function renderDocxFromTemplate(templatePath, data) {
  const content = await fs.readFile(templatePath);
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    //delimiters: { start: "[[", end: "]]" }, // matches your templates
    nullGetter: () => "",                       // return empty string for missing values
  });

  const safe = (v) => (v === null || v === undefined ? "" : v);
  function sanitize(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "string") return v.trim();
    if (Array.isArray(v)) return v.map(sanitize);
    if (typeof v === "object") {
      const o = {};
      for (const k of Object.keys(v)) o[k] = sanitize(v[k]);
      return o;
    }
    return v;
  }

  const model = sanitize({
    ...data,
    customer: {
      ...data?.customer,
      // trim padded account numbers from core banking
      account_number: (data?.customer?.account_number || "").toString().trim(),
      customer_number: data?.customer?.customer_number ?? "",
    },
    loan: {
      ...data?.loan,
      // if you have numeric copies, keep them; else keep strings
      days_in_arrears: data?.loan?.days_in_arrears ?? "",
      outstanding_balance: data?.loan?.outstanding_balance ?? "",
    },
    guarantors: Array.isArray(data?.guarantors) ? data.guarantors : [],
  });


  // Helpful diagnostics for common mistakes
  // (a) quick presence check for keys you mentioned
  const dbg = {
    "customer.name": model?.customer?.name,
    "customer.account_number": model?.customer?.account_number,
    "loan.outstanding_balance": model?.loan?.outstanding_balance,
    "loan.days_in_arrears": model?.loan?.days_in_arrears,
  };
  Object.entries(dbg).forEach(([k, v]) => {
    if (v === "") console.warn(`[DOCX] value empty for tag: ${k}`);
  });


  doc.render(model);
  return doc.getZip().generate({ type: "nodebuffer" });
}

async function docxToPdfBuffer(docxBuffer) {
  const soffice = await resolveSoffice();
  return withTempDir(async (dir) => {
    const inPath = path.join(dir, `in-${crypto.randomUUID()}.docx`);
    await fs.writeFile(inPath, docxBuffer);

    await execFileAsync(
      soffice,
      [
        "--headless",
        "--nologo",
        "--nodefault",
        "--norestore",
        "--nolockcheck",
        "--convert-to", "pdf",
        "--outdir", dir,
        inPath,
      ],
      { windowsHide: true }
    );

    const pdfPath = inPath.replace(/\.docx$/i, ".pdf");
    const pdf = await fs.readFile(pdfPath);
    return pdf;
  });
}



// Convert a PDF buffer to a PNG buffer using `pdftoppm`
async function pdfToPngBuffer(pdfBuffer, { page = 1, dpi = 144 } = {}) {
  const pdftoppm = await resolvePdftoppm();
  return withTempDir(async (dir) => {
    const inPath = path.join(dir, `in-${crypto.randomUUID()}.pdf`);
    const outBase = path.join(dir, `out-${crypto.randomUUID()}`);
    await fs.writeFile(inPath, pdfBuffer);

    await execFileAsync(
      pdftoppm,
      ["-png", "-rx", String(dpi), "-ry", String(dpi), "-f", String(page), "-l", String(page), "-singlefile", inPath, outBase],
      { windowsHide: true }
    );

    const pngPath = `${outBase}.png`;
    const png = await fs.readFile(pngPath);
    return png;
  });
}

async function saveLetterToMinioAndLog({
  template_code,           // e.g. 'demand1'
  data,                    // model used to render
  blob,                    // Buffer (PDF or DOCX)
  ext,                     // 'pdf' | 'docx'
  contentType,             // mime
  sent_by,                 // from Keycloak / request header
  provider_ref,            // e.g., email messageId (optional)
  our_ref,
  status = "SAVED",        // or "SENT"
}) {
  const traceId = crypto.randomUUID();
  const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
  const idem_key = generateIdemKey(template_code, account);

  const tmpl = (template_code || "demand").replace(/[^\w.-]+/g, "_");
  const ts = dayjs().format("YYYY/MM/DD");
  const tsName = dayjs().format("YYYYMMDD_HHmmss");
  const document_name = `${account}_${tmpl}_${tsName}.${ext}`;
  const object_key = `letters/${tmpl}/${ts}/${document_name}`;

  console.log("[saveLetterToMinioAndLog] STEP 1 start", {
    traceId,
    template_code,
    account,
    ext,
    contentType,
    status,
    hasBlob: !!blob,
    blobLength: Buffer.isBuffer(blob) ? blob.length : null,
    document_name,
    object_key,
  });

  let bucket;
  let key;
  try {
    console.log("[saveLetterToMinioAndLog] STEP 2 uploadToS3 begin", { traceId, object_key });
    const uploaded = await uploadToS3({ key: object_key, body: blob, contentType });
    bucket = uploaded?.bucket;
    key = uploaded?.key;
    console.log("[saveLetterToMinioAndLog] STEP 3 uploadToS3 success", { traceId, bucket, key });
  } catch (e) {
    console.error("[saveLetterToMinioAndLog] STEP 3 uploadToS3 failed", {
      traceId,
      message: e?.message,
      name: e?.name,
      code: e?.code,
      statusCode: e?.statusCode,
      stack: e?.stack,
      bucket: process.env.S3_BUCKET || null,
      endpoint: process.env.S3_ENDPOINT || null,
      object_key,
    });
    throw e;
  }

  let signedUrl;
  let signedUrlExpiryUtc;
  try {
    console.log("[saveLetterToMinioAndLog] STEP 4 presignGet begin", { traceId, bucket, key });
    signedUrl = await presignGet({ bucket, key, expiresInSec: process.env.S3_SIGN_URL_EXP_SECONDS });
    signedUrlExpiryUtc = dayjs().add(Number(process.env.S3_SIGN_URL_EXP_SECONDS || 3600), "second")
      .toDate();
    console.log("[saveLetterToMinioAndLog] STEP 5 presignGet success", {
      traceId,
      bucket,
      key,
      signedUrlExpiryUtc,
      hasSignedUrl: !!signedUrl,
    });
  } catch (e) {
    console.error("[saveLetterToMinioAndLog] STEP 5 presignGet failed", {
      traceId,
      message: e?.message,
      name: e?.name,
      code: e?.code,
      statusCode: e?.statusCode,
      stack: e?.stack,
      bucket,
      key,
    });
    throw e;
  }

  let id;
  try {
    console.log("[saveLetterToMinioAndLog] STEP 6 insertHistory begin", {
      traceId,
      account_number: data?.customer?.account_number || null,
      customer_number: data?.customer?.customer_number || null,
      demand_type: template_code,
      document_name,
      bucket,
      key,
      provider_ref,
      our_ref,
      status,
      idem_key,
    });

    id = await insertHistory({
      account_number: data?.customer?.account_number || null,
      customer_number: data?.customer?.customer_number || null,
      demand_type: template_code,
      date_sent: new Date(),
      days_in_arrears: data?.loan?.days_in_arrears ?? null,
      outstanding_balance: (data?.loan?.outstanding_balance ?? null),
      arrears_amount: (data?.loan?.arrears_amount ?? null),
      sent_by,
      document_name,
      bucket,
      object_key: key,
      provider_ref: provider_ref,
      our_ref,
      status,
      signed_url_expiry_utc: signedUrlExpiryUtc,
      idem_key,
    });

    console.log("[saveLetterToMinioAndLog] STEP 7 insertHistory success", { traceId, id, bucket, key });
  } catch (e) {
    console.error("[saveLetterToMinioAndLog] STEP 7 insertHistory failed", {
      traceId,
      message: e?.message,
      name: e?.name,
      code: e?.code,
      number: e?.number,
      state: e?.state,
      class: e?.class,
      lineNumber: e?.lineNumber,
      serverName: e?.serverName,
      procName: e?.procName,
      stack: e?.stack,
      account_number: data?.customer?.account_number || null,
      customer_number: data?.customer?.customer_number || null,
      document_name,
      bucket,
      key,
      provider_ref,
      our_ref,
      status,
    });
    throw e;
  }

  console.log("[saveLetterToMinioAndLog] STEP 8 done", { traceId, id, bucket, key, document_name });
  return { id, bucket, key, document_name, signedUrl, signedUrlExpiryUtc };
}

// --- MSSQL
let _sqlPool = null;
async function getSqlPool() {
  if (_sqlPool) return _sqlPool;
  _sqlPool = await new sql.ConnectionPool({
    server: process.env.MSSQL_SERVER,
    port: parseInt(process.env.MSSQL_PORT, 10) || 1435,
    database: process.env.MSSQL_DATABASE,
    user: process.env.MSSQL_USER,
    password: process.env.MSSQL_PASSWORD,
    options: {
      encrypt: String(process.env.MSSQL_ENCRYPT || "false") === "true",
      trustServerCertificate: String(process.env.MSSQL_TRUST_SERVER_CERTIFICATE || "true") === "true",
      enableArithAbort: true,
    },
  }).connect();
  return _sqlPool;
}
function toDecimalOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const parsed = Number(String(value).replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

async function insertHistory({
  account_number, customer_number, demand_type, date_sent,
  days_in_arrears, outstanding_balance, arrears_amount, sent_by,
  document_name, bucket, object_key, provider_ref, our_ref, status, signed_url_expiry_utc, idem_key
}) {
  const pool = await getSqlPool();
  const r = await pool.request()
    .input("account_number", sql.NVarChar(100), account_number)
    .input("customer_number", sql.NVarChar(100), customer_number || null)
    .input("demand_type", sql.NVarChar(50), demand_type)
    .input("date_sent", sql.DateTime2(0), date_sent)
    .input("days_in_arrears", sql.Int, days_in_arrears ?? null)
    .input("outstanding_balance", sql.Decimal(18, 2), toDecimalOrNull(outstanding_balance))
    .input("arrears_amount", sql.Decimal(18, 2), toDecimalOrNull(arrears_amount))
    .input("sent_by", sql.NVarChar(128), sent_by || null)
    .input("document_name", sql.NVarChar(260), document_name)
    .input("bucket", sql.NVarChar(128), bucket)
    .input("object_key", sql.NVarChar(512), object_key)
    .input("provider_ref", sql.NVarChar(200), provider_ref || null)
    .input("our_ref", sql.NVarChar(120), our_ref || null)
    .input("status", sql.NVarChar(30), status || "SAVED")
    .input("signed_url_expiry_utc", sql.DateTime2(0), signed_url_expiry_utc || null)
    .input("idem_key", sql.NVarChar(100), idem_key || null)
    .query(`
      INSERT INTO dbo.demand_letter_history
      (account_number, customer_number, demand_type, date_sent, days_in_arrears, outstanding_balance,arrears_amount,
       sent_by, document_name, bucket, object_key, provider_ref, our_ref, status, signed_url_expiry_utc, idem_key)
      OUTPUT inserted.id
      VALUES (@account_number, @customer_number, @demand_type, @date_sent, @days_in_arrears, @outstanding_balance,@arrears_amount,
              @sent_by, @document_name, @bucket, @object_key, @provider_ref,@our_ref, @status, @signed_url_expiry_utc, @idem_key)
    `);
  return r.recordset?.[0]?.id;
}

// ---- S3/MinIO client + upload + signed URL
const S3 = new S3Client({
  region: process.env.S3_REGION || "us-east-1",
  endpoint: process.env.S3_ENDPOINT || undefined,
  forcePathStyle: String(process.env.S3_FORCE_PATH_STYLE || "true") === "true",
  credentials: (process.env.S3_ACCESS_KEY && process.env.S3_SECRET_KEY) ? {
    accessKeyId: process.env.S3_ACCESS_KEY,
    secretAccessKey: process.env.S3_SECRET_KEY,
  } : undefined,
});

async function uploadToS3({ key, body, contentType }) {
  const Bucket = process.env.S3_BUCKET;
  await S3.send(new PutObjectCommand({ Bucket, Key: key, Body: body, ContentType: contentType }));
  return { bucket: Bucket, key };
}
async function presignGet({ bucket, key, expiresInSec }) {
  return getSignedUrl(S3, new GetObjectCommand({ Bucket: bucket, Key: key }),
    { expiresIn: Number(expiresInSec || process.env.S3_SIGN_URL_EXP_SECONDS || 3600) });
}

async function withTempDir(run) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "demand-"));
  try {
    return await run(dir);
  } finally {
    // best-effort cleanup
    try { await fs.rm(dir, { recursive: true, force: true }); } catch { }
  }
}

async function resolveSoffice() {
  const candidates = process.platform === "win32"
    ? [
      "C:\\Program Files\\LibreOffice\\program\\soffice.com",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
      "soffice" // last resort if PATH is set
    ]
    : ["soffice"];

  for (const c of candidates) {
    try {
      await execFileAsync(c, ["--version"], { windowsHide: true });
      return c;
    } catch { /* try next */ }
  }
  throw new Error("LibreOffice (soffice) not found. Install LibreOffice and ensure it's on PATH.");
}

async function resolvePdftoppm() {
  const candidates = process.platform === "win32"
    ? ["pdftoppm"] // ensure poppler is installed and on PATH (e.g., via Chocolatey)
    : ["pdftoppm"];
  for (const c of candidates) {
    try {
      await execFileAsync(c, ["-v"], { windowsHide: true });
      return c;
    } catch { }
  }
  throw new Error("pdftoppm not found. Install poppler-utils and ensure it's on PATH to enable PNG previews.");
}

function makeMailer() {
  const host = process.env.EMAIL_HOST;
  const port = Number(process.env.EMAIL_PORT || 587);
  const secure = String(process.env.EMAIL_SECURE || "false") === "true";
  const user = process.env.EMAIL_USER;
  const pass = process.env.EMAIL_PASS;
  const from = process.env.EMAIL_FROM || user;


  if (!host || !user || !pass) {
    throw new Error("Email not configured: set EMAIL_HOST, EMAIL_USER, EMAIL_PASS");
  }

  const transport = nodemailer.createTransport({
    host, port, secure,
    auth: { user, pass },
  });

  return { transport, from };
}

function generateIdemKey(template_code, account_number) {
  const t = (template_code || "DEMAND").replace(/[^\w.-]+/g, "_").toUpperCase();
  const acc = (account_number || "UNKNOWN").replace(/[^\w.-]+/g, "_").toUpperCase();
  const ts = dayjs().utc().format("YYYYMMDDTHHmmss[Z]");
  const rand = crypto.randomBytes(3).toString("hex"); // 6-char random suffix
  return `${t}_${acc}_${ts}_${rand}`;
}

// If the sequence exists, we’ll use it; else fallback to time+random (still unique).
async function generateOurRef({ template_code, account_number, customer_number }) {
  const prefix = (process.env.OUR_REF_PREFIX || "KB/REC").trim();
  const tmpl = (template_code || "DEMAND").toUpperCase().replace(/[^\w/-]+/g, "");
  const yyyy = dayjs().utc().format("YYYY");

  // For DL1, append the customer_number as the final ref segment.
  const custSuffix =
    tmpl === "DL1" && customer_number ? `/${String(customer_number).trim()}` : "";

  let seq = null;
  try {
    const pool = await getSqlPool();
    const r = await pool.request().query("SELECT NEXT VALUE FOR dbo.seq_demand_ref AS seq");
    seq = r?.recordset?.[0]?.seq;
  } catch {
    // sequence missing → fallback
  }

  if (!seq) {
    const ts = dayjs().utc().format("YYYYMMDDHHmmss");
    const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
    return `${prefix}/${tmpl}/${yyyy}/${ts}-${rand}${custSuffix}`; // e.g. KB/REC/DEMAND1/2025/20251107...-ABCD
  }

  return `${prefix}/${tmpl}/${yyyy}/${seq}${custSuffix}`; // e.g. KB/REC/DL1/2025/100321/301952171
}


/* ---------- Repossession order helpers ---------- */

function ordinalDay(day) {
  const n = Number(day);
  const mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return `${n}TH`;
  switch (n % 10) {
    case 1: return `${n}ST`;
    case 2: return `${n}ND`;
    case 3: return `${n}RD`;
    default: return `${n}TH`;
  }
}

// e.g. "MONDAY 3rd AUGUST 2026" — day name & month uppercase, ordinal suffix lowercase
function formatFullDateMixedCase(value = null) {
  const d = value ? dayjs(value) : dayjs();
  if (!d.isValid()) return value;
  const day = d.date();
  const suffix = ordinalDay(day).replace(/\d+/, "").toLowerCase();
  return `${d.format("dddd").toUpperCase()} ${day}${suffix} ${d.format("MMMM").toUpperCase()} ${d.format("YYYY")}`;
}

function formatRepossessionDate(value = null) {
  if (typeof value === "string") {
    const text = value.trim();
    if (/^[A-Za-z]+\s+\d{1,2}(?:ST|ND|RD|TH)\s+[A-Za-z]+\s+\d{4}$/i.test(text)) {
      return text.toUpperCase();
    }
  }

  const d = value ? dayjs(value) : dayjs();
  if (!d.isValid()) throw new Error("Invalid repossession order date");
  return `${d.format("dddd")} ${ordinalDay(d.date())} ${d.format("MMMM YYYY")}`.toUpperCase();
}

function formatMoney(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return new Intl.NumberFormat("en-KE", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(value);
  }

  const text = String(value).trim();
  const parsed = Number(text.replace(/,/g, ""));
  if (Number.isFinite(parsed)) {
    return new Intl.NumberFormat("en-KE", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(parsed);
  }
  return text;
}

// Renders a money value with its sign dropped (e.g. "-1.28" -> "1.28")
function formatAbsMoney(value) {
  if (value === null || value === undefined || value === "") return "";
  const num = typeof value === "number" ? value : Number(String(value).trim().replace(/,/g, ""));
  if (!Number.isFinite(num)) return value;
  return formatMoney(Math.abs(num));
}

// DPD -> classification band, per bank policy
const DPD_CLASSIFICATION_BANDS = [
  { min: 0, max: 30, label: "Normal" },
  { min: 31, max: 60, label: "Watch 1" },
  { min: 61, max: 90, label: "Watch 2" },
  { min: 91, max: 120, label: "Substandard 1" },
  { min: 121, max: 180, label: "Substandard 2" },
  { min: 181, max: 360, label: "Doubtful" },
  { min: 361, max: Infinity, label: "Loss" },
];

function classifyByDaysPastDue(days) {
  const n = Number(days);
  if (!Number.isFinite(n) || n < 0) return "";
  const band = DPD_CLASSIFICATION_BANDS.find((b) => n >= b.min && n <= b.max);
  return band ? band.label : "";
}

function formatInterestRate(value) {
  if (value === null || value === undefined || value === "") return "";
  const text = String(value).trim();
  return text.endsWith("%") ? text : `${text}%`;
}

function numberToWordsUnder100(value) {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 0 || n > 99) return String(value || "").toUpperCase();
  const ones = ["ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"];
  const tens = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"];
  if (n < 20) return ones[n];
  const t = Math.floor(n / 10);
  const o = n % 10;
  return o ? `${tens[t]}-${ones[o]}` : tens[t];
}

function actorFromRequest(req) {
  const value =
    req.user?.preferred_username ||
    req.user?.username ||
    req.user?.email ||
    req.headers["x-user"] ||
    "unknown";
  return typeof value === "string" ? value : String(value?.name || value?.username || "unknown");
}

function buildRepossessionModel(input = {}) {
  const data = input && typeof input === "object" ? input : {};
  const validityDays = Number(data?.repossession?.validity_days ?? 30);

  return {
    ...data,
    date: data.date || formatRepossessionDate(),
    bank: {
      name: "KINGDOM BANK LIMITED",
      address_line_1: "Kingdom Bank Towers, Argwings Kodhek Rd, Kilimani",
      address_line_2: "P. O. Box 22741-00100",
      town: "Nairobi",
      ...(data.bank || {}),
    },
    auctioneer: {
      name: "",
      address_line_1: "",
      address_line_2: "",
      town: "",
      contact_person: "",
      ...(data.auctioneer || {}),
    },
    customer: {
      account_number: "",
      customer_number: "",
      name: "",
      address_line_1: "",
      address_line_2: "",
      town: "",
      phone: "",
      ...(data.customer || {}),
    },
    collateral: {
      physical_address: "",
      legal_description: "",
      ...(data.collateral || {}),
    },
    loan: {
      ...(data.loan || {}),
      outstanding_balance: formatMoney(data?.loan?.outstanding_balance),
      interest_rate: formatInterestRate(data?.loan?.interest_rate),
    },
    repossession: {
      statutory_provision: "Movable Property Security Rights Act",
      legal_costs: "T.B.A",
      auctioneer_fees: "AUCTIONEERS SCALE",
      reserve_price: "",
      advertising_instructions: "To be assessed",
      validity_days: validityDays,
      validity_days_words:
        data?.repossession?.validity_days_words || numberToWordsUnder100(validityDays),
      ...(data.repossession || {}),
    },
    signatory_1: {
      name: "Samuel Murimi",
      title: "Debt Recovery Manager",
      ...(data.signatory_1 || {}),
    },
    signatory_2: {
      name: "Josphat Thiaine",
      title: "Head of Credit",
      ...(data.signatory_2 || {}),
    },
  };
}

function validateRepossessionModel(data) {
  const required = [
    ["customer.account_number", data?.customer?.account_number],
    ["customer.name", data?.customer?.name],
    ["auctioneer.name", data?.auctioneer?.name],
    ["collateral.legal_description", data?.collateral?.legal_description],
    ["loan.outstanding_balance", data?.loan?.outstanding_balance],
  ];

  const missing = required
    .filter(([, value]) => value === null || value === undefined || String(value).trim() === "")
    .map(([field]) => field);

  if (missing.length) {
    const error = new Error(`Missing required repossession fields: ${missing.join(", ")}`);
    error.status = 400;
    throw error;
  }
}

async function generateRepossessionOurRef() {
  const prefix = (process.env.REPO_REF_PREFIX || "KB/DRU_MM").trim().replace(/\/+$/, "");
  const month = dayjs().utc().format("M");
  const year = dayjs().utc().format("YYYY");

  let seq = null;
  try {
    const pool = await getSqlPool();
    const r = await pool.request().query("SELECT NEXT VALUE FOR dbo.seq_demand_ref AS seq");
    seq = r?.recordset?.[0]?.seq;
  } catch {
    // Sequence may not exist in a fresh environment; use a unique time-based suffix.
  }

  const suffix = seq || `${dayjs().utc().format("YYYYMMDDHHmmss")}${Math.floor(Math.random() * 90 + 10)}`;
  return `${prefix}/${month}/${year}/${suffix}`;
}

/* ---------- Tiny cache (per code+version) ---------- */

const cache = new Map(); // key = `${code}:${version||"current"}`
async function getTemplateBuffer(code, version) {
  const key = `${safeCode(code)}:${version || "current"}`;
  if (cache.has(key)) return cache.get(key);
  const p = await resolveTemplatePath(code, version);
  const buf = await fs.readFile(p);
  cache.set(key, buf);
  return buf;
}

/* ---------- Routes ---------- */


// List templates
app.get("/demand-letters-api/templates", async (_req, res) => {
  try {
    const pool = await getSqlPool(); 

    // Get metadata from database
    const dbResult = await pool.request().query(`
            SELECT 
                template_code,
                template_name,
                description,
                available_fields,
                is_active
            FROM dbo.demand_letter_template
            ORDER BY template_name
        `);

    const dbTemplates = dbResult.recordset || [];
    const templates = [];

    console.log(dbTemplates)

    // Check filesystem for versions
    if (await exists(TEMPLATES_DIR)) {
      const entries = await fs.readdir(TEMPLATES_DIR, { withFileTypes: true });

      for (const d of entries) {
        if (!d.isDirectory()) continue;
        const code = d.name;
        const dir = path.join(TEMPLATES_DIR, code);
        const files = await fs.readdir(dir).catch(() => []);
        const versions = files.filter((f) => /\.docx$/i.test(f)).sort();

        // Find matching database record
        const dbRecord = dbTemplates.find(t => t.template_code === code);

        templates.push({
          code,
          hasCurrent: versions.includes('current.docx'),
          versions,
          meta: dbRecord ? {
            name: dbRecord.template_name,
            description: dbRecord.description,
            fields: dbRecord.available_fields ? JSON.parse(dbRecord.available_fields) : [],
            is_active: dbRecord.is_active
          } : null
        });
      }
    }

    res.json(templates);
  } catch (err) {
    console.error('Failed to list templates:', err);
    // Fallback to old method if database fails
    const list = await listTemplates();
    res.json(list);
  }
});


// Get meta/fields for a template
app.get("/demand-letters-apixx/templates/:code/meta", async (req, res) => {
  try {
    const code = safeCode(req.params.code);
    const meta = await readMeta(code);
    if (!meta) return res.status(404).json({ error: "No meta.json for template" });
    res.json(meta);
  } catch (e) {
    res.status(400).json({ error: e.message || String(e) });
  }
});

// Get meta/fields for a template
app.get("/demand-letters-api/templates/:code/meta", async (req, res) => {
  try {
    const code = safeCode(req.params.code);

    const pool = await getSqlPool();  // ← Use your existing function
    const result = await pool.request()
      .input('code', sql.NVarChar(100), code)
      .query(`
                SELECT 
                    template_name,
                    description,
                    available_fields,
                    is_active
                FROM dbo.demand_letter_template
                WHERE template_code = @code
            `);

    if (!result.recordset || result.recordset.length === 0) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const record = result.recordset[0];
    res.json({
      name: record.template_name,
      description: record.description,
      fields: record.available_fields ? JSON.parse(record.available_fields) : [],
      is_active: record.is_active
    });
  } catch (e) {
    console.error('Failed to read template metadata:', e);
    res.status(500).json({ error: 'Failed to read template metadata' });
  }
});

// Upload or version a template
// form-data: code=<string>, version=<optional string>, file=<.docx>, meta=<optional json as text>
app.post("/demand-letters-api/templates", upload.fields([{ name: "file" }, { name: "meta" }]), async (req, res) => {
  try {
    const code = safeCode(req.body.code);
    if (!code) throw new Error("Missing template code");
    const version = safeCode(req.body.version) || "current";
    const file = (req.files?.file || [])[0];
    if (!file) throw new Error("Missing file");
    if (!/\.docx$/i.test(file.originalname)) throw new Error("Only .docx files allowed");

    const dir = path.join(TEMPLATES_DIR, code);
    await fs.mkdir(dir, { recursive: true });
    const outPath = path.join(dir, `${version}.docx`);
    await fs.writeFile(outPath, file.buffer);

    // Optional meta.json
    const metaField = (req.body?.meta || "").toString().trim();
    if (metaField) {
      let parsed;
      try { parsed = JSON.parse(metaField); } catch { throw new Error("Invalid meta JSON"); }
      await fs.writeFile(path.join(dir, "meta.json"), JSON.stringify(parsed, null, 2));
    }

    // Invalidate cache
    cache.delete(`${code}:${version}`);
    if (version !== "current" && !(await exists(path.join(dir, "current.docx")))) {
      // if first upload with a version, also set current if missing
      await fs.copyFile(outPath, path.join(dir, "current.docx"));
    }

    res.json({ ok: true, code, version });
  } catch (e) {
    res.status(400).json({ error: e.message || String(e) });
  }
});

// Switch current to a specific version
app.put("/demand-letters-api/templates/:code/current", async (req, res) => {
  try {
    const code = safeCode(req.params.code);
    const version = safeCode(req.body.version);
    if (!version) throw new Error("Missing version");
    const dir = path.join(TEMPLATES_DIR, code);
    const src = path.join(dir, `${version}.docx`);
    const dst = path.join(dir, "current.docx");
    if (!(await exists(src))) throw new Error("Source version not found");
    await fs.copyFile(src, dst);
    cache.delete(`${code}:current`);
    res.json({ ok: true, code, current: version });
  } catch (e) {
    res.status(400).json({ error: e.message || String(e) });
  }
});

app.put("/demand-letters-api/templates/:code/meta", async (req, res) => {
  try {
    const code = safeCode(req.params.code);
    if (!code) throw new Error("Missing template code");

    const { name, description, fields } = req.body;
    const user = req.user?.preferred_username || req.user?.email || req.headers['x-user'] || 'system';

    const pool = await getSqlPool();  // ← Use your existing function
    const result = await pool.request()
      .input('code', sql.NVarChar(100), code)
      .input('name', sql.NVarChar(200), name)
      .input('description', sql.NVarChar(500), description || null)
      .input('fields', sql.NVarChar(sql.MAX), fields ? JSON.stringify(fields) : null)
      .input('user', sql.NVarChar(128), user)
      .query(`
                UPDATE dbo.demand_letter_template
                SET 
                    template_name = @name,
                    description = @description,
                    available_fields = @fields,
                    updated_by = @user,
                    updated_at = GETDATE()
                OUTPUT 
                    INSERTED.template_name,
                    INSERTED.description,
                    INSERTED.available_fields,
                    INSERTED.is_active
                WHERE template_code = @code
            `);

    if (!result.recordset || result.recordset.length === 0) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const updated = result.recordset[0];
    res.json({
      ok: true,
      code,
      meta: {
        name: updated.template_name,
        description: updated.description,
        fields: updated.available_fields ? JSON.parse(updated.available_fields) : [],
        is_active: updated.is_active
      }
    });
  } catch (e) {
    console.error('Failed to update template metadata:', e);
    res.status(400).json({ error: e.message || String(e) });
  }
});

// PATCH /demand-letters-api/templates/:code/status
// ENABLE/DISABLE: Toggle template active status
app.patch("/demand-letters-api/templates/:code/status", async (req, res) => {
  try {
    const code = safeCode(req.params.code);
    if (!code) throw new Error("Missing template code");

    const { is_active } = req.body;
    if (typeof is_active !== 'boolean') {
      return res.status(400).json({ error: "is_active must be a boolean" });
    }

    const user = req.user?.preferred_username || req.user?.email || req.headers['x-user'] || 'system';

    const pool = await getSqlPool();  // ← Use your existing function
    const result = await pool.request()
      .input('code', sql.NVarChar(100), code)
      .input('is_active', sql.Bit, is_active ? 1 : 0)
      .input('user', sql.NVarChar(128), user)
      .query(`
                UPDATE dbo.demand_letter_template
                SET 
                    is_active = @is_active,
                    updated_by = @user,
                    updated_at = GETDATE()
                WHERE template_code = @code
            `);

    if (result.rowsAffected[0] === 0) {
      return res.status(404).json({ error: 'Template not found' });
    }

    res.json({
      ok: true,
      code,
      is_active
    });
  } catch (e) {
    console.error('Failed to update template status:', e);
    res.status(400).json({ error: e.message || String(e) });
  }
});



// Generate (DOCX/PDF) from a specific template code (+optional version)
app.post("/demand-letters-api/letters", authenticate, async (req, res, next) => {

  try {
    const {
      template_code = "DL1_KB",
      template_version = null,
      format = "docx",
      sendoption = 'PREVIEW',
      data = {},
      provider_ref = null
    } = req.body || {};

    if (!data.our_ref) {
      data.our_ref = await generateOurRef({
        template_code,
        account_number: data?.customer?.account_number,
        customer_number: data?.customer?.customer_number
      });
    }

    if (String(template_code).trim().toUpperCase() === "DL1") {
      if (data.date) data.date = formatFullDateMixedCase(data.date);
      if (data.as_of_date) data.as_of_date = formatFullDateMixedCase(data.as_of_date);

      if (data.loan?.total_customer_balance !== undefined) {
        data.loan.total_customer_balance = formatAbsMoney(data.loan.total_customer_balance);
      }

      if (Array.isArray(data.accounts)) {
        data.accounts = data.accounts.map((acc) => ({
          ...acc,
          outstanding_balance: formatAbsMoney(acc.outstanding_balance),
          arrears_amount: formatAbsMoney(acc.arrears_amount),
          classification: classifyByDaysPastDue(acc.arrears_days) || acc.classification,
        }));
      }
    }

    // Resolve & render
    const p = await resolveTemplatePath(template_code, template_version);
    const docxBuffer = await renderDocxFromTemplate(p, data);
    const isPdf = String(format).toLowerCase() === "pdf";

    // Build filename like:  <account>_<template>_<YYYYMMDDHHmmss>.docx/pdf
    const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
    const template = (template_code || "demand").replace(/[^\w.-]+/g, "_");
    const timestamp = dayjs().format("YYYYMMDD_HHmmss");
    const ext = isPdf ? "pdf" : "docx";
    // Convert once if needed
    const blob = isPdf ? await docxToPdfBuffer(docxBuffer) : docxBuffer;
    const contentType = isPdf
      ? "application/pdf"
      : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    // Who sent (from Keycloak/req header/user claim)
    const sent_by = (req.user?.username || req.user?.email || req.headers['x-user'] || 'unknown');

    // Filename for non-persist responses
    const baseName = `${account}_${template}_${timestamp}.${ext}`;

    // Common headers
    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition, Content-Type, Content-Length");
    res.setHeader("Content-Type", contentType);

    // Save to MinIO + insert history if PRINT
    if (sendoption === 'PRINT') {
      const saved = await saveLetterToMinioAndLog({
        template_code,
        data,
        blob,
        ext,
        contentType,
        sent_by,
        provider_ref,
        our_ref: data.our_ref,
        status: "SAVED"
      });

      // Return the actual binary with the stored name so the user downloads what we logged
      res.setHeader("Content-Disposition", `attachment; filename="${saved.document_name}"`);
      return res.send(blob);

    }

    // Default: just stream back (no save/log)
    res.setHeader("Content-Disposition", `attachment; filename="${baseName}"`);
    return res.send(blob);

  } catch (err) {
    console.log('LETTERS_ROUTE_ERROR', {
      message: err?.message,
      name: err?.name,
      code: err?.code,
      stack: err?.stack,
    });
    //res.status(400).json({ error: err?.message || String(err) });
    req.log.error('account info error', {
      error: err?.message,
      name: err?.name,
      code: err?.code,
      stack: err?.stack,
    });
    next(err);
  }
});


// POST /demand-letters-api/repossession-orders
// Generates a repossession instruction from templates/REPO/current.docx.
// Body: { template_version?, format: "docx"|"pdf", sendoption: "PREVIEW"|"PRINT", data: {...} }
app.post("/demand-letters-api/repossession-orders", authenticate, async (req, res, next) => {
  try {
    const {
      template_version = null,
      format = "pdf",
      sendoption = "PREVIEW",
      data = {},
      provider_ref = null,
    } = req.body || {};

    const normalizedFormat = String(format).trim().toLowerCase();
    if (!["docx", "pdf"].includes(normalizedFormat)) {
      return res.status(400).json({ error: "format must be either 'docx' or 'pdf'" });
    }

    const normalizedSendOption = String(sendoption).trim().toUpperCase();
    if (!["PREVIEW", "PRINT"].includes(normalizedSendOption)) {
      return res.status(400).json({ error: "sendoption must be either 'PREVIEW' or 'PRINT'" });
    }

    const model = buildRepossessionModel(data);
    validateRepossessionModel(model);

    if (!model.our_ref) {
      model.our_ref = await generateRepossessionOurRef();
    }
    if (data.date) {
      model.date = formatRepossessionDate(data.date);
    }

    const templateCode = "REPO";
    const templatePath = await resolveTemplatePath(templateCode, template_version);
    const docxBuffer = await renderDocxFromTemplate(templatePath, model);
    const isPdf = normalizedFormat === "pdf";
    const blob = isPdf ? await docxToPdfBuffer(docxBuffer) : docxBuffer;
    const ext = isPdf ? "pdf" : "docx";
    const contentType = isPdf
      ? "application/pdf"
      : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    const account = String(model.customer.account_number).replace(/[^\w.-]+/g, "_");
    const timestamp = dayjs().format("YYYYMMDD_HHmmss");
    const baseName = `${account}_REPO_${timestamp}.${ext}`;

    res.setHeader("Access-Control-Expose-Headers", "Content-Disposition, Content-Type, Content-Length, X-Our-Ref");
    res.setHeader("Content-Type", contentType);
    res.setHeader("X-Our-Ref", model.our_ref);

    if (normalizedSendOption === "PRINT") {
      const saved = await saveLetterToMinioAndLog({
        template_code: templateCode,
        data: model,
        blob,
        ext,
        contentType,
        sent_by: actorFromRequest(req),
        provider_ref,
        our_ref: model.our_ref,
        status: "SAVED",
      });

      res.setHeader("Content-Disposition", `attachment; filename="${saved.document_name}"`);
      return res.send(blob);
    }

    res.setHeader("Content-Disposition", `attachment; filename="${baseName}"`);
    return res.send(blob);
  } catch (err) {
    req.log?.error?.("repossession order generation failed", {
      error: err?.message,
      name: err?.name,
      code: err?.code,
      stack: err?.stack,
    });
    return next(err);
  }
});

// POST /letters/preview
// Body: { template_code, template_version?, data, kind: "pdf"|"png", page?, dpi? }
app.post("/demand-letters-api/letters/preview", async (req, res) => {
  try {
    const {
      template_code = "DL1",
      template_version = null,
      data = {},
      kind = "png",           // default png preview
      page = 1,
      dpi = 144
    } = req.body || {};

    const p = await resolveTemplatePath(template_code, template_version);
    const docx = await renderDocxFromTemplate(p, data);
    const pdf = await docxToPdfBuffer(docx);

    if (String(kind).toLowerCase() === "pdf") {
      const b64 = Buffer.from(pdf).toString("base64");
      return res.json({ kind: "pdf", base64: b64, contentType: "application/pdf" });
    }

    // default: PNG (first page unless specified)
    const png = await pdfToPngBuffer(pdf, { page: Number(page) || 1, dpi: Number(dpi) || 144 });
    const b64 = Buffer.from(png).toString("base64");
    res.json({ kind: "png", page: Number(page) || 1, dpi: Number(dpi) || 144, base64: b64, contentType: "image/png" });
  } catch (err) {
    console.log(err)
    res.status(400).json({ error: err?.message || String(err) });
  }
});

function maskAccountNumber(accountNumber) {
  if (!accountNumber) return '';

  // Convert to string just in case
  const str = String(accountNumber).trim();

  if (str.length <= 3) return str;

  // Keep first 3 characters and mask the rest with *
  const visible = str.slice(0, 3);
  const hidden = '*'.repeat(str.length - 3);

  return visible + hidden;
}

// POST /demand-letters-api/letters/email
// Body: { template_code, template_version?, data, to, cc?, bcc?, subject?, body? }
app.post("/demand-letters-api/letters/email", async (req, res) => {
  const traceId = req.headers['x-request-id'] || req.id || crypto.randomUUID();
  const logStep = (step, extra = {}) => {
    console.log(`[letters/email][${traceId}] ${step}`, extra);
  };

  try {
    logStep('STEP 1 - request accepted');

    const {
      template_code = "DL1",
      template_version = null,
      data = {},
      to,
      cc,
      bcc,
      subject,
      body,
    } = req.body || {};

    logStep('STEP 2 - payload parsed', {
      template_code,
      template_version,
      to,
      has_cc: !!cc,
      has_bcc: !!bcc,
      has_subject: !!subject,
      account_number: data?.customer?.account_number || null,
      customer_number: data?.customer?.customer_number || null,
      provider_ref: req.body?.provider_ref || null,
    });

    // ⬇️ NEW: our_ref if absent
    if (!data.our_ref) {
      logStep('STEP 3 - generating our_ref');
      data.our_ref = await generateOurRef({ template_code, account_number: data?.customer?.account_number });
      logStep('STEP 4 - our_ref generated', { our_ref: data.our_ref });
    } else {
      logStep('STEP 4 - existing our_ref detected', { our_ref: data.our_ref });
    }

    if (!to) {
      logStep('STEP 5 - validation failed', { error: "Missing 'to' email address" });
      return res.status(400).json({ error: "Missing 'to' email address" });
    }

    // Render DOCX -> PDF
    logStep('STEP 6 - resolving template path');
    const p = await resolveTemplatePath(template_code, template_version);
    logStep('STEP 7 - template path resolved', { template_path: p });

    logStep('STEP 8 - rendering DOCX');
    const docxBuffer = await renderDocxFromTemplate(p, data);
    logStep('STEP 9 - DOCX rendered', { docx_size: docxBuffer?.length || 0 });

    logStep('STEP 10 - converting DOCX to PDF');
    const pdf = await docxToPdfBuffer(docxBuffer);
    logStep('STEP 11 - PDF rendered', { pdf_size: pdf?.length || 0 });

    // Build filename e.g. L0012142_demand1_YYYYMMDD_HHmmss.pdf
    const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
    const template = (template_code || "demand").replace(/[^\w.-]+/g, "_");
    const timestamp = dayjs().format("YYYYMMDD_HHmmss");
    const filename = `${account}_${template}_${timestamp}.pdf`;
    logStep('STEP 12 - filename built', { filename });

    logStep('STEP 13 - creating mail transport');
    const { transport, from } = makeMailer();
    logStep('STEP 14 - mail transport created', { from });

    // build a nice HTML version
    const htmlBody = `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      body {
        margin: 0;
        padding: 0;
        font-family: "Segoe UI", Arial, sans-serif;
        background-color: #f4f4f4;
        color: #2c2c2c;
      }
      .container {
        max-width: 640px;
        margin: 2rem auto;
        background: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
        border-top: 4px solid #6d9ad1ff;
      }
      .header {
        background-color: #e87722;
        color: #ffffff;
        padding: 1.25rem 1.75rem;
        font-size: 1.25rem;
        font-weight: 600;
        letter-spacing: 0.3px;
      }
      .content {
        padding: 1.75rem;
        line-height: 1.6;
      }
      .content p {
        margin: 0.9rem 0;
      }
      .btn {
        display: inline-block;
        padding: 0.6rem 1.25rem;
        background-color: #3e93e7ff;
        color: #5697d3ff !important;
        border-radius: 4px;
        text-decoration: none;
        font-weight: 600;
        margin-top: 1rem;
      }
      .footer {
        background-color: #fafafa;
        padding: 1rem 1.75rem;
        font-size: 0.85rem;
        color: #555555;
        border-top: 1px solid #eee;
      }
      a {
        color: #227ee8ff;
        text-decoration: none;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">Kingdom Bank Kenya – Demand Letter</div>
      <div class="content">
        <p>Dear <strong>${data?.customer?.name || "Member"}</strong>,</p>

        <p>
          We hope this message finds you well. This is a reminder that your
          loan account <strong>${maskAccountNumber(data?.customer?.account_number) || maskAccountNumber(account)}</strong> 
          is currently in arrears.
        </p>

        <p>
          Please review the attached <strong>Demand Letter</strong> for details 
          on your outstanding balance and repayment obligations.
        </p>

        <p>
          To avoid additional interest or penalties, kindly make payment or 
          contact our Recoveries Team immediately for assistance.
        </p>

        <p style="margin-top:1rem;">
          <a href="mailto:recoveries@kingdombankltd.co.ke" class="btn">Contact Recoveries</a>
        </p>

        <p>
          Thank you for being a valued member of Kingdom Bank Kenya.
          We appreciate your prompt attention to this matter.
        </p>

        <p>Warm regards,<br />
        <strong>Recoveries Department</strong><br />
        Kingdom Bank Kenya</p>
      </div>
      <div class="footer">
        <p>
          This email and any attachments are confidential and intended solely 
          for the addressed recipient. If you received this message in error, 
          please notify us immediately and delete it.
        </p>
        <p> 
          Kingdom Bank Kenya | P.O. Box 22741- 00400 Nairobi | 
          <a href="https://www.kingdombankltd.co.ke">www.kingdombankltd.co.ke</a>
        </p>
      </div>
    </div> 
  </body>
</html> 
`;


    // now send using nodemailer
    logStep('STEP 15 - sending email', {
      to,
      has_attachment: true,
      attachment_name: filename,
      attachment_size: pdf?.length || 0,
    });
    const mail = await transport.sendMail({
      from,
      to,
      cc,
      bcc,
      subject: subject || `Demand Letter - ${maskAccountNumber(data?.customer?.account_number)}`,
      text:
        body ||
        `Dear Customer,

Please find attached your demand letter for account ${data?.customer?.account_number}.

Regards,
Recoveries Team`,
      html: htmlBody,
      attachments: [
        {
          filename,
          content: pdf,
          contentType: "application/pdf",
        },
      ],
    });
    logStep('STEP 16 - email sent', {
      messageId: mail?.messageId || null,
      accepted: mail?.accepted || [],
      rejected: mail?.rejected || [],
      response: mail?.response || null,
    });

    logStep('STEP 17 - saving letter to MinIO/history');
    const saved = await saveLetterToMinioAndLog({
      template_code,
      data,
      blob: pdf,                  // you already rendered to PDF for email
      ext: 'pdf',
      contentType: 'application/pdf',
      sent_by: from,
      provider_ref: mail.messageId,
      our_ref: data.our_ref,
      status: "SENT",
    });
    logStep('STEP 18 - saved letter to MinIO/history', {
      history_id: saved?.id || null,
      object_key: saved?.key || null,
      document_name: saved?.document_name || null,
    });

    logStep('STEP 19 - sending API response');
    res.json({
      ok: true,
      messageId: mail.messageId,
      history_id: saved.id,
      document_name: saved.document_name,
      object_key: saved.key,
      idem_key: saved.idem_key,
      our_ref: data.our_ref,
      url: saved.signedUrl,
    });
    logStep('STEP 20 - API response sent');
  } catch (e) {
    console.log(`[letters/email][${traceId}] ERROR`, {
      message: e?.message,
      name: e?.name,
      code: e?.code,
      command: e?.command,
      responseCode: e?.responseCode,
      response: e?.response,
      stack: e?.stack,
    });
    res.status(400).json({
      ok: false,
      error: e?.message || String(e),
      traceId,
      code: e?.code || null,
      command: e?.command || null,
      responseCode: e?.responseCode || null,
    });
  }
});

// GET /letters/download/:id
// Look up history row by id, issue a presigned GET and redirect (302)
app.get("/demand-letters-api/letters/download/:id", async (req, res) => {
  const pool = await getSqlPool();
  const r = await pool.request()
    .input("id", sql.BigInt, Number(req.params.id))
    .query("SELECT TOP 1 bucket, object_key, document_name FROM dbo.demand_letter_history WHERE id=@id");
  const row = r.recordset?.[0];
  if (!row) return res.status(404).send("Not found");

  const url = await presignGet({ bucket: row.bucket, key: row.object_key });
  res.setHeader("Content-Disposition", `attachment; filename="${row.document_name}"`);
  res.redirect(302, url);
});

// GET /demand-letters-api/templates/:code/:version.docx
// DOWNLOAD: Download a specific template version
// Enhanced version with better error handling
app.get("/demand-letters-api/templates/:code/:version.docx", async (req, res) => {
  try {
    const code = safeCode(req.params.code);
    const version = req.params.version.replace('.docx', ''); // Remove .docx if present

    const dir = path.join(TEMPLATES_DIR, code);
    const filename = `${version}.docx`;
    const filepath = path.join(dir, filename);

    console.log('dir:::' + dir);
    console.log('filename:::' + filename);
    console.log('filepath:::' + filepath);

    // Verify file exists
    if (!(await exists(filepath))) {
      return res.status(404).json({ error: "Template file not found" });
    }

    // Send file with proper headers
    const downloadName = `${code}_${version}.docx`;
    res.setHeader("Content-Disposition", `attachment; filename="${downloadName}"`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

    const fileContent = await fs.readFile(filepath);
    res.send(fileContent);
  } catch (e) {
    res.status(400).json({ error: e.message || String(e) });
  }
});

// GET /demand-letters-api/letters/history?account=ACC123&page=0&pageSize=10
app.get("/demand-letters-api/letters/history", async (req, res) => {
  try {
    const account = (req.query.account || "").trim();
    if (!account) {
      return res.status(400).json({ error: "Missing ?account parameter" });
    }

    const page = Number(req.query.page || 0);
    const pageSize = Number(req.query.pageSize || 20);

    const pool = await getSqlPool();

    const q = `
      SELECT
        id,
        account_number,
        customer_number,
        demand_type,
        date_sent,
        days_in_arrears,
        outstanding_balance,
        sent_by,
        document_name,
        bucket,
        object_key,
        our_ref,
        provider_ref,
        status
      FROM dbo.demand_letter_history
      WHERE account_number = @account
      ORDER BY date_sent DESC
      OFFSET @offset ROWS FETCH NEXT @pageSize ROWS ONLY;
    `;

    const r = await pool.request()
      .input("account", sql.NVarChar(100), account)
      .input("offset", sql.Int, page * pageSize)
      .input("pageSize", sql.Int, pageSize)
      .query(q);

    // Return a plain array for Angular
    res.json(r.recordset || []);
  } catch (err) {
    console.error("Error fetching demand letter history:", err);
    res.status(500).json({ error: err.message || "Server error" });
  }
});

app.use((err, req, res, next) => {
  const requestId =
    err.requestId ||
    req.requestId ||
    req.headers['x-request-id']

  res.header('Access-Control-Expose-Headers', 'X-Request-Id');
  res.setHeader('x-request-id', requestId);


  return res.status(err.status || 500).json({
    ok: false,
    error: err.message || 'Internal server error',
    requestId,
  });
});

const PORT = process.env.PORT || 8004;
app.listen(PORT, () => console.log(`Demand Letter API listening on :${PORT}`));
