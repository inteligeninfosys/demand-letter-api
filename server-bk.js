// server.js
import 'dotenv/config';
import express from "express";
import fs from "fs/promises";
import { spawn } from "child_process";
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
import { authenticate } from './auth.js';
import { minioClient } from './middleware/minioClient.js';

dayjs.extend(utc);
const require = createRequire(import.meta.url);


import { S3Client, PutObjectCommand, GetObjectCommand } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";


const __dirname = path.dirname(fileURLToPath(import.meta.url));


const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: "4mb" }));

const TEMPLATES_DIR = path.join(__dirname, "templates");
const upload = multer({ storage: multer.memoryStorage() });

/* ---------- Helpers ---------- */

const execFileAsync = promisify(execFile);
const safeCode = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9_\-]/g, "");
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

    // 1) sanitize inputs (trim spaces, normalize)
    const code = String(template_code).trim();                // e.g. 'DL1'
    const verIn = (template_version ?? "current").toString().trim();

    // If caller passed a full filename like "current.docx" or "DL1_v2.docx", use it as-is
    const isFileName = /\.(docx)$/i.test(verIn);
    const fileName = isFileName ? verIn : `${verIn}.docx`;    // "current" -> "current.docx"

    // 2) candidate roots (absolute!)
    const roots = [
        process.env.TEMPLATE_DIR,                          // prefer explicit env
        path.join(__dirname, "templates"),                // ./templates next to server.js
        "/app/templates",                                 // default in your image
        "/data/templates",                                // optional external mount
    ].filter(Boolean);

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
        //nullGetter: () => "",                      // return empty string for missing values
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
    const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
    const idem_key = generateIdemKey(template_code, account);

    const tmpl = (template_code || "demand").replace(/[^\w.-]+/g, "_");
    const ts = dayjs().format("YYYY/MM/DD");
    const tsName = dayjs().format("YYYYMMDD_HHmmss");
    const document_name = `${account}_${tmpl}_${tsName}.${ext}`;
    const object_key = `letters/${tmpl}/${ts}/${document_name}`;

    // 1) Upload to MinIO
    const { bucket, key } = await uploadToS3({ key: object_key, body: blob, contentType });

    // 2) (Optional) Pre-sign a GET URL for quick UI open
    const signedUrl = await presignGet({ bucket, key, expiresInSec: process.env.S3_SIGN_URL_EXP_SECONDS });
    const signedUrlExpiryUtc = dayjs().add(Number(process.env.S3_SIGN_URL_EXP_SECONDS || 3600), "second")
        .toDate();

    // 3) Insert history row
    const id = await insertHistory({
        account_number: data?.customer?.account_number || null,
        customer_number: data?.customer?.customer_number || null,
        demand_type: template_code,
        date_sent: new Date(),
        days_in_arrears: data?.loan?.days_in_arrears ?? null,
        outstanding_balance: (data?.loan?.outstanding_balance ?? null), // numeric if you have it
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
        .input("outstanding_balance", sql.Decimal(18, 2), outstanding_balance ?? null)
        .input("arrears_amount", sql.Decimal(18, 2), arrears_amount ?? null)
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
async function generateOurRef({ template_code, account_number }) {
    const prefix = (process.env.OUR_REF_PREFIX || "STIMA/REC").trim();
    const tmpl = (template_code || "DEMAND").toUpperCase().replace(/[^\w/-]+/g, "");
    const yyyy = dayjs().utc().format("YYYY");

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
        return `${prefix}/${tmpl}/${yyyy}/${ts}-${rand}`; // e.g. STIMA/REC/DEMAND1/2025/20251107...-ABCD
    }

    return `${prefix}/${tmpl}/${yyyy}/${seq}`; // e.g. STIMA/REC/DEMAND1/2025/100321
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
    const list = await listTemplates();
    res.json(list);
});

// Get meta/fields for a template
app.get("/demand-letters-api/templates/:code/meta", async (req, res) => {
    try {
        const code = safeCode(req.params.code);
        const meta = await readMeta(code);
        if (!meta) return res.status(404).json({ error: "No meta.json for template" });
        res.json(meta);
    } catch (e) {
        res.status(400).json({ error: e.message || String(e) });
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

// Generate (DOCX/PDF) from a specific template code (+optional version)
app.post("/demand-letters-api/letters", authenticate, async (req, res) => {

    try {
        const {
            template_code = "DL1",
            template_version = null,
            format = "docx",
            sendoption = 'PREVIEW',
            data = {},
            provider_ref = null
        } = req.body || {};

        if (!data.our_ref) {
            data.our_ref = await generateOurRef({
                template_code,
                account_number: data?.customer?.account_number
            });
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
        console.log(err)
        res.status(400).json({ error: err?.message || String(err) });
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
app.post("/demand-letters-api/letters/emailxx", async (req, res) => {
    try {
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

        // ⬇️ NEW: our_ref if absent
        if (!data.our_ref) {
            data.our_ref = await generateOurRef({ template_code, account_number: data?.customer?.account_number });
        }

        if (!to) return res.status(400).json({ error: "Missing 'to' email address" });

        // Render DOCX -> PDF
        const p = await resolveTemplatePath(template_code, template_version);
        const docxBuffer = await renderDocxFromTemplate(p, data);
        const pdf = await docxToPdfBuffer(docxBuffer);

        // Build filename e.g. L0012142_demand1_YYYYMMDD_HHmmss.pdf
        const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
        const template = (template_code || "demand").replace(/[^\w.-]+/g, "_");
        const timestamp = dayjs().format("YYYYMMDD_HHmmss");
        const filename = `${account}_${template}_${timestamp}.pdf`;

        const { transport, from } = makeMailer();

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
        border-top: 4px solid #ffc20e;
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
        background-color: #e87722;
        color: #ffffff !important;
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
        color: #e87722;
        text-decoration: none;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">Stima-Sacco – Demand Letter</div>
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
          <a href="mailto:recoveries@stima-sacco.com" class="btn">Contact Recoveries</a>
        </p>

        <p>
          Thank you for being a valued member of Stima Sacco Kenya.
          We appreciate your prompt attention to this matter.
        </p>

        <p>Warm regards,<br />
        <strong>Recoveries Department</strong><br />
        Stima Sacco Kenya</p>
      </div>
      <div class="footer">
        <p>
          This email and any attachments are confidential and intended solely 
          for the addressed recipient. If you received this message in error, 
          please notify us immediately and delete it.
        </p>
        <p>
          Stima Sacco Kenya | P.O. Box 75629–00200 Nairobi | 
          <a href="https://www.stima-sacco.com">www.stima-sacco.com</a>
        </p>
      </div>
    </div>
  </body>
</html>
`;


        // now send using nodemailer
        const mail = await transport.sendMail({
            from,
            to,
            cc,
            bcc,
            subject: `Demand Letter - ${maskAccountNumber(data?.customer?.account_number)}`,
            text:
                body ||
                `Dear Customer,\n\nPlease find attached your demand letter for account ${data?.customer?.account_number}.\n\nRegards,\nRecoveries Team`,
            html: htmlBody,
            attachments: [
                {
                    filename,
                    content: pdf,
                    contentType: "application/pdf",
                },
            ],
        });

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
    } catch (e) {
        console.log(e)
        res.status(400).json({ error: e?.message || String(e) });
    }
});

app.post("/demand-letters-api/letters/email", async (req, res) => {
    try {
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

        // ⬇️ NEW: our_ref if absent
        if (!data.our_ref) {
            data.our_ref = await generateOurRef({ template_code, account_number: data?.customer?.account_number });
        }

        if (!to) return res.status(400).json({ error: "Missing 'to' email address" });

        // Render DOCX -> PDF
        const p = await resolveTemplatePath(template_code, template_version);
        const docxBuffer = await renderDocxFromTemplate(p, data);
        const pdf = await docxToPdfBuffer(docxBuffer);

// Build filename e.g. L0012142_demand1_YYYYMMDD_HHmmss.pdf
const account = (data?.customer?.account_number || "unknown").replace(/[^\w.-]+/g, "_");
const template = (template_code || "demand").replace(/[^\w.-]+/g, "_");
const timestamp = dayjs().format("YYYYMMDD_HHmmss");
const filename = `${account}_${template}_${timestamp}.pdf`;

// Mask the account number for email subject (show first and last character only)
const maskAccountNumber = (acc) => {
  if (!acc || acc.length < 2) return acc;
  const str = String(acc).trim();
  if (str.length < 2) return str;
  const firstChar = str.charAt(0);
  const lastChar = str.charAt(str.length - 1);
  const masked = 'X'.repeat(Math.max(0, str.length - 2));
  return `${firstChar}${masked}${lastChar}`;
};

const maskedAccount = maskAccountNumber(data?.customer?.account_number || account);

const { transport, from } = makeMailer();

// Read logo file for embedding
const logoPath = path.join(__dirname, "assets", "images", "auth.jpg");
let logoBuffer;
try {
    logoBuffer = await fs.readFile(logoPath);
} catch (err) {
    console.warn("Logo file not found at", logoPath, "- email will be sent without logo");
}

// Build professional email with embedded logo
const htmlBody = `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      * { 
        box-sizing: border-box; 
        margin: 0;
        padding: 0;
      }
      body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        background-color: #f5f7fa;
        line-height: 1.6;
      }
      .email-wrapper {
        width: 100%;
        background-color: #f5f7fa;
        padding: 20px 10px;
      }
      .container {
        max-width: 600px;
        margin: 0 auto;
        background: #ffffff;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-radius: 8px;
      }
      
      /* HEADER with Logo */
      .header {
        background: #ffffff;
        padding: 30px 30px 20px;
        text-align: center;
        border-bottom: 1px solid #e8edf2;
      }
      .logo-container {
        margin: 0 auto 10px;
        max-width: 320px;
      }
      .logo-container img {
        max-width: 100%;
        height: auto;
        display: block;
        margin: 0 auto;
      }
      
      /* CONTENT */
      .content {
        padding: 35px 30px;
        background: #ffffff;
      }
      .content p {
        margin: 0 0 18px 0;
        color: #4a5568;
        font-size: 15px;
        line-height: 1.7;
      }
      .greeting {
        font-size: 16px;
        color: #2d3748;
        font-weight: 500;
      }
      
      /* ACCOUNT BOX */
      .account-box {
        background: linear-gradient(135deg, #f0f4ff 0%, #e8f0fe 100%);
        border-left: 4px solid #3949AB;
        padding: 20px 24px;
        margin: 26px 0;
        border-radius: 6px;
      }
      .account-box p {
        margin: 0;
        color: #2d3748;
      }
      .account-label {
        font-size: 12px;
        color: #64748b;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 8px;
      }
      .account-number {
        font-family: "SF Mono", "Monaco", "Consolas", "Courier New", monospace;
        font-size: 20px;
        font-weight: 700;
        color: #3949AB;
        letter-spacing: 1.5px;
      }
      
      /* ALERT BOX */
      .alert-box {
        background: #fff3e0;
        border-left: 4px solid #FFA726;
        padding: 16px 20px;
        margin: 22px 0;
        border-radius: 4px;
        font-size: 14px;
        color: #e65100;
      }
      
      /* BUTTON */
      .btn-wrapper {
        text-align: center;
        margin: 30px 0;
      }
      .btn {
        display: inline-block;
        padding: 14px 32px;
        background: #E87722;
        color: #ffffff !important;
        border-radius: 6px;
        text-decoration: none;
        font-weight: 600;
        font-size: 15px;
        box-shadow: 0 4px 12px rgba(232, 119, 34, 0.25);
        transition: all 0.2s;
      }
      .btn:hover {
        background: #d66a1a;
        box-shadow: 0 6px 16px rgba(232, 119, 34, 0.35);
      }
      
      /* DIVIDER */
      .divider {
        height: 1px;
        background: linear-gradient(to right, transparent, #e2e8f0, transparent);
        margin: 30px 0;
      }
      
      /* SIGNATURE */
      .signature {
        margin-top: 30px;
        padding-top: 20px;
        border-top: 2px solid #f1f5f9;
      }
      .signature p {
        margin: 6px 0;
      }
      .signature-intro {
        color: #64748b;
        font-size: 15px;
        margin-bottom: 10px;
      }
      .dept-name {
        color: #3949AB;
        font-weight: 700;
        font-size: 16px;
      }
      .company-name {
        color: #64748b;
        font-size: 14px;
        font-style: italic;
      }
      
      /* FOOTER */
      .footer {
        background: #f8fafc;
        padding: 28px 30px;
        border-top: 3px solid #FFC107;
        font-size: 13px;
        color: #64748b;
      }
      .footer-heading {
        font-weight: 700;
        color: #3949AB;
        margin-bottom: 12px;
        font-size: 15px;
      }
      .footer p {
        margin: 8px 0;
        line-height: 1.6;
      }
      .footer a {
        color: #3949AB;
        text-decoration: none;
        font-weight: 500;
      }
      .footer a:hover {
        text-decoration: underline;
      }
      .footer-legal {
        font-size: 11px;
        margin-top: 20px;
        padding-top: 20px;
        border-top: 1px solid #e2e8f0;
        color: #94a3b8;
        line-height: 1.5;
      }
      
      /* RESPONSIVE */
      @media only screen and (max-width: 600px) {
        .email-wrapper {
          padding: 10px 5px;
        }
        .content {
          padding: 25px 20px;
        }
        .header {
          padding: 25px 20px 15px;
        }
        .logo-container {
          max-width: 280px;
        }
        .account-number {
          font-size: 17px;
        }
        .btn {
          display: block;
          padding: 12px 24px;
        }
        .footer {
          padding: 22px 20px;
        }
      }
    </style>
  </head>
  <body>
    <div class="email-wrapper">
      <div class="container">
        
        <!-- HEADER with Logo -->
        <div class="header">
          <div class="logo-container">
            ${logoBuffer ? '<img src="cid:stimalogo" alt="Stima Sacco" />' : '<div style="font-size: 28px; font-weight: 700; color: #3949AB; letter-spacing: 2px;">STIMA SACCO</div><div style="font-size: 13px; font-style: italic; color: #64748b; margin-top: 8px;">towards a prosperous future together</div>'}
          </div>
        </div>
        
        <!-- CONTENT -->
        <div class="content">
          <p class="greeting">Dear <strong>${data?.customer?.name || "Valued Member"}</strong>,</p>

          <p>
            We hope this message finds you well. This is an important reminder regarding 
            your loan account with Stima Sacco Kenya.
          </p>

          <div class="account-box">
            <div class="account-label">Your Account Number</div>
            <div class="account-number">${maskedAccount}</div>
          </div>

          <div class="alert-box">
            <strong>⚠️ Action Required:</strong> Your loan account is currently in arrears. 
            Immediate attention is needed to avoid additional penalties.
          </div>

          <p>
            Please review the attached <strong>Demand Letter</strong> carefully. It contains 
            complete details about your outstanding balance, payment obligations, and the 
            steps required to regularize your account.
          </p>

          <p>
            We understand that financial challenges can arise. To avoid additional interest 
            charges or legal action, we encourage you to:
          </p>

          <p style="padding-left: 20px; color: #2d3748;">
            • Make payment immediately, or<br>
            • Contact our Recoveries Team to discuss flexible repayment arrangements
          </p>

          <div class="btn-wrapper">
            <a href="mailto:recoveries@stima-sacco.com" class="btn">📧 Contact Recoveries Team</a>
          </div>

          <div class="divider"></div>

          <p>
            At Stima Sacco, we value your membership and are committed to working with you 
            towards financial stability. We appreciate your prompt attention to this matter.
          </p>

          <div class="signature">
            <p class="signature-intro"><strong>Warm regards,</strong></p>
            <p class="dept-name">Recoveries Department</p>
            <p class="company-name">Stima Sacco Society Kenya Limited</p>
          </div>
        </div>
        
        <!-- FOOTER -->
        <div class="footer">
          <p class="footer-heading">Stima Sacco Society Kenya Limited</p>
          <p>
            <strong>Head Office:</strong> Stima Sacco Plaza, Kolobot Road, Off Red Hill Road<br>
            P.O. Box 75629–00200, Nairobi, Kenya
          </p>
          <p>
            <strong>Website:</strong> <a href="https://www.stima-sacco.com">www.stima-sacco.com</a><br>
            <strong>Email:</strong> <a href="mailto:info@stima-sacco.com">info@stima-sacco.com</a><br>
            <strong>Recoveries:</strong> <a href="mailto:recoveries@stima-sacco.com">recoveries@stima-sacco.com</a>
          </p>
          <p class="footer-legal">
            <strong>CONFIDENTIALITY NOTICE:</strong> This email and any attachments are confidential 
            and intended solely for the use of the individual or entity to whom they are addressed. 
            If you have received this message in error, please notify the sender immediately and 
            delete this message from your system. Unauthorized disclosure, copying, or distribution 
            of this email is strictly prohibited.
          </p>
        </div>
        
      </div>
    </div>
  </body>
</html>
`;

// Build attachments array
const attachments = [
    {
        filename,
        content: pdf,
        contentType: "application/pdf",
    }
];

// Add logo as embedded image if available
if (logoBuffer) {
    attachments.push({
        filename: "stima-logo.jpg",
        content: logoBuffer,
        contentType: "image/jpeg",
        cid: "stimalogo"
    });
}

// Send email with embedded logo
const mail = await transport.sendMail({
    from,
    to,
    cc,
    bcc,
    subject: `Loan Account Notice - ${maskedAccount}`,
    text: body ||
          `Dear ${data?.customer?.name || "Valued Member"},\n\nYour loan account (${maskedAccount}) is currently in arrears.\n\nPlease find attached your demand letter with complete details on your outstanding balance and repayment obligations.\n\nTo avoid additional interest or penalties, kindly make payment or contact our Recoveries Team at recoveries@stima-sacco.com\n\nWarm regards,\nRecoveries Department\nStima Sacco Kenya`,
    html: htmlBody,
    attachments,
});
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
    } catch (e) {
        console.log(e)
        res.status(400).json({ error: e?.message || String(e) });
    }
});

// GET /letters/download/:id
// Look up history row by id, issue a presigned GET and redirect (302)
app.get("/demand-letters-apixx/letters/download/:id", async (req, res) => {
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

app.get("/demand-letters-api/letters/download/:id", async (req, res) => {
  const pool = await getSqlPool();

  const id = Number(req.params.id);
  if (!Number.isFinite(id)) return res.status(400).send("Invalid id");

  const r = await pool.request()
    .input("id", sql.BigInt, id)
    .query("SELECT TOP 1 bucket, object_key, document_name FROM dbo.demand_letter_history WHERE id=@id");

  const row = r.recordset?.[0];
  if (!row) return res.status(404).send("Not found");

  try {
    // If you use minio-js client:
    // const stream = await minioClient.getObject(row.bucket, row.object_key);

    // If you wrap minio in helper, adapt accordingly:
    const stream = await minioClient.getObject(row.bucket, row.object_key);

    const filename = (row.document_name || row.object_key?.split("/").pop() || `letter_${id}`).toString();

    res.setHeader("Content-Type", row.content_type || "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader("Cache-Control", "private, max-age=0, no-cache");

    stream.on("error", (e) => {
      // MinIO returns errors on stream for missing key, perms, etc.
      console.error("minio stream error", e);
      if (!res.headersSent) res.status(500).send("Download failed");
      else res.end();
    });

    stream.pipe(res);
  } catch (e) {
    console.error("download error", e);
    return res.status(500).send("Download failed");
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


const PORT = process.env.PORT || 8004;
app.listen(PORT, () => console.log(`Demand Letter API listening on :${PORT}`));
