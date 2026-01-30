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
    const code = String(template_code).trim();                // e.g. 'F_F'
    const verIn = (template_version ?? "current").toString().trim();

    // If caller passed a full filename like "current.docx" or "F_F_v2.docx", use it as-is
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
            path.join(root, code, fileName),                // /app/templates/F_F/current.docx
            path.join(root, `${code}.docx`)                 // /app/templates/F_F.docx (fallback)
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
        delimiters: { start: "[[", end: "]]" }, // matches your templates
        nullGetter: () => "",                      // return empty string for missing values
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

    // Enhanced mapping to handle all DB field variations
    const model = sanitize({
        ...data,

        // Top-level fields
        our_ref: data?.our_ref ?? "",
        date: data?.date ?? dayjs().format("YYYY-MM-DD"),

        customer: {
            ...data?.customer,

            // Map DB field → template field (handle all variations)
            name: data?.customer?.name
                ?? data?.CLIENTNAME
                ?? data?.customer.name
                ?? "",

            account_number: (
                data?.customer?.account_number
                ?? data?.ACCNUMBER
                ?? data?.customer.account_number
                ?? ""
            ).toString().trim(),

            customer_number:
                data?.customer?.customer_number
                ?? data?.CUSTOMERNO
                ?? data?.CLIENTID
                ?? "",

            // Address fields
            address_line_1: data?.customer?.address_line_1
                ?? data?.ADDRESS_1
                ?? data?.customer?.address_1
                ?? "",

            address_line_2: data?.customer?.address_line_2
                ?? data?.ADDRESS_2
                ?? data?.customer?.address_2
                ?? "",

            town: data?.customer?.town
                ?? data?.TOWN
                ?? "",

            email: data?.customer?.email
                ?? data?.EMAIL
                ?? "",

            phone_1: data?.customer?.phone_1
                ?? data?.PHONE_1
                ?? "",

            phone_2: data?.customer?.phone_2
                ?? data?.PHONE_2
                ?? "",
        },

        loan: {
            ...data?.loan,

            outstanding_balance:
                data?.loan?.outstanding_balance
                ?? data?.OUSTBALANCE
                ?? data?.OUTSTANDING_BALANCE
                ?? data?.TOTALOVERDUEAMOUNT
                ?? "",

            original_balance:
                data?.loan?.original_balance
                ?? data?.ORIGBALANCE
                ?? "",

            days_in_arrears:
                data?.loan?.days_in_arrears
                ?? data?.DAYSINARR
                ?? data?.DAYS_IN_ARREARS
                ?? "",

            arrears_amount:
                data?.loan?.arrears_amount
                ?? data?.PRINCARREARS
                ?? data?.arrears_amount
                ?? "",

            penalty_arrears:
                data?.loan?.penalty_arrears
                ?? data?.PENALARREARS
                ?? "",

            interest_rate:
                data?.loan?.interest_rate
                ?? data?.INTRATE
                ?? "",

            installment_amount:
                data?.loan?.installment_amount
                ?? data?.INSTAMOUNT
                ?? "",

            product_code:
                data?.loan?.product_code
                ?? data?.PRODUCTCODE
                ?? "",

            maturity_date:
                data?.loan?.maturity_date
                ?? data?.MATDATE
                ?? "",
        },

        guarantors: Array.isArray(data?.guarantors)
            ? data.guarantors
            : [],
    });



    // Helpful diagnostics for common mistakes
    // Enhanced diagnostics to check all fields
    const dbg = {
        "our_ref": model?.our_ref,
        "date": model?.date,
        "customer.name": model?.customer?.name,
        "customer.account_number": model?.customer?.account_number,
        "customer.address_line_1": model?.customer?.address_line_1,
        "customer.town": model?.customer?.town,
        "loan.outstanding_balance": model?.loan?.outstanding_balance,
        "loan.days_in_arrears": model?.loan?.days_in_arrears,
    };
    console.log("[DOCX] Rendering with data model:", JSON.stringify(dbg, null, 2));
    Object.entries(dbg).forEach(([k, v]) => {
        if (v === "") console.warn(`[DOCX] ⚠️ value empty for tag: ${k}`);
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

// If the sequence exists, we'll use it; else fallback to time+random (still unique).
async function generateOurRef({ template_code, account_number }) {
    const prefix = (process.env.OUR_REF_PREFIX || "KB/REC").trim();
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
        return `${prefix}/${tmpl}/${yyyy}/${ts}-${rand}`; // e.g. KB/REC/DEMAND1/2025/20251107...-ABCD
    }

    return `${prefix}/${tmpl}/${yyyy}/${seq}`; // e.g. KB/REC/DEMAND1/2025/100321
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
app.post("/demand-letters-api/letters", authenticate, async (req, res, next) => {

    try {
        const {
            template_code = "DL_7",
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
        //res.status(400).json({ error: err?.message || String(err) });
        req.log.error('account info error', { error: err.message }, req);
        next(err);
    }
});

// POST /letters/preview
// Body: { template_code, template_version?, data, kind: "pdf"|"png", page?, dpi? }
app.post("/demand-letters-api/letters/preview", async (req, res) => {
    try {
        const {
            template_code = "F_F",
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
    try {
        const {
            template_code = "F_F",
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