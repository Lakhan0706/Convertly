const express = require("express");
const cors = require("cors");
const multer = require("multer");
const { PDFDocument, StandardFonts, rgb, degrees } = require("pdf-lib");
const XLSX = require("xlsx");
const JSZip = require("jszip");
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");
const execFileAsync = promisify(execFile);
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");
const { encryptPDF } = require("@pdfsmaller/pdf-encrypt");
const { decryptPDF } = require("@pdfsmaller/pdf-decrypt");

const app = express();
const port = Number(process.env.PORT || 8080);
const maxFileSize = 25 * 1024 * 1024;
const history = [];
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: maxFileSize, files: 10 },
});

const toolSpecs = {
  "jpg-pdf": { input: "image", minFiles: 1, maxFiles: 10 },
  "pdf-jpg": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "word-pdf": { input: "word", minFiles: 1, maxFiles: 1 },
  "excel-pdf": { input: "spreadsheet", minFiles: 1, maxFiles: 1 },
  "merge-pdf": { input: "pdf", minFiles: 2, maxFiles: 10 },
  "compress-image": { input: "image", minFiles: 1, maxFiles: 1 },
  "split-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "compress-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "rotate-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "watermark": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "page-numbers": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "protect-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "unlock-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-text": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-markdown": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "compare-pdf": { input: "pdf", minFiles: 2, maxFiles: 2 },
  "crop-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "organize-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "remove-pages": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "extract-pages": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-pdfa": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "repair-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "sign-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-forms": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "edit-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "redact-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "ocr-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "ai-summarizer": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "translate-pdf": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-word": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-excel": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "pdf-to-powerpoint": { input: "pdf", minFiles: 1, maxFiles: 1 },
  "powerpoint-pdf": { input: "powerpoint", minFiles: 1, maxFiles: 1 },
  "html-pdf": { input: "html", minFiles: 0, maxFiles: 1 },
  "scan-pdf": { input: "image", minFiles: 1, maxFiles: 10 },
};
const supportedTools = new Set(Object.keys(toolSpecs));
const nativeEngines = {
  libreoffice: { command: "libreoffice", label: "LibreOffice", tools: ["word-pdf", "excel-pdf", "powerpoint-pdf", "html-pdf"] },
  poppler: { command: "pdftoppm", label: "Poppler", tools: ["pdf-jpg", "compress-pdf", "redact-pdf", "ocr-pdf", "pdf-to-powerpoint"] },
  imagemagick: { command: "magick", label: "ImageMagick", tools: ["compress-image", "redact-pdf"] },
  tesseract: { command: "tesseract", label: "Tesseract OCR", tools: ["ocr-pdf"] },
};

function commandAvailable(command) {
  const pathEntries = String(process.env.PATH || "").split(path.delimiter);
  const extensions = process.platform === "win32" ? String(process.env.PATHEXT || ".EXE;.CMD;.BAT").split(";") : [""];
  return pathEntries.some((entry) => extensions.some((extension) => fs.existsSync(path.join(entry, command + extension))));
}

function engineStatus() {
  return Object.fromEntries(Object.entries(nativeEngines).map(([key, engine]) => [key, {
    label: engine.label,
    available: commandAvailable(engine.command),
    tools: engine.tools,
  }]));
}

function assertRequiredEngines(tool) {
  const missing = Object.values(nativeEngines)
    .filter((engine) => engine.tools.includes(tool) && !commandAvailable(engine.command))
    .map((engine) => `${engine.label} (${engine.command})`);
  if (missing.length) {
    const error = new Error(`${tool} needs ${missing.join(" and ")}. Install it and add it to PATH, then restart Convertly.`);
    error.statusCode = 503;
    error.code = "ENGINE_UNAVAILABLE";
    throw error;
  }
}

app.use(cors());
app.use(express.json({ limit: "1mb" }));
app.use(express.static(__dirname));

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, service: "convertly-api", time: new Date().toISOString(), engines: engineStatus() });
});

app.get("/api/tools", (_req, res) => {
  const limitations = Object.fromEntries(Object.entries(toolSpecs)
    .filter(([, spec]) => spec.limitation)
    .map(([tool, spec]) => [tool, spec.limitation]));
  res.json({ tools: Array.from(supportedTools), specs: toolSpecs, limitations, engines: engineStatus() });
});

app.get("/api/conversions", (_req, res) => {
  res.json({ items: history.slice(-50).reverse() });
});

app.delete("/api/conversions", (_req, res) => {
  history.length = 0;
  res.json({ ok: true });
});

app.post("/api/convert", upload.array("files", 10), async (req, res, next) => {
  try {
    const tool = String(req.body.tool || "").trim();
    const files = req.files || [];
    if (!supportedTools.has(tool)) return res.status(400).json({ error: "Unknown conversion tool." });
    const options = normalizeOptions(req.body);
    validateRequest(tool, files, options);

    const results = await convert(tool, files, options);
    const record = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      tool,
      input: files.map((file) => file.originalname),
      output: results.map((result) => result.name),
      createdAt: new Date().toISOString(),
    };
    history.push(record);
    res.json({
      ...record,
      files: results.map((result) => ({
        name: result.name,
        type: result.type,
        data: result.buffer.toString("base64"),
      })),
    });
  } catch (error) {
    next(error);
  }
});

async function convert(tool, files, options = {}) {
  options = normalizeOptions(options);
  if (!supportedTools.has(tool)) throw badRequest("Unknown conversion tool.");
  if (!Array.isArray(files)) throw badRequest("files must be an array.");
  validateRequest(tool, files, options);
  assertRequiredEngines(tool);
  if (tool === "jpg-pdf") return [await imagesToPdf(files, options)];
  if (tool === "pdf-jpg") return await pdfToImages(files[0], options);
  if (tool === "word-pdf") return [await documentToPdf(files[0])];
  if (tool === "excel-pdf") return [await spreadsheetToPdf(files[0])];
  if (tool === "merge-pdf") return [await mergePdfs(files)];
  if (tool === "compress-image") return [await compressImage(files[0], options)];
  if (tool === "split-pdf") return await splitPdf(files[0], options.pages || options.ranges);
  if (tool === "compress-pdf") return [await compressPdf(files[0], options.compression)];
  if (tool === "rotate-pdf") return [await rotatePdf(files[0], options.rotation)];
  if (tool === "watermark") return [await watermarkPdf(files[0], options.text, options.opacity)];
  if (tool === "page-numbers") return [await pageNumbers(files[0], options)];
  if (tool === "protect-pdf") return [await protectPdf(files[0], options)];
  if (tool === "unlock-pdf") return [await unlockPdf(files[0], options)];
  if (tool === "pdf-to-text") return [await pdfToTextFile(files[0])];
  if (tool === "pdf-to-markdown") return [await pdfToMarkdown(files[0])];
  if (tool === "compare-pdf") return [await comparePdfs(files)];
  if (tool === "crop-pdf") return [await cropPdf(files[0], options)];
  if (tool === "organize-pdf") return [await organizePdf(files[0], options)];
  if (tool === "remove-pages") return [await removePages(files[0], options)];
  if (tool === "extract-pages") return [await extractPages(files[0], options)];
  if (tool === "pdf-to-pdfa") return [await pdfArchive(files[0])];
  if (tool === "repair-pdf") return [await rewritePdf(files[0], "-repaired")];
  if (tool === "sign-pdf") return [await signPdf(files[0], options)];
  if (tool === "pdf-forms") return [await createForms(files[0], options.fields)];
  if (tool === "edit-pdf") return [await editPdf(files[0], options)];
  if (tool === "redact-pdf") return [await redactPdf(files[0], options)];
  if (tool === "ocr-pdf") return [await ocrPdf(files[0], options)];
  if (tool === "ai-summarizer") return [await summarizePdf(files[0])];
  if (tool === "translate-pdf") return [await translatePdf(files[0], options.target || options.targetLanguage || "en")];
  if (tool === "pdf-to-word") return [await pdfToWord(files[0])];
  if (tool === "pdf-to-excel") return [await pdfToExcel(files[0])];
  if (tool === "pdf-to-powerpoint") return [await pdfToPowerpoint(files[0])];
  if (tool === "powerpoint-pdf") return [await powerpointToPdf(files[0])];
  if (tool === "html-pdf") return [await htmlToPdf(options.html || (files[0] ? files[0].buffer.toString("utf8") : ""))];
  if (tool === "scan-pdf") return [await imagesToPdf(files)];
  throw new Error("Unknown conversion tool.");
}

function unsupported(message) {
  const error = new Error(message);
  error.statusCode = 422;
  error.code = "UNSUPPORTED_TOOL";
  return error;
}

function badRequest(message, code = "INVALID_REQUEST") {
  const error = new Error(message);
  error.statusCode = 400;
  error.code = code;
  return error;
}

function normalizeOptions(options) {
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    throw badRequest("Conversion options must be an object.");
  }
  return options;
}

function stringOption(value, name, { fallback, required = false, maxLength = 2000 } = {}) {
  if (value === undefined || value === null) {
    if (required) throw badRequest(`${name} is required.`);
    return fallback;
  }
  if (typeof value !== "string") throw badRequest(`${name} must be a string.`);
  if (!value.trim()) {
    if (required) throw badRequest(`${name} must not be empty.`);
    return fallback;
  }
  if (value.length > maxLength) throw badRequest(`${name} must be ${maxLength} characters or fewer.`);
  return value;
}

function passwordOption(value, name, required = false) {
  if (value === undefined || value === null) {
    if (required) throw badRequest(`${name} is required.`);
    return undefined;
  }
  if (typeof value !== "string") throw badRequest(`${name} must be a string.`);
  if (!value.length) {
    if (required) throw badRequest(`${name} must not be empty.`);
    return undefined;
  }
  if (value.length > 256) throw badRequest(`${name} must be 256 characters or fewer.`);
  return value;
}

function booleanOption(value, name, fallback = true) {
  if (value === undefined || value === null || value === "") return fallback;
  if (value === true || value === "true") return true;
  if (value === false || value === "false") return false;
  throw badRequest(`${name} must be true or false.`);
}

function validateRequest(tool, files, options) {
  const spec = toolSpecs[tool];
  if (files.length < spec.minFiles) {
    if (spec.minFiles !== 0) {
      throw badRequest(spec.minFiles === 2 ? "Choose at least two files." : "At least one file is required.");
    }
  }
  if (files.length > spec.maxFiles) throw badRequest(`This tool accepts at most ${spec.maxFiles} file${spec.maxFiles === 1 ? "" : "s"}.`);
  if (tool === "html-pdf" && options.html !== undefined && typeof options.html !== "string") {
    throw badRequest("html must be a string.");
  }
  if (tool === "html-pdf" && options.html && options.html.length > 5 * 1024 * 1024) {
    throw badRequest("html must be 5 MB or smaller.");
  }
  if (tool === "html-pdf" && !files.length && (!options.html || !options.html.trim())) {
    throw badRequest("Provide an HTML file or an html field.");
  }
  for (const file of files) {
    if (!file.buffer || !file.buffer.length) throw badRequest(`The file "${file.originalname || "unnamed"}" is empty.`);
    if (!matchesInput(file, spec.input)) throw badRequest(`"${file.originalname}" is not a supported ${spec.input} file.`);
  }
}

function matchesInput(file, input) {
  const name = String(file.originalname || "").toLowerCase();
  const mime = String(file.mimetype || "").toLowerCase();
  if (input === "pdf") return mime === "application/pdf" || name.endsWith(".pdf");
  if (input === "image") return ["image/jpeg", "image/jpg", "image/png"].includes(mime)
    || /\.(jpe?g|png)$/i.test(name);
  if (input === "word") return mime.includes("wordprocessingml") || /\.docx?$/i.test(name);
  if (input === "spreadsheet") return mime.includes("spreadsheet") || mime === "text/csv" || /\.(xlsx?|csv)$/i.test(name);
  if (input === "powerpoint") return mime.includes("presentation") || /\.(pptx?|ppsx?)$/i.test(name);
  if (input === "html") return mime.includes("html") || /\.html?$/i.test(name);
  return true;
}

function numberOption(value, name, { min = -Infinity, max = Infinity, fallback, integer = false } = {}) {
  if (value === undefined || value === null || value === "") return fallback;
  if (typeof value !== "number" && typeof value !== "string") throw badRequest(`${name} must be a number${integer ? " (whole number)" : ""}.`);
  if (typeof value === "string" && !value.trim()) throw badRequest(`${name} must be a number${integer ? " (whole number)" : ""}.`);
  const number = Number(value);
  if (!Number.isFinite(number) || (integer && !Number.isInteger(number)) || number < min || number > max) {
    throw badRequest(`${name} must be a number${integer ? " (whole number)" : ""} between ${min} and ${max}.`);
  }
  return number;
}

function jsonOption(value, name, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  if (typeof value === "object") return value;
  try { return JSON.parse(value); } catch (_error) { throw badRequest(`${name} must be valid JSON.`); }
}

function safeBaseName(name) {
  return String(name || "convertly-file").replace(/\.[^.]*$/, "").replace(/[<>:"/\\|?*\x00-\x1f]/g, "-").slice(0, 120) || "convertly-file";
}

async function loadPdf(file) {
  if (!file || !/pdf/i.test(file.mimetype || "") && !/\.pdf$/i.test(file.originalname || "")) {
    throw badRequest("Please provide a PDF file.");
  }
  try {
    return await PDFDocument.load(file.buffer, { ignoreEncryption: false });
  } catch (error) {
    const message = /encrypt|password/i.test(error.message || "")
      ? "This PDF is password-protected and cannot be opened by the local PDF engine."
      : "The PDF is invalid or damaged and could not be read.";
    const wrapped = badRequest(message, "INVALID_PDF");
    wrapped.cause = error;
    throw wrapped;
  }
}

async function savePdf(pdf, name) {
  return { name, type: "application/pdf", buffer: Buffer.from(await pdf.save({ useObjectStreams: true })) };
}

async function mergePdfs(files) {
  if (files.length < 2) throw new Error("Choose at least two PDF files.");
  const merged = await PDFDocument.create();
  for (const file of files) {
    const source = await loadPdf(file);
    const pages = await merged.copyPages(source, source.getPageIndices());
    pages.forEach((page) => merged.addPage(page));
  }
  return savePdf(merged, "convertly-merged.pdf");
}

async function compressImage(file, options = {}) {
  if (!matchesInput(file, "image")) throw badRequest("Only JPG and PNG files can be compressed.");
  const quality = numberOption(options.quality, "quality", { min: 1, max: 100, fallback: 72, integer: true });
  const maxWidth = numberOption(options.maxWidth, "maxWidth", { min: 64, max: 10000, fallback: 1800, integer: true });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-img-"));
  const input = path.join(dir, "input"); const output = path.join(dir, "output.jpg");
  try {
    fs.writeFileSync(input, file.buffer);
    await execFileAsync("magick", [input, "-auto-orient", "-resize", `${maxWidth}x>`, "-strip", "-quality", String(quality), output], { timeout: 60000, maxBuffer: 1024 * 1024 });
    return { name: safeBaseName(file.originalname) + "-compressed.jpg", type: "image/jpeg", buffer: fs.readFileSync(output) };
  } catch (error) { const wrapped=badRequest(`"${file.originalname}" could not be compressed.`, "IMAGE_COMPRESSION_FAILED"); wrapped.cause=error; throw wrapped; }
  finally { fs.rmSync(dir,{recursive:true,force:true}); }
}

async function pdfToImages(file, options = {}) {
  await loadPdfJs(file);
  const scale = numberOption(options.scale, "scale", { min: 0.5, max: 4, fallback: 1.5 });
  const quality = numberOption(options.quality, "quality", { min: 1, max: 100, fallback: 90, integer: true });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-render-"));
  try {
    const input=path.join(dir,"input.pdf"); fs.writeFileSync(input,file.buffer);
    const dpi=Math.round(72*scale);
    await execFileAsync("pdftoppm", ["-jpeg", "-r", String(dpi), "-jpegopt", `quality=${quality}`, input, path.join(dir,"page")], { timeout: 180000, maxBuffer: 2*1024*1024 });
    const files=fs.readdirSync(dir).filter(n=>/^page-\d+\.jpg$/i.test(n)).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0]));
    return files.map((name,i)=>({name:`${safeBaseName(file.originalname)}-page-${i+1}.jpg`,type:"image/jpeg",buffer:fs.readFileSync(path.join(dir,name))}));
  } catch(error){const wrapped=badRequest("The PDF could not be rendered to images.","PDF_RENDER_FAILED");wrapped.cause=error;throw wrapped;}
  finally{fs.rmSync(dir,{recursive:true,force:true});}
}

async function imagesToPdf(files, options = {}) {
  if (!files.length) throw new Error("Choose at least one image.");
  const pdf = await PDFDocument.create();
  const margin = numberOption(options.margin, "margin", { min: 0, max: 200, fallback: 20 });
  const pageSize = String(options.pageSize || "a4").toLowerCase();
  const sizes = { a4: [595.28, 841.89], letter: [612, 792], original: null };
  if (!Object.prototype.hasOwnProperty.call(sizes, pageSize)) throw badRequest("pageSize must be a4, letter, or original.");
  for (const file of files) {
    const isPng = file.mimetype === "image/png" || /\.png$/i.test(file.originalname || "");
    const isJpg = file.mimetype === "image/jpeg" || file.mimetype === "image/jpg" || /\.(jpe?g)$/i.test(file.originalname || "");
    if (!isPng && !isJpg) throw badRequest(`"${file.originalname}" is not a JPG or PNG image.`);
    let image;
    try {
      image = isPng ? await pdf.embedPng(file.buffer) : await pdf.embedJpg(file.buffer);
    } catch (error) {
      const wrapped = badRequest(`"${file.originalname}" is not a valid readable image.`, "INVALID_IMAGE");
      wrapped.cause = error;
      throw wrapped;
    }
    const page = pdf.addPage(sizes[pageSize] || sizes.a4);
    if (pageSize === "original") {
      page.setSize(image.width + margin * 2, image.height + margin * 2);
    }
    const scale = Math.min((page.getWidth() - margin * 2) / image.width, (page.getHeight() - margin * 2) / image.height);
    page.drawImage(image, {
      x: (page.getWidth() - image.width * scale) / 2, y: (page.getHeight() - image.height * scale) / 2,
      width: image.width * scale, height: image.height * scale,
    });
  }
  return savePdf(pdf, "convertly-images.pdf");
}

async function documentToPdf(file) {
  return officeToPdf(file, [".doc", ".docx"]);
}

async function spreadsheetToPdf(file) {
  return officeToPdf(file, [".xls", ".xlsx", ".csv"]);
}

async function officeToPdf(file, allowedExtensions) {
  const ext = path.extname(file.originalname || "").toLowerCase();
  if (!allowedExtensions.includes(ext)) throw badRequest(`Unsupported office format: ${ext || "unknown"}.`);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-"));
  const input = path.join(dir, `input${ext}`);
  try {
    fs.writeFileSync(input, file.buffer);
    await execFileAsync("libreoffice", ["--headless", "--convert-to", "pdf", "--outdir", dir, input], { timeout: 120000, maxBuffer: 2 * 1024 * 1024 });
    const output = path.join(dir, "input.pdf");
    if (!fs.existsSync(output)) throw new Error("LibreOffice did not create a PDF.");
    return { name: safeBaseName(file.originalname) + ".pdf", type: "application/pdf", buffer: fs.readFileSync(output) };
  } catch (error) {
    const wrapped = badRequest(`Could not convert ${ext.toUpperCase()} to PDF.`, "OFFICE_CONVERSION_FAILED");
    wrapped.cause = error;
    throw wrapped;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

async function textPdf(name, text) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  let page = pdf.addPage();
  let y = page.getHeight() - 40;
  for (const line of String(text).split(/\r?\n/).flatMap((value) => value.match(/.{1,105}/g) || [""])) {
    if (y < 40) { page = pdf.addPage(); y = page.getHeight() - 40; }
    page.drawText(line, { x: 30, y, size: 10, font, color: rgb(0.1, 0.12, 0.2) });
    y -= 15;
  }
  return { name, type: "application/pdf", buffer: Buffer.from(await pdf.save({ useObjectStreams: true })) };
}

function pageRange(value, pageCount) {
  if (!value) return Array.from({ length: pageCount }, (_item, index) => index + 1);
  if (typeof value !== "string") throw badRequest("Page ranges must be a string.");
  const pages = new Set();
  for (const part of value.split(",")) {
    const bits = part.trim().split("-").map(Number);
    if (!/^\d+(?:-\d+)?$/.test(part.trim()) || !Number.isInteger(bits[0])) throw badRequest(`Invalid page range "${part.trim()}".`);
    const start = Math.max(1, bits[0]);
    const end = Math.min(pageCount, Number.isFinite(bits[1]) ? bits[1] : start);
    if (start > pageCount || end < 1) throw badRequest(`Page range "${part.trim()}" is outside the document.`);
    for (let page = Math.min(start, end); page <= Math.max(start, end); page += 1) pages.add(page);
  }
  return Array.from(pages).sort((a, b) => a - b);
}

function parsePageNumbers(value, name) {
  if (value === undefined || value === null || value === "") return [];
  const values = Array.isArray(value) ? value : String(value).split(",");
  if (values.length > 1000) throw badRequest(`${name} contains too many page numbers.`);
  return values.map((item) => {
    if (typeof item === "number" && Number.isInteger(item)) return item;
    if (typeof item !== "string" || !/^\d+$/.test(item.trim())) throw badRequest(`${name} must contain whole page numbers separated by commas.`);
    return Number(item.trim());
  });
}

async function splitPdf(file, ranges) {
  const source = await loadPdf(file);
  if (ranges !== undefined && ranges !== null && typeof ranges !== "string" && !Array.isArray(ranges)) {
    throw badRequest("ranges must be a string or an array of page groups.");
  }
  const rangeText = Array.isArray(ranges) ? ranges.join("|") : ranges;
  const groups = String(rangeText || "").trim()
    ? String(rangeText).split(/[;|]/).map((range) => pageRange(range, source.getPageCount()))
    : pageRange("", source.getPageCount()).map((page) => [page]);
  if (groups.some((group) => !group.length)) throw new Error("No valid page range was supplied.");
  const results = [];
  for (let index = 0; index < groups.length; index += 1) {
    const output = await PDFDocument.create();
    const pages = await output.copyPages(source, groups[index].map((page) => page - 1));
    pages.forEach((page) => output.addPage(page));
    results.push(await savePdf(output, `${safeBaseName(file.originalname)}-part-${index + 1}.pdf`));
  }
  return results;
}

async function rewritePdf(file, suffix) {
  const pdf = await loadPdf(file);
  return savePdf(pdf, safeBaseName(file.originalname) + suffix + ".pdf");
}

async function protectPdf(file, options) {
  options = normalizeOptions(options);
  const userPassword = passwordOption(
    options.userPassword === undefined ? options.password : options.userPassword,
    "userPassword",
    true,
  );
  const ownerPassword = passwordOption(options.ownerPassword, "ownerPassword");
  const algorithm = stringOption(options.algorithm, "algorithm", { fallback: "AES-256", maxLength: 20 }).toUpperCase();
  if (!["AES-256", "RC4"].includes(algorithm)) throw badRequest("algorithm must be AES-256 or RC4.");
  const encryptionOptions = {
    ownerPassword,
    algorithm,
    allowPrinting: booleanOption(options.allowPrinting, "allowPrinting"),
    allowModifying: booleanOption(options.allowModifying, "allowModifying"),
    allowCopying: booleanOption(options.allowCopying, "allowCopying"),
    allowAnnotating: booleanOption(options.allowAnnotating, "allowAnnotating"),
    allowFillingForms: booleanOption(options.allowFillingForms, "allowFillingForms"),
  };

  // Validate that the input is a readable, unencrypted PDF before handing it
  // to the encryption package. This also prevents accidentally nesting layers.
  await loadPdf(file);
  try {
    const encrypted = await encryptPDF(new Uint8Array(file.buffer), userPassword, encryptionOptions);
    return {
      name: safeBaseName(file.originalname) + "-protected.pdf",
      type: "application/pdf",
      buffer: Buffer.from(encrypted),
    };
  } catch (error) {
    const wrapped = badRequest("The PDF could not be protected.", "PDF_PROTECTION_FAILED");
    wrapped.cause = error;
    throw wrapped;
  }
}

async function unlockPdf(file, options) {
  options = normalizeOptions(options);
  const password = passwordOption(options.password, "password", true);
  if (!matchesInput(file, "pdf")) throw badRequest("Please provide a PDF file.");
  try {
    const decrypted = await decryptPDF(new Uint8Array(file.buffer), password);
    // Verify the package produced a valid, readable PDF before returning it.
    await PDFDocument.load(decrypted);
    return {
      name: safeBaseName(file.originalname) + "-unlocked.pdf",
      type: "application/pdf",
      buffer: Buffer.from(decrypted),
    };
  } catch (error) {
    const wrongPassword = /wrong password|incorrect password|password does not match/i.test(error.message || "");
    const message = wrongPassword
      ? "The PDF password is incorrect."
      : /not encrypted/i.test(error.message || "")
        ? "The PDF is not password-protected."
        : /unsupported encryption/i.test(error.message || "")
          ? "This PDF uses an encryption format not supported by the local unlocker."
          : "The PDF could not be unlocked. It may be invalid or damaged.";
    const wrapped = badRequest(message, wrongPassword ? "INVALID_PASSWORD" : "PDF_UNLOCK_FAILED");
    wrapped.cause = error;
    throw wrapped;
  }
}

async function compressPdf(file, compression) {
  const level = stringOption(compression, "compression", { fallback: "medium", maxLength: 20 }).toLowerCase();
  if (!["low","medium","high"].includes(level)) throw badRequest("compression must be low, medium, or high.");
  if (level === "low") return rewritePdf(file, "-compressed-low");
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),"convertly-compress-"));
  try {
    const input=path.join(dir,"input.pdf"); fs.writeFileSync(input,file.buffer);
    const dpi=level==="high"?100:130; const quality=level==="high"?55:72;
    await execFileAsync("pdftoppm",["-jpeg","-r",String(dpi),"-jpegopt",`quality=${quality}`,input,path.join(dir,"page")],{timeout:240000,maxBuffer:2*1024*1024});
    const pages=fs.readdirSync(dir).filter(n=>/^page-\d+\.jpg$/i.test(n)).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0]));
    if(!pages.length) throw new Error("No pages rendered.");
    const output=await PDFDocument.create();
    for(const name of pages){const image=await output.embedJpg(fs.readFileSync(path.join(dir,name)));output.addPage([image.width*72/dpi,image.height*72/dpi]).drawImage(image,{x:0,y:0,width:image.width*72/dpi,height:image.height*72/dpi});}
    const raster=await savePdf(output,`${safeBaseName(file.originalname)}-compressed-${level}.pdf`);
    if(level==="high"){const rewritten=await rewritePdf(file,"-compressed-high");if(rewritten.buffer.length<raster.buffer.length)return rewritten;}
    return raster;
  }catch(error){const wrapped=badRequest("The PDF could not be compressed.","PDF_COMPRESSION_FAILED");wrapped.cause=error;throw wrapped;}
  finally{fs.rmSync(dir,{recursive:true,force:true});}
}

async function rotatePdf(file, rotation) {
  const pdf = await loadPdf(file);
  const amount = numberOption(rotation, "rotation", { min: -3600, max: 3600, fallback: 90 });
  pdf.getPages().forEach((page) => page.setRotation(degrees(amount)));
  return savePdf(pdf, safeBaseName(file.originalname) + "-rotated.pdf");
}

async function watermarkPdf(file, text, opacity) {
  const pdf = await loadPdf(file);
  const font = await pdf.embedFont(StandardFonts.HelveticaBold);
  const label = stringOption(text, "text", { fallback: "CONFIDENTIAL", maxLength: 120 });
  const alpha = numberOption(opacity, "opacity", { min: 0.05, max: 1, fallback: 0.25 });
  pdf.getPages().forEach((page) => {
    const size = Math.min(48, Math.max(18, page.getWidth() / 12));
    page.drawText(label, {
      x: page.getWidth() / 2 - font.widthOfTextAtSize(label, size) / 2,
      y: page.getHeight() / 2,
      size, font, color: rgb(0.75, 0.1, 0.1), opacity: alpha, rotate: degrees(35),
    });
  });
  return savePdf(pdf, safeBaseName(file.originalname) + "-watermarked.pdf");
}

async function pageNumbers(file, options = {}) {
  const pdf = await loadPdf(file);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const startAt = numberOption(options.startAt === undefined ? options.start : options.startAt, "startAt", { min: 0, max: 1000000, fallback: 1, integer: true });
  const position = String(options.position || "bottom-center").toLowerCase();
  const positions = {
    bottom: { vertical: "bottom", horizontal: "center" },
    top: { vertical: "top", horizontal: "center" },
    "bottom-left": { vertical: "bottom", horizontal: "left" },
    "bottom-center": { vertical: "bottom", horizontal: "center" },
    "bottom-right": { vertical: "bottom", horizontal: "right" },
    "top-left": { vertical: "top", horizontal: "left" },
    "top-center": { vertical: "top", horizontal: "center" },
    "top-right": { vertical: "top", horizontal: "right" },
  };
  if (!positions[position]) throw badRequest("position must be top, bottom, or a position such as bottom-right.");
  pdf.getPages().forEach((page, index) => {
    const label = String(startAt + index);
    const placement = positions[position];
    const labelWidth = font.widthOfTextAtSize(label, 9);
    const x = placement.horizontal === "left" ? 24 : placement.horizontal === "right" ? page.getWidth() - labelWidth - 24 : (page.getWidth() - labelWidth) / 2;
    page.drawText(label, { x, y: placement.vertical === "top" ? page.getHeight() - 24 : 18, size: 9, font, color: rgb(0.2, 0.2, 0.2) });
  });
  return savePdf(pdf, safeBaseName(file.originalname) + "-numbered.pdf");
}

async function extractText(file) {
  const document = await loadPdfJs(file);
  const pages = [];
  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => item.str).join(" ").replace(/\s+/g, " ").trim());
  }
  return pages;
}

async function loadPdfJs(file) {
  try {
    return await pdfjsLib.getDocument({ data: new Uint8Array(file.buffer) }).promise;
  } catch (error) {
    const wrapped = badRequest(/encrypt|password/i.test(error.message || "")
      ? "This PDF is password-protected and cannot be processed by the local PDF engine."
      : "The PDF is invalid or damaged and could not be processed.", "INVALID_PDF");
    wrapped.cause = error;
    throw wrapped;
  }
}

async function pdfToTextFile(file) {
  const pages = await extractText(file);
  return { name: safeBaseName(file.originalname) + ".txt", type: "text/plain", buffer: Buffer.from(pages.join("\n\n"), "utf8") };
}

async function pdfToMarkdown(file) {
  const pages = await extractText(file);
  const markdown = pages.map((text, index) => `## Page ${index + 1}\n\n${text || "_No selectable text found on this page._"}`).join("\n\n");
  return { name: safeBaseName(file.originalname) + ".md", type: "text/markdown", buffer: Buffer.from(markdown, "utf8") };
}

async function comparePdfs(files) {
  if (files.length < 2) throw new Error("Choose two PDF files to compare.");
  const [left, right] = await Promise.all([extractText(files[0]), extractText(files[1])]);
  const lines = ["# PDF comparison", "", `- Left: ${files[0].originalname}`, `- Right: ${files[1].originalname}`, ""];
  const count = Math.max(left.length, right.length);
  for (let index = 0; index < count; index += 1) {
    if ((left[index] || "") === (right[index] || "")) continue;
    lines.push(`## Page ${index + 1}`, `- **Before:** ${left[index] || "(missing)"}`, `- **After:** ${right[index] || "(missing)"}`, "");
  }
  if (lines.length === 5) lines.push("No text differences found.");
  return { name: "convertly-comparison.md", type: "text/markdown", buffer: Buffer.from(lines.join("\n")) };
}

async function cropPdf(file, options) {
  const pdf = await loadPdf(file);
  const margin = numberOption(options.margin, "margin", { min: 0, max: 10000, fallback: 24 });
  pdf.getPages().forEach((page) => {
    const width = page.getWidth();
    const height = page.getHeight();
    const left = numberOption(options.left, "left", { min: 0, max: width - 1, fallback: margin });
    const right = numberOption(options.right, "right", { min: 0, max: width - left - 1, fallback: margin });
    const top = numberOption(options.top, "top", { min: 0, max: height - 1, fallback: margin });
    const bottom = numberOption(options.bottom, "bottom", { min: 0, max: height - top - 1, fallback: margin });
    if (width - left - right < 1 || height - top - bottom < 1) throw badRequest("Crop margins leave no usable page area.");
    page.setCropBox(left, bottom, width - left - right, height - top - bottom);
  });
  return savePdf(pdf, safeBaseName(file.originalname) + "-cropped.pdf");
}

async function removePages(file, options) {
  const source=await loadPdf(file);
  const remove=parsePageNumbers(options.deletePages || options.pages || options.ranges, "deletePages");
  if(!remove.length) throw badRequest("Enter pages to remove, for example 2,4-6 is not accepted here; use 2,4,5,6.");
  const removeSet=new Set(remove);
  if([...removeSet].some(p=>p<1||p>source.getPageCount())) throw badRequest("deletePages contains an invalid page number.");
  const keep=Array.from({length:source.getPageCount()},(_,i)=>i+1).filter(p=>!removeSet.has(p));
  if(!keep.length) throw badRequest("You cannot remove every page from the PDF.");
  const out=await PDFDocument.create();
  const pages=await out.copyPages(source,keep.map(p=>p-1)); pages.forEach(p=>out.addPage(p));
  return savePdf(out,safeBaseName(file.originalname)+"-pages-removed.pdf");
}

async function extractPages(file, options) {
  const source=await loadPdf(file);
  const pages=pageRange(options.ranges || options.pages, source.getPageCount());
  const out=await PDFDocument.create();
  const copied=await out.copyPages(source,pages.map(p=>p-1)); copied.forEach(p=>out.addPage(p));
  return savePdf(out,safeBaseName(file.originalname)+"-extracted.pdf");
}

async function organizePdf(file, options) {
  const source = await loadPdf(file);
  options = normalizeOptions(options);
  let order;
  if (options.order === undefined || options.order === "") {
    order = [];
  } else if (Array.isArray(options.order)) {
    order = options.order;
  } else {
    try { order = JSON.parse(options.order); } catch (_error) {
      order = String(options.order).split(",").filter(Boolean).map(Number);
    }
  }
  if (!Array.isArray(order) || order.some((page) => !Number.isInteger(page))) throw badRequest("order must be a JSON array or comma-separated page numbers.");
  if (!order.length) order = Array.from({ length: source.getPageCount() }, (_item, index) => index + 1);
  if (order.some((page) => page < 1 || page > source.getPageCount())) throw badRequest("order contains an invalid page number.");
  const deleted = new Set(parsePageNumbers(options.deletePages, "deletePages"));
  if ([...deleted].some((page) => page < 1 || page > source.getPageCount())) throw badRequest("deletePages contains an invalid page number.");
  order = order.filter((page) => !deleted.has(page));
  if (!order.length) throw badRequest("The requested page order contains no pages.");
  const output = await PDFDocument.create();
  const pages = await output.copyPages(source, order.map((page) => page - 1));
  pages.forEach((page) => output.addPage(page));
  return savePdf(output, safeBaseName(file.originalname) + "-organized.pdf");
}

async function pdfArchive(file) {
  const pdf = await loadPdf(file);
  pdf.setTitle(pdf.getTitle() || "Convertly PDF/A export");
  pdf.setSubject("Archive-ready PDF generated by Convertly");
  pdf.setProducer("Convertly");
  return savePdf(pdf, safeBaseName(file.originalname) + "-pdfa.pdf");
}

async function editPdf(file, options) {
  const pdf = await loadPdf(file);
  options = normalizeOptions(options);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const pageNumber = numberOption(options.page, "page", { min: 1, max: pdf.getPageCount(), fallback: 1, integer: true });
  const page = pdf.getPages()[pageNumber - 1];
  if (!page) throw new Error("The requested page does not exist.");
  const text = stringOption(options.text, "text", { fallback: "Edited with Convertly", maxLength: 2000 });
  if (!text.trim()) throw badRequest("text must not be empty.");
  const x = numberOption(options.x, "x", { min: 0, max: page.getWidth(), fallback: 36 });
  const y = numberOption(options.y, "y", { min: 0, max: page.getHeight(), fallback: page.getHeight() - 60 });
  const size = numberOption(options.size, "size", { min: 1, max: 144, fallback: 14 });
  page.drawText(text, { x, y, size, font, color: rgb(0.1, 0.1, 0.1) });
  return savePdf(pdf, safeBaseName(file.originalname) + "-edited.pdf");
}

async function signPdf(file, options) {
  options = normalizeOptions(options);
  const result = await editPdf(file, {
    text: options.signature === undefined ? options.text : options.signature,
    page: options.page === undefined ? 1 : options.page,
    x: options.x === undefined ? 36 : options.x,
    y: options.y === undefined ? 40 : options.y,
    size: options.size === undefined ? 14 : options.size,
  });
  result.name = safeBaseName(file.originalname) + "-signed.pdf";
  return result;
}

async function redactPdf(file, options) {
  const document = await loadPdfJs(file);
  let redactions = jsonOption(options.redactions, "redactions", []);
  if (!Array.isArray(redactions)) throw badRequest("redactions must be a JSON array.");
  if (!redactions.length) { const first=await document.getPage(1); redactions=[{page:1,x:0,y:0,width:Math.min(180,first.getViewport({scale:1}).width),height:24}]; }
  const byPage=new Map();
  redactions.forEach(item=>{if(!item||typeof item!=="object"||Array.isArray(item))throw badRequest("Each redaction must be an object.");const page=numberOption(item.page,"redaction page",{min:1,max:document.numPages,integer:true,fallback:1});const x=numberOption(item.x,"redaction x",{min:0,max:100000,fallback:0});const y=numberOption(item.y,"redaction y",{min:0,max:100000,fallback:0});const width=numberOption(item.width,"redaction width",{min:.01,max:100000});const height=numberOption(item.height,"redaction height",{min:.01,max:100000});if(!byPage.has(page))byPage.set(page,[]);byPage.get(page).push({x,y,width,height});});
  if(redactions.length>1000)throw badRequest("redactions may contain at most 1000 rectangles.");
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),"convertly-redact-"));
  try {
    const input=path.join(dir,"input.pdf");fs.writeFileSync(input,file.buffer);const dpi=144;
    await execFileAsync("pdftoppm",["-png","-r",String(dpi),input,path.join(dir,"page")],{timeout:240000,maxBuffer:2*1024*1024});
    const imgs=fs.readdirSync(dir).filter(n=>/^page-\d+\.png$/i.test(n)).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0]));
    const out=await PDFDocument.create();
    for(let i=0;i<imgs.length;i++){const pageNumber=i+1;const imagePath=path.join(dir,imgs[i]);const sourcePage=await document.getPage(pageNumber);const viewport=sourcePage.getViewport({scale:1});const commands=[];for(const r of byPage.get(pageNumber)||[]){if(r.x+r.width>viewport.width||r.y+r.height>viewport.height)throw badRequest("A redaction rectangle is outside the page.");const x1=Math.round(r.x*dpi/72),x2=Math.round((r.x+r.width)*dpi/72);const y1=Math.round((viewport.height-r.y-r.height)*dpi/72),y2=Math.round((viewport.height-r.y)*dpi/72);commands.push("rectangle",`${x1},${y1} ${x2},${y2}`);}const redactedPath=path.join(dir,`redacted-${pageNumber}.png`);const args=[imagePath,"-fill","#000000"];for(let j=0;j<commands.length;j+=2)args.push("-draw",`${commands[j]} ${commands[j+1]}`);args.push(redactedPath);await execFileAsync("magick",args,{timeout:60000,maxBuffer:1024*1024});const img=await out.embedPng(fs.readFileSync(redactedPath));const page=out.addPage([viewport.width,viewport.height]);page.drawImage(img,{x:0,y:0,width:viewport.width,height:viewport.height});}
    return savePdf(out,safeBaseName(file.originalname)+"-redacted.pdf");
  }catch(error){if(error.statusCode)throw error;const wrapped=badRequest("The PDF could not be redacted.","REDACTION_FAILED");wrapped.cause=error;throw wrapped;}finally{fs.rmSync(dir,{recursive:true,force:true});}
}

async function createForms(file, fields) {
  const pdf = await loadPdf(file);
  const form = pdf.getForm();
  let definitions = jsonOption(fields, "fields", []);
  if (!Array.isArray(definitions)) throw badRequest("fields must be a JSON array.");
  if (!definitions.length) definitions = [{ name: "name", label: "Name" }, { name: "email", label: "Email" }];
  if (definitions.length > 100) throw badRequest("fields may contain at most 100 form fields.");
  const page = pdf.getPages()[0];
  const names = new Set();
  definitions.forEach((field, index) => {
    if (!field || typeof field !== "object" || Array.isArray(field)) throw badRequest("Each form field must be an object.");
    const name = field.name === undefined ? `field-${index + 1}` : stringOption(field.name, "field name", { maxLength: 64 });
    if (!/^[A-Za-z][A-Za-z0-9_.-]{0,63}$/.test(name)) throw badRequest("Form field names must start with a letter and contain only letters, numbers, ., _, or -.");
    if (names.has(name)) throw badRequest(`Form field "${name}" is duplicated.`);
    names.add(name);
    const input = form.createTextField(name);
    const value = field.value === undefined || field.value === null ? "" : stringOption(field.value, "field value", { maxLength: 10000 });
    input.setText(value);
    const y = page.getHeight() - 80 - index * 34;
    if (y < 22) throw badRequest("Too many form fields for the first page.");
    const x = Math.min(48, Math.max(0, page.getWidth() - 14));
    const width = Math.min(240, page.getWidth() - x - 12);
    const height = Math.min(22, page.getHeight() - y);
    if (width < 1 || height < 1) throw badRequest("The first page is too small for form fields.");
    input.addToPage(page, { x, y, width, height, borderWidth: 1 });
  });
  return savePdf(pdf, safeBaseName(file.originalname) + "-form.pdf");
}

async function summarizePdf(file) {
  const pages = await extractText(file);
  const text = pages.join(" ").replace(/\s+/g, " ").trim();
  const sentences = text.match(/[^.!?]+[.!?]+/g) || [];
  const summary = ["# Summary", "", text ? sentences.slice(0, 5).join(" ").trim() : "No selectable text was found in this PDF.", "", `Pages: ${pages.length}`].join("\n");
  return { name: safeBaseName(file.originalname) + "-summary.md", type: "text/markdown", buffer: Buffer.from(summary, "utf8") };
}

async function ocrPdf(file, options = {}) {
  const language = stringOption(options.ocrLang || options.lang, "ocrLang", { fallback: "eng", maxLength: 100 });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-ocr-"));
  try {
    const input = path.join(dir, "input.pdf");
    fs.writeFileSync(input, file.buffer);
    await execFileAsync("pdftoppm", ["-png", "-r", "160", input, path.join(dir, "page")], { timeout: 180000, maxBuffer: 2 * 1024 * 1024 });
    const pages = fs.readdirSync(dir).filter(n => /^page-\d+\.png$/i.test(n)).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0]));
    if (!pages.length) throw new Error("Could not render PDF pages for OCR.");
    const ocrPdfs=[];
    for (let i=0;i<pages.length;i++) {
      const base=path.join(dir, `ocr-${i+1}`);
      await execFileAsync("tesseract", [path.join(dir,pages[i]), base, "pdf", "-l", language], { timeout: 180000, maxBuffer: 2 * 1024 * 1024 });
      ocrPdfs.push(base + ".pdf");
    }
    const merged=await PDFDocument.create();
    for(const pdfPath of ocrPdfs){const source=await PDFDocument.load(fs.readFileSync(pdfPath));const pages=await merged.copyPages(source,source.getPageIndices());pages.forEach(p=>merged.addPage(p));}
    return { name: safeBaseName(file.originalname) + "-ocr.pdf", type: "application/pdf", buffer: Buffer.from(await merged.save({useObjectStreams:true})) };
  } catch(error) {
    const wrapped=badRequest("OCR could not process this PDF. Check that the requested Tesseract language is installed.", "OCR_FAILED"); wrapped.cause=error; throw wrapped;
  } finally { fs.rmSync(dir,{recursive:true,force:true}); }
}

async function translatePdf(file, target) {
  const lang=String(target||"en").toLowerCase();
  if(!/^[a-z]{2,5}$/.test(lang)) throw badRequest("target must be a language code such as en, hi, or gu.");
  const pages=await extractText(file);
  const translated=[];
  for(const page of pages){
    if(!page){translated.push("");continue;}
    const chunks=page.match(/.{1,1800}(?:\s|$)/g)||[page];
    const out=[];
    for(const chunk of chunks){
      const url=`https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${encodeURIComponent(lang)}&dt=t&q=${encodeURIComponent(chunk)}`;
      const response=await fetch(url);
      if(!response.ok) throw new Error(`Translation service returned ${response.status}.`);
      const data=await response.json();
      out.push(Array.isArray(data?.[0])?data[0].map(x=>x?.[0]||"").join(""):"");
    }
    translated.push(out.join(" "));
  }
  return textPdf(safeBaseName(file.originalname)+`-translated-${lang}.pdf`, translated.map((t,i)=>`Page ${i+1}\n${t}`).join("\n\n"));
}

async function pdfToWord(file) {
  const text = (await extractText(file)).join("\n\n");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.folder("_rels").file(".rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.folder("word").file("document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${escapeXml(text).split(/\r?\n/).map((line) => `<w:p><w:r><w:t xml:space="preserve">${line || " "}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`);
  return { name: safeBaseName(file.originalname) + ".docx", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", buffer: await zip.generateAsync({ type: "nodebuffer" }) };
}

async function pdfToExcel(file) {
  const rows = (await extractText(file)).map((text, index) => [index + 1, text]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Page", "Text"], ...rows]), "PDF text");
  return { name: safeBaseName(file.originalname) + ".xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) };
}

async function pdfToPowerpoint(file) {
  const document = await loadPdfJs(file);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-ppt-"));
  try {
    const pdfPath = path.join(dir, "input.pdf");
    fs.writeFileSync(pdfPath, file.buffer);
    await execFileAsync("pdftoppm", ["-png", "-r", "120", pdfPath, path.join(dir, "page")], { timeout: 120000, maxBuffer: 2 * 1024 * 1024 });
    const images = fs.readdirSync(dir).filter(n => /^page-\d+\.png$/i.test(n)).sort((a,b)=>Number(a.match(/\d+/)[0])-Number(b.match(/\d+/)[0]));
    if (!images.length) throw new Error("Could not render PDF pages.");
    const presentation = new pptxgen();
    presentation.layout = "LAYOUT_WIDE";
    for (const name of images) {
      const slide = presentation.addSlide();
      slide.background = { color: "FFFFFF" };
      slide.addImage({ path: path.join(dir, name), x: 0, y: 0, w: 13.333, h: 7.5 });
    }
    return { name: safeBaseName(file.originalname) + ".pptx", type: "application/vnd.openxmlformats-officedocument.presentationml.presentation", buffer: await presentation.write({ outputType: "nodebuffer" }) };
  } catch (error) {
    const wrapped = badRequest("Could not convert the PDF to PowerPoint.", "PDF_TO_PPT_FAILED");
    wrapped.cause = error;
    throw wrapped;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

async function powerpointToPdf(file) {
  return officeToPdf(file, [".ppt", ".pptx"]);
}

async function htmlToPdf(html) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "convertly-html-"));
  const input = path.join(dir, "input.html");
  try {
    const source = String(html || "").trim();
    if (!source) throw badRequest("Provide HTML content.");
    fs.writeFileSync(input, source, "utf8");
    await execFileAsync("libreoffice", ["--headless", "--convert-to", "pdf", "--outdir", dir, input], { timeout: 120000, maxBuffer: 2 * 1024 * 1024 });
    const output = path.join(dir, "input.pdf");
    if (!fs.existsSync(output)) throw new Error("LibreOffice did not create a PDF.");
    return { name: "convertly-html.pdf", type: "application/pdf", buffer: fs.readFileSync(output) };
  } catch (error) {
    if (error.statusCode) throw error;
    const wrapped = badRequest("Could not convert the HTML document to PDF.", "HTML_CONVERSION_FAILED");
    wrapped.cause = error;
    throw wrapped;
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function escapeXml(value) {
  return String(value).replace(/[<>&'"]/g, (character) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", "\"": "&quot;" }[character]));
}

app.use((error, _req, res, _next) => {
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    return res.status(413).json({ error: "Each file must be 25 MB or smaller.", code: error.code });
  }
  if (error instanceof multer.MulterError) {
    return res.status(400).json({ error: "The upload is invalid. Send files using the 'files' field.", code: error.code });
  }
  res.status(error.statusCode || 500).json({
    error: error.message || "Conversion failed.",
    ...(error.code ? { code: error.code } : {}),
  });
});

if (require.main === module) {
  const server = app.listen(port, () => {
    console.log(`Convertly running at http://localhost:${port}`);
  });
  server.on("error", (error) => {
    if (error.code === "EADDRINUSE") {
      console.error(`Port ${port} is already in use. Close the older Convertly server, then start this project again.`);
      process.exitCode = 1;
      return;
    }
    console.error("Convertly could not start:", error.message);
    process.exitCode = 1;
  });
}

module.exports = { app, convert, supportedTools, toolSpecs };
