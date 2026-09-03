# Convertly — PDF Workspace

A polished local PDF toolkit inspired by the simplicity of modern PDF utilities. The UI is redesigned from scratch and the tool catalogue is wired to the local API.

## Included
- Merge, split, remove, extract, organize and rotate PDFs
- Compress, repair and crop PDFs
- Watermark, page numbers, edit, sign, forms and redaction
- Protect and unlock PDFs
- OCR, local summarizer, compare and translate extracted PDF text
- PDF ↔ JPG, Word, Excel, PowerPoint and HTML workflows
- Image compression and image/scan to PDF
- Favorites, recent activity, search, responsive UI and favicon

## Run
1. Install Node.js 18+.
2. On Windows, double-click `start.bat` or run `npm install` then `npm start`.
3. Open `http://localhost:8080`.

## Public deployment

This repository includes `Dockerfile` and `render.yaml` for a full Convertly deployment with LibreOffice, Poppler, ImageMagick, and Tesseract included.

1. Upload the project to a GitHub repository.
2. In Render, select **New → Blueprint** and select that repository.
3. Render reads `render.yaml`, builds the Docker image, and publishes a public `onrender.com` URL.
4. Set a custom domain in Render, then replace `https://convertly.app/` in `index.html`, `robots.txt`, and `sitemap.xml` with that domain before the next deploy.

The included health check is available at `/api/health`.

### System conversion engines

Install these command-line tools and make sure they are on your `PATH` to enable every Convertly workflow:

- LibreOffice (`libreoffice`) — Word, Excel, PowerPoint, and HTML to PDF
- Poppler (`pdftoppm`) — PDF to JPG, PDF compression, redaction, OCR, and PDF to PowerPoint
- ImageMagick (`magick`) — image compression and redaction
- Tesseract (`tesseract`) — OCR

Use `GET /api/health` or `GET /api/tools` to see exactly which engines the running backend detects. Tools that do not need a native engine work with the Node dependencies alone.

## Security note
This is a local/self-hosted prototype. Uploaded files are kept in memory by the Express API and conversion history is held in memory only. Add authentication, rate limiting, file scanning and production storage policies before public deployment.
