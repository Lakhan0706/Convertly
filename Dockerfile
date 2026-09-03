FROM node:22-bookworm-slim

ENV NODE_ENV=production
WORKDIR /app

# Native engines used by Convertly's office conversion, rendering, image,
# compression, redaction, and OCR tools.
RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    imagemagick \
    tesseract-ocr \
  && rm -rf /var/lib/apt/lists/*

COPY package*.json ./
RUN npm ci --omit=dev

COPY . ./

EXPOSE 10000
CMD ["node", "server.js"]
