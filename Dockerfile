# ---- Base
FROM node:22-slim

# Keep layer small; install LO + poppler + fonts (Noto covers most scripts)
# Using tini for clean SIGTERM handling in containers.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    poppler-utils \
    fonts-dejavu \
    fonts-liberation \
    fonts-noto \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    tini \
 && rm -rf /var/lib/apt/lists/*

# Optional: if you want Arabic or other locales strongly rendered,
# you can add e.g. fonts-noto-extra, fonts-noto-unhinted, etc.

ENV NODE_ENV=production \
    # let your Node code find soffice at the standard path
    SOFFICE_BIN=/usr/bin/soffice

WORKDIR /app

# Install deps first for better caching
COPY package*.json ./
RUN npm ci --omit=dev

# App source
COPY . .

# Run as non-root (recommended)
RUN chown -R node:node /app
USER node

EXPOSE 8004

# tini forwards signals so LibreOffice/child procs exit cleanly 
ENTRYPOINT ["/usr/bin/tini", "--"]

CMD ["node", "server.js"]

# docker build -t docker.io/inteligeninfosys/demand-letters-api-sid:2026022130 .