# Dockerfile for ContaGO API

FROM node:20-slim AS builder

WORKDIR /app

ENV PUPPETEER_SKIP_DOWNLOAD=true

# Copy package files
COPY package*.json ./
COPY scripts ./scripts

# Install dependencies (incl. dev for build)
RUN npm install

# Copy source
COPY tsconfig.json ./
COPY src ./src
COPY templates ./templates

# Build TypeScript
RUN npm run build

FROM node:20-slim

# Install Chromium dependencies
RUN apt-get update && apt-get install -y \
    chromium \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    xdg-utils \
    python3 \
    make \
    g++ \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Set Puppeteer to use installed Chromium
ENV NODE_ENV=production
ENV PUPPETEER_SKIP_DOWNLOAD=true
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium

WORKDIR /app

# Copy package files
COPY package*.json ./
COPY scripts ./scripts

# Install production dependencies
RUN npm install --omit=dev

# Copy build output
COPY --from=builder /app/dist ./dist
COPY --from=builder /app/templates ./templates

# Create directories writable by node user
RUN mkdir -p /app/downloads /app/.cache/puppeteer \
    && chown node:node /app/downloads \
    && chown -R node:node /app/.cache

# Expose port
EXPOSE 8000

USER node

# Start server
CMD ["node", "dist/index.js"]
