FROM node:22-alpine AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src/ ./src/
RUN npm run build

FROM node:22-alpine

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY --from=builder /app/dist/ ./dist/
COPY docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh

ENV MCP_TRANSPORT=http
ENV MCP_PORT=3001

ARG DOCKER_GID=985
# docker-cli: the server spawns every sandbox with `docker run` against the
# host's socket, so the binary has to be in this image.
RUN apk add --no-cache su-exec docker-cli && \
    chmod +x /usr/local/bin/docker-entrypoint.sh && \
    addgroup -S mcp && adduser -S mcp -G mcp && \
    (addgroup -g ${DOCKER_GID} docker 2>/dev/null || addgroup docker) && \
    adduser mcp docker && \
    mkdir -p /app/.tmp && chown mcp:mcp /app/.tmp

EXPOSE 3001

ENTRYPOINT ["docker-entrypoint.sh"]
CMD ["node", "dist/index.js"]
