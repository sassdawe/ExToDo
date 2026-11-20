FROM node:24-alpine

ENV NODE_ENV=production
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --ignore-scripts && npm cache clean --force
COPY server.js LICENSE ./
COPY public/ ./public/

USER node
EXPOSE 8080
CMD ["node", "server.js"]
