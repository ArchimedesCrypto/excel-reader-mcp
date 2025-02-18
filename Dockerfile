# Generated by https://smithery.ai. See: https://smithery.ai/docs/config#dockerfile
# Stage 1: Build the application
FROM node:22-alpine AS builder

# Set the working directory
WORKDIR /app

# Copy package files and install dependencies
COPY package.json package-lock.json ./
RUN npm install --ignore-scripts

# Copy the entire source code and compile TypeScript to JavaScript
COPY tsconfig.json tsconfig.json
COPY src src
RUN npm run build

# Stage 2: Run the application
FROM node:22-alpine AS release

# Set the working directory
WORKDIR /app

# Copy the compiled JavaScript files from the builder stage
COPY --from=builder /app/build /app/build
COPY --from=builder /app/package.json /app/package.json
COPY --from=builder /app/package-lock.json /app/package-lock.json

# Install only production dependencies
RUN npm ci --omit=dev

# Set the entry point to the application
ENTRYPOINT ["node", "build/index.js"]
