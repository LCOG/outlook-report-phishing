# syntax=docker/dockerfile:1

# Dockerfile created using docker init
# See https://docs.docker.com/guides/nodejs/containerize for context.

ARG NODE_VERSION=24.13.0

#===============================================================================
# Use node image for base image for all stages.
#===============================================================================
FROM node:${NODE_VERSION}-alpine AS base

# Set working directory for all build stages.
WORKDIR /app

# Create non-root user for security
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nodejs -u 1001 -G nodejs && \
    chown -R nodejs:nodejs /app

# Enable pnpm via corepack
# pnpm version is determined by packageManager field in package.json
RUN corepack enable

#===============================================================================
# Production dependencies
#===============================================================================
FROM base AS deps

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./

# Download dependencies as a separate step to take advantage of Docker's caching.
# Leverage a cache mount to /root/.local/share/pnpm/store to speed up subsequent builds.
RUN --mount=type=cache,target=/root/.local/share/pnpm/store \
    pnpm install --prod --frozen-lockfile && \
    chown -R nodejs:nodejs /app

#===============================================================================
# Build dependencies
#===============================================================================
FROM base AS build-deps

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./

# Download additional development dependencies before building, as some projects require
# "devDependencies" to be installed to build. If you don't need this, remove this step.
RUN --mount=type=cache,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile && \
    chown -R nodejs:nodejs /app

#===============================================================================
# Build stage
#===============================================================================
FROM build-deps AS build

# Copy only necessary files for building (respects .dockerignore)
COPY --chown=nodejs:nodejs . .

# Run the build script.
RUN pnpm run build

# Set proper ownership
RUN chown -R nodejs:nodejs /app/dist

#===============================================================================
# Development stage
#===============================================================================
FROM build-deps AS development

# Set environment
ENV NODE_ENV=development \
    NPM_CONFIG_LOGLEVEL=warn

# Copy source files
COPY . .

# Ensure all directories have proper permissions
RUN chown -R nodejs:nodejs /app && \
    chmod -R 755 /app

# Switch to non-root user
USER nodejs

# Expose ports
EXPOSE 3000

# Start development server
CMD ["pnpm", "run", "dev-server"]

#===============================================================================
# Create a new stage to run the application with minimal runtime dependencies
# where the necessary files are copied from the build stage.
#===============================================================================
FROM base AS final

# Use production node environment by default.
ENV NODE_ENV=production

# Run the application as a non-root user.
USER node

# Copy package.json so that package manager commands can be used.
COPY package.json .

# Copy the production dependencies from the deps stage and also
# the built application from the build stage into the image.
COPY --from=deps /usr/src/app/node_modules ./node_modules
COPY --from=build /usr/src/app/dist ./dist


# Expose the port that the application listens on.
EXPOSE 3000

# Run the application.
CMD ["pnpm", "run", "dev-server"]
