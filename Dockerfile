# Use Node.js LTS version that matches your project requirements
FROM node:18-alpine

# Set working directory
WORKDIR /app

# Install dependencies for building native modules if needed
RUN apk add --no-cache python3 make g++

# Copy package files first for better layer caching
COPY package*.json ./

# Install dependencies
RUN npm ci --prefer-offline --no-audit

# Copy the rest of the application
COPY . .

# Build the application (if needed)
RUN npm run build || echo "Build step optional"

# Default command runs tests
CMD ["npm", "test"]

