FROM node:18-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production

# Copy application files
COPY server.js ./
COPY public/ ./public/

# Create necessary directories
RUN mkdir -p uploads output

# Expose port
EXPOSE 3000

# Start the application
CMD ["node", "server.js"]
