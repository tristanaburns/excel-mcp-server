FROM python:3.10-slim

WORKDIR /app

# Install uv for Python package management
RUN pip install uv

# Copy project files
COPY pyproject.toml uv.lock README.md ./
COPY src ./src

# Install the package in development mode (using --system flag)
RUN uv pip install --system -e .

# Install dev dependencies for testing
RUN uv pip install --system flake8 mypy pytest

# Create directory for Excel files and set permissions
RUN mkdir -p /app/excel_files && chmod 777 /app/excel_files

# Copy health check script, test scripts, and entrypoint
COPY health/health.py ./
COPY tests/test_*.py ./
COPY build/docker-entrypoint.sh ./
COPY build/lint.sh build/test.sh build/build.sh ./
RUN chmod +x /app/docker-entrypoint.sh /app/lint.sh /app/test.sh /app/build.sh

# Set environment variables
ENV EXCEL_FILES_PATH=/app/excel_files
ENV FASTMCP_PORT=8000
ENV FASTAPI_PORT=8080
ENV PYTHONUNBUFFERED=1

# Expose ports for MCP server and web interface
EXPOSE 8000 8080

# Run the server
ENTRYPOINT ["/app/docker-entrypoint.sh"]