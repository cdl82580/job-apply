FROM python:3.12-slim

# Install Node.js 20 + pandoc in one layer
RUN apt-get update && apt-get install -y --no-install-recommends curl pandoc \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python deps first (layer-cache friendly)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Node deps (docx package for cover letter + ATS resume generation)
COPY package.json .
RUN npm install

# Application code
COPY . .

RUN mkdir -p output

EXPOSE 8080

CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8080"]
