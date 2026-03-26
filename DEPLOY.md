# Streamlit Cloud Deployment

## Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

1. Push this repository to GitHub (make sure `generators/` is included).

2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.

3. Select your repo, branch, and set **Main file path** to `app.py`.

4. Click **Deploy**.

### Font note

The PDF generator registers DejaVu or Liberation Sans fonts if they exist at
standard Linux paths. On Streamlit Cloud (Ubuntu) these are typically available:

```
/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf
/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf
```

If fonts are missing the generator falls back to Helvetica (a built-in ReportLab
font), so the app will still work.

### packages.txt (optional, for guaranteed font availability)

Create a `packages.txt` in the repo root to install system fonts:

```
fonts-dejavu-core
```

## Docker (self-hosted)

```dockerfile
FROM python:3.11-slim
RUN apt-get update && apt-get install -y fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

```bash
docker build -t esg-doc-generator .
docker run -p 8501:8501 esg-doc-generator
```
