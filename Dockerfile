FROM mcr.microsoft.com/devcontainers/python:1-3.11-bookworm

WORKDIR /app

COPY requirements.txt ./requirements.txt

RUN if [ -f requirements.txt ]; then \
    pip install --no-cache-dir -r requirements.txt; \
    fi && \
    pip install --no-cache-dir streamlit

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.enableCORS=false", "--server.enableXsrfProtection=false", "--server.address=0.0.0.0"]
