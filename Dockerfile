FROM python:3.11-slim

WORKDIR /app

RUN pip install --no-cache-dir flask curl-cffi openpyxl gunicorn

COPY app.py .

ENV HOST=0.0.0.0 PORT=7860

EXPOSE 7860

# fetch timeout is 90s inside the app -> worker timeout 180s
CMD ["gunicorn", "-w", "2", "--threads", "4", "-t", "180", "-b", "0.0.0.0:7860", "app:app"]
