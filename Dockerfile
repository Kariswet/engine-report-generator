FROM python:3.10-slim
WORKDIR /app

COPY source/requirements.txt /app/requirements.txt
RUN pip install --upgrade pip
RUN pip install -r requirements.txt
COPY source/. /app
ENV PYTHONUNBUFFERED=1
ENV TZ=Asia/Jakarta
ENTRYPOINT ["python3","main.py"]
