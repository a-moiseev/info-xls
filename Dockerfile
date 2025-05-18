FROM --platform=linux/amd64 python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y \
    binutils \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
COPY . .

RUN pip install -r requirements.txt

CMD pyinstaller --clean --onefile --name info-xls \
    --add-data "config.yaml:." \
    --hidden-import xlwings \
    --hidden-import openpyxl \
    main.py
