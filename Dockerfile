FROM python:latest

WORKDIR /usr/app

COPY vendas-combustiveis-m3.xlsx .
COPY fuel-sales-bfaleiros-test.py .

RUN pip install openpyxl pandas

CMD ["python", "./fuel-sales-bfaleiros-test.py"]