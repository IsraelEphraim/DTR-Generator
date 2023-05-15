FROM python:3.9

WORKDIR /app

COPY main.py .
COPY template template/

RUN pip install flask
RUN pip install pandas
RUN pip install openpyxl

CMD ["python", "./main.py"]
