# Dockerfile, Image, Container
FROM python:3.9

ADD pizzas_maven_excel.py .
ADD order_details.csv .
ADD orders.csv .
ADD pizzas.csv .
ADD pizza_types.csv .

RUN pip install -r requirements.txt


CMD ["python", "./pizzas_maven_excel.py"]
