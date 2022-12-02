# Bloque_4

## A simple data analysis of the orders of a pizza restaurant in order to optimize the weekly ingredient purchases.
This project consists of a simple pandas program which calculates the mode of the quantity spent each week for each ingredient.
Also, it supplies a data analysis of nulls and nans for each table.
This program makes use of pandas in order to calculate the quantity spent of ingredients every week. It also gives you a data analysis of nulls and nans for each file.
Added support for analyzing data from 2016. This new data had to be cleaned so I also added a new python script to do this.
New files to support this, were also created by modifying the original ones.
Added support to save data into XML files. This files include a data typology analysis and the final reccomendation.

## Practica_1
Added support for creating an executive report in pdf using matplotlib to be able to analyze the data and visualize it.

## Practica_2
Added support for creating an excel report using pandas and xlsxwriter. Generating a report in excel of the executive part, the ingredient part and the orders part.

### Instructions:
To execute the program, run "pizzas_maven_ejecutivo.py" or "pizzas_maven_excel.py"
Also, it is possible to create a docker image to deploy the program in a safer way.
To do that, just run the following command in the console, inside the directory where you clone this repository:

docker build . -t Bloque_4
