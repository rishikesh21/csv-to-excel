# CSV to Excel Converter

A project to converts a set of csv files into a single Excel Workbook

### Prerequisites

What things you need to install the software and how to install them

```
Maven
Java 8
```

## Getting Started
1. Clone your project 
```
git clone https://github.com/rishikesh21/csv-to-excel.git
```
2. Open in you IDE
3. Run 
```
clean install package -DSkipTests
```

The jar will be in your target folder.

## How To Run 
```
java -jar excelconverter-1.0-SNAPSHOT-shaded.jar file_1.csv file_2.csv file_3.csv result
```
file_1.csv -Input file 1 
file_2.csv -Input file 2
file_3.csv -Input file 3
result -resulting Excel File
