# parquet-to-excel

## Overview
Solution to dynamically build, format and load data from parquet and json configuration files. The solution will include features to perform the following tasks:
- Dynamically create worksheets from querying the values in a dataframe imported from parquet
- Load data into the worksheet using specific locations, for single and multiple loads
- Format the worksheets using the json styles sheet

The solution is developed and tested using the following dependencies:
- Python 3.10
- Pandas 1.4
- Openpyxl 3.0.10
Ensure the above is available within the working environment that package will be loaded and used.

## What's in the box
This current release allows the user to **Dynamically create worksheets from querying the values in a dataframe imported from parquet**. Follow up releases will include the remaining planned features listed in Overview. A folder of mocked parquet data for a list of cars, is used to demonstrate the functionality of the solution.

## How to use
- Clone the repo working directory 

![image](https://user-images.githubusercontent.com/59668937/184549343-ed8934b5-5e0d-4b7e-8af2-72267057b461.png)

- Change to the working directory, in this example it will be the directory that was created for the clone

![image](https://user-images.githubusercontent.com/59668937/184549511-d07549fe-569b-4caf-ac21-cc06217da2a8.png)

- The following command is used to create and save Excel workbook called **excel_template.xlsx**. It will use the **Make** column from the dataframe read in from the mocked parquet data. This column includes a list a cars , that will create a worksheet for each car in the workbook.

![image](https://user-images.githubusercontent.com/59668937/184549777-6af2f426-9ac1-4586-b4aa-2f5b9ba559dc.png)



