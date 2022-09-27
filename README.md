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
Ensure the above is available within the working environment that package will be loaded and used. Run **pip show pandas** and **pip show openpyxl**, to check the versions of the packages are the same or newer listed above.

![image](https://user-images.githubusercontent.com/59668937/184554267-63fc77b9-1711-4b4b-872a-8c0bb966e0e5.png)


## What's in the box
This current release allows the user to **Dynamically create worksheets from querying the values in a dataframe imported from parquet**. Follow up releases will include the remaining planned features listed in Overview. A folder of mocked parquet data for a list of cars, is used to demonstrate the functionality of the solution.

## How to use
- Clone the repo working directory 

![image](https://user-images.githubusercontent.com/59668937/184549343-ed8934b5-5e0d-4b7e-8af2-72267057b461.png)

- Change to the working directory, in this example it will be the directory that was created for the clone

![image](https://user-images.githubusercontent.com/59668937/184549511-d07549fe-569b-4caf-ac21-cc06217da2a8.png)

- The following command is used to create and save Excel workbook called **excel_template.xlsx**. It will use the **Make** column from the dataframe read in from the mocked parquet data. This column includes a list a cars , that will create a worksheet for each car in the workbook.

![image](https://user-images.githubusercontent.com/59668937/184554028-6f4be6c4-7931-4335-a629-9bce40d42000.png)


![image](https://user-images.githubusercontent.com/59668937/184550427-5de80561-ff53-42e6-bd24-f53bb25db62a.png)

- This is the code executed in a python command shell

![image](https://user-images.githubusercontent.com/59668937/184554099-d67491dc-756a-4bb4-a1d6-8d65231e664a.png)

- Check the working directory where the excel file would have been saved, and open up the workbook to review the worksheets created

![image](https://user-images.githubusercontent.com/59668937/184554140-879bb044-7087-4a3e-8680-77be052c84c2.png)


![image](https://user-images.githubusercontent.com/59668937/184554181-d322ebdf-4749-4a61-a5a1-faf26949b025.png)




