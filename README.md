# Author
Omkar Tikudave – Data Engineer

# Excel-Pivot-Automation App

## Overview

This Java application automates the creation of Pivot Tables in Excel based on a JSON configuration file.

## Features

- Reads an Excel (.xlsx) file and a JSON configuration.
- Generates pivot tables per sheet as defined.
- Saves the result to a new Excel file with pivot sheets.

---

## Input

- **args[0]**: Path to the input Excel file.
- **args[1]**: JSON configuration string.
- **args[2]**: Path to the output Excel file.

### Sample JSON:
'{
  "house_data": [
    {
      "sheet_name": "pivot_house_data",
      "dimension": ["Unit", "Camp Name"],
      "matrix": [
        { "column": "Date", "operation": "column" },
        { "column": "House Id", "operation": "Count" }
      ],
      "Filter": ["Sub Group"]
    }
  ]
}'


# Setup Instructions
sudo apt update
sudo apt install openjdk-8-jdk

# Install Maven
sudo apt install maven

# Check Version

java -version
mvn -version

# Sample Call of the attached Input and Output sheet 

java -cp /home/yukta-13/Desktop/Personal_projects/Excel-Pivot-Automation/pivot_automation/target/pivot_automation-0.0.1-SNAPSHOT.jar com.pivot.pivot_automation.App "/home/yukta-13/Desktop/Personal_projects/Excel-Pivot-Automation/sample_sales_data.xlsx" '{
    "Sales_data": [
        {
            "sheet_name": "Pivot_1",
            "dimension": [
                "Product"
            ],
            "matrix": [
                {
                    "column": "Total_sales",
                    "operation": "sum"
                }
            ],
            "Filter": [
                "Region"
            ]
        },
        {
            "sheet_name": "Pivot_2",
            "dimension": [
                "Region",
                "Product"
            ],
            "matrix": [
                {
                    "column": "Total_sales",
                    "operation": "sum"
                }
            ],
            "Filter": []
        },
        {
            "sheet_name": "Pivot_3",
            "dimension": [
                "Region",
                "Date"
            ],
            "matrix": [
                {
                    "column": "Total_sales",
                    "operation": "sum"
                }
            ],
            "Filter": []
        },
        {
            "sheet_name": "Pivot_3",
            "dimension": [
                "Region",
                "Product"
            ],
            "matrix": [
                {
                    "column": "Date",
                    "operation": "column"
                },
                {
                    "column": "Quantity",
                    "operation": "sum"
                }
            ],
            "Filter": []
        }
    ],
    "New_sales_data": [
        {
            "sheet_name": "New_Pivot_1",
            "dimension": [
                "Product"
            ],
            "matrix": [
                {
                    "column": "Quantity",
                    "operation": "avg"
                }
            ],
            "Filter": [
                "Region"
            ]
        }
    ]
}' "/home/yukta-13/Desktop/Personal_projects/Excel-Pivot-Automation/pivot_sales_data_report.xlsx"
