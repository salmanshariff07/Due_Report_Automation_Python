# Sales Report Automation with Python
## Automating country-wise sales reports for faster insights and smarter tracking

## Project Overview
This Python project automates the entire sales reporting process — from raw data to clean, country-specific Excel reports — enabling teams to track regional performance instantly and reduce manual work.

Using a single sales data file, the script reads, filters, splits, and populates country-wise sales reports based on a predefined Excel template. It also evaluates performance for each country and provides summary comments to help interpret key metrics.

<a tref= "https://github.com/salmanshariff07/Due_Report_Automation_Python/blob/main/Sales_Data.xlsx"> Sales_Data

## How It Works
Input Sales Data
The script reads a master SalesData.xlsx file that contains raw sales records.

Country-Based Filtering
Based on a list of countries provided in the ‘Summary’ sheet, the script filters the sales data for each country.

Template-Based Report Creation
For each country, it:

Opens a preformatted Excel report template
Populates data such as Total Sales, Amount Received, and Balance Due
<a tref="https://github.com/salmanshariff07/Due_Report_Automation_Python/blob/main/Template.xlsx"> Template

Saves a new Excel report using the country name
Performance Evaluation

After each report is generated, the terminal outputs:
Country-wise totals

## A comment on performance (e.g. “Doing Great” or “Needs Improvement”) based on company thresholds as below

<a tref="https://github.com/salmanshariff07/Due_Report_Automation_Python/blob/main/ReportGenerator.py"> ReportGenerator_Code

![Generation](https://github.com/user-attachments/assets/1d756e50-9610-4590-afb0-765cb21c9b57)


## Key Features
Automated multi-country report generation
Exports clean Excel files using a predefined template
Includes sales summaries and balance analysis per country
Performance feedback provided via terminal
Eliminates manual sorting and calculation efforts

