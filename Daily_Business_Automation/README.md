# Daily Business Performance Automation System

## Purpose
Automates business reporting workflow:
- Cleans raw sales/transaction data
- Generates formatted Excel reports
- Creates dashboard with pivot and chart
- Sends email summary to manager

## Folder Structure
- raw_data_1000.xlsx : Input data
- automation.py       : Main script
- config.json         : Email config
- reports/            : Output reports and dashboard

## How to Run
1. Update email credentials in config.json
2. Run: python automation.py
3. Reports and dashboard saved in reports/
4. Summary email sent automatically
