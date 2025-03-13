# Jacquard Textile Reporting Software

## Overview
This Python-based application is designed to streamline data management and reporting for Jacquard textile operations. It facilitates efficient handling of customer information, weaving items, lacing details, and weaver data, ensuring accurate and organized record-keeping.

## Features
- **Data Entry:** Easily input and manage data for customers, items, lacing, and weavers.
- **Payment Management:** Track and manage payments related to customers and lacing processes.
- **Report Generation:** Generate comprehensive reports for customers and lacing activities to gain insights into operations.

## Prerequisites
Ensure you have Python installed on your system. If not, download it from the [official Python website](https://www.python.org/downloads/).

## Installation
1. **Clone the Repository:**
   ```sh
   git clone https://github.com/Vimalraj002/jaquard-textile-report-sw.git
   ```

2. **Navigate to the Project Directory:**
   ```sh
   cd jaquard-textile-report-sw
   ```
## Usage
1. **Data Entry:**
   - Use `dataEntry.py` to input and manage data.
   - Ensure your Excel files (`Customers.xlsx`, `Items.xlsx`, `Lacing.xlsx`, `Weaver.xlsx`) are updated with the latest information.

2. **Manage Payments:**
   - `managePaymentsCustomer.py`: Handle customer payment records.
   - `managePaymentsLacing.py`: Manage payments related to lacing processes.

3. **Generate Reports:**
   - `reportGenerateCustomer.py`: Produce reports detailing customer-related data.
   - `reportGenerateLacing.py`: Generate reports focusing on lacing activities.

4. **Main Execution:**
   - Run `main.py` to execute the primary functionalities of the application.

## File Structure
- **`dataEntry.py`**: Script for data input and management.
- **`managePaymentsCustomer.py`**: Handles customer payment tracking.
- **`managePaymentsLacing.py`**: Manages lacing payment records.
- **`reportGenerateCustomer.py`**: Generates customer data reports.
- **`reportGenerateLacing.py`**: Produces lacing activity reports.
- **`Customers.xlsx`**: Excel file containing customer information.
- **`Items.xlsx`**: Excel file detailing weaving items.
- **`Lacing.xlsx`**: Excel file with lacing details.
- **`Weaver.xlsx`**: Excel file listing weaver data.

## Contributing
Contributions are welcome! Please fork the repository and create a pull request with your enhancements.
