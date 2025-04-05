# IBL Sales Data Generator

A web application to generate sales and stock data reports from SQL Server database and export them to Excel format.

## Features

- Date range selection for data filtering
- Supplier selection with checkboxes
- Two form types:
  - Form 1: Sales Data
  - Form 2: Stock Data
- Excel export functionality
- Modern and responsive UI

## Prerequisites

- Python 3.8 or higher
- SQL Server database
- ODBC Driver 17 for SQL Server

## Installation

1. Clone the repository
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Configure the database connection:
   - Copy the `.env.example` file to `.env`
   - Update the database connection string in `.env` with your credentials

## Usage

1. Start the application:
   ```bash
   python app.py
   ```
2. Open your web browser and navigate to `http://localhost:5000`
3. Select the date range
4. Choose the form type
5. Select one or more suppliers
6. Click "Generate Excel" to download the report

## Database Setup

The application expects the following tables and functions in your SQL Server database:

1. `Suppliers` table with columns:
   - `SupplierId`
   - `SupplierName`

2. `PAP_SI_ALL` function for Form 1
3. `UDF_PAM_OPG_IBL` function for Form 2

## Error Handling

The application includes error handling for:
- Database connection issues
- Invalid date selections
- No supplier selection
- Query execution errors

## Security Notes

- Never commit the `.env` file with real credentials
- Use appropriate database user permissions
- Implement proper authentication if deploying to production 