# Stock-buy-email-report C# Program Summary

This C# program is designed to perform several tasks related to managing product stock, calculating margins, creating a PDF report, and sending an email with the report attached. 

## Purpose

The program is created to assist with inventory management by:

- Connecting to a database to retrieve product information.
- Calculating profit margins based on sales and costs.
- Generating a PDF report with sales data and margins.
- Sending an email with the PDF report attached.

## Functionalities

### Database Connection

- Establishes a connection to the specified SQL Server database.

### Product Data Retrieval

- Retrieves product information from the database, including:
  - Stock quantities.
  - Product IDs.
  - Descriptions.
  - Family categories.
  - Sale prices.
  - Other relevant details.

### XML Data Processing

- Reads XML data from a specific URL.
- Deserializes the XML into C# objects (using `XmlSerializer`).

### Margin Calculation

- Calculates profit margins for each product based on:
  - Sale prices.
  - Cost prices (retrieved from the database).
  - Margins are calculated in both percentage and Euro values.

### PDF Report Generation

- Uses the iText library to create a PDF report.
- The report includes:
  - Header with company information.
  - Details of products:
    - Product ID.
    - Description.
    - Stock quantities.
    - Units sold in the last 30 days.
    - Profit margins in percentage and Euro values.
  - Data is organized by product family.

### Email Notification

- Sends an email with the generated PDF report attached.
- Uses the SMTP protocol for email transmission.
- The email includes:
  - Recipient addresses (specified in the code).
  - Subject with the report's date.
  - Body with additional information and credits.

## Dependencies

- iText library: Used for PDF generation.
- .NET `System.Data.SqlClient`: For SQL Server database connectivity.
- .NET `System.Net`: For web requests and email transmission.
- `System.Xml.Serialization`: For XML data deserialization.
- `System.Linq`: For LINQ queries.

## Notes

- The program assumes a specific database structure and XML format.
- File paths for the PDF report and SMTP settings are hardcoded.
- Margins are calculated based on specific formulae for cost and sale prices.

### Contributors

- Developed by Samuel Parente.


