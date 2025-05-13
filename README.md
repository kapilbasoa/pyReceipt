# pyReceipt
This Python program is designed to automate the generation, display, and storage of payment receipts for a business transaction. It captures essential company and customer details, transaction information, and itemized purchases, generating a formal receipt complete with taxes and total amounts. The receipt can be printed on the console and stored in an Excel workbook, either by appending to existing records or replacing old data. This solution improves business transaction management efficiency by reducing manual entry errors and enabling seamless digital record-keeping.

The program follows a structured, object-oriented approach by defining three primary classes:
**Company** — to hold company details like name, address, email, and contact number.
**Customer** — to store customer details such as name, email, phone number, and address.
**Receipt** — to manage receipt generation, including:
1. Capturing transaction data (items, prices, quantities)
2. Automatically generating unique receipt numbers using timestamps
3. Calculating totals and taxes
4. Printing the receipt in a well-formatted layout
5. Saving receipt details into an Excel workbook using the openpyxl library
The program takes customer and transaction details via console inputs, dynamically allows adding multiple purchased items, computes subtotals and tax (fixed at 7%), and then formats this information for both console display and Excel recordkeeping.
Users can choose to:
1. Append new receipt records to an existing Excel file
2. Delete and overwrite an existing file before saving new records
3. Optionally open the saved Excel file automatically after saving
4. This makes the program a practical tool for both small business transaction management and record archiving.

**Technique	Purpose**
**Object-Oriented Programming (OOP)**:	Encapsulation of related data and operations inside Company, Customer, and Receipt classes
**Date and Time Handling**:	Auto-generating unique receipt numbers and timestamps using the datetime module
**Console Input/Output (I/O)**:	Interacting with users for entering transaction details and displaying the receipt
**Dynamic Data Storage (Lists)**:	Storing purchased items and their details inside a dynamic list of dictionaries
**Iterative Loops (while loop)**:	Allowing users to enter multiple purchased items dynamically until finished
**Mathematical Operations**:	Calculating subtotals, tax, and total transaction amounts
**Excel File Handling (openpyxl)**:	Creating, reading, updating, and formatting Excel workbooks to store transaction records
**File System Operations (os module)**:  Checking for and deleting existing files when requested
**Excel Cell Formatting**:	Wrapping text and aligning cells for better readability in Excel receipts
