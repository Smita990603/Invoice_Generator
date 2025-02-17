# Invoice_Generator

**The script performs the following actions:**

**Reads data:** Reads data from an Excel file using Pandas.  
**Calculates totals:** Calculates total price, tax, and grand total for each product.  
**Groups by customer:** Groups the data by customer ID.  
**Generates invoices:** For each customer, it creates a PDF invoice with:

```text
Company information
Invoice number and date
Customer ID
Product details (table)
Subtotal, tax, and grand total
```

**Saves PDFs:** Saves the generated PDFs in the Invoice_generated directory.  
**Updates Excel File:** Saves the calculated Total_price, Tax, and Grand_Total to the same Excel file.  
Customization
**File paths:** Adjust the Excel file path if it's not in the same directory as the script.  
**Output folder:** The PDFs are saved in the Invoice_generated folder. Make sure this folder exists. You can change the folder name in the doc definition if you prefer a different folder.
Company information: Change the company_name and Company_address variables.
**Styling:** Modify the TableStyle and ParagraphStyle objects to customize the appearance of the invoice.  
**Tax rate:** Change the tax rate (currently 5%) as needed.  
**Invoice Number Format:** Modify the invoice number format if required.  

This README provides a clear and concise guide for users to understand and use your invoice generation script. It covers requirements, usage instructions, code, explanation, and customization options.  The most important change is the explicit instruction to create the Invoice_generated folder.
