# Quicken_Sankey
D3-based Sankey diagram for Quicken transactions

Uses XLSX from Quicken Transactions as described below. The input reader is very specific to Quicken's output format, especially how category hierachies are represented as colon-delimited string in a Category column. Only uses Category and Amount fields (which are also hard-coded). The expected file name is also hard coded as quicken_export.xlsx. No attempt to generalize the input reader.

1. To create the required XLSX in Quicken, create a custom Transaction report. I use the following options for my output:
- Date Range (top): Year to Date
- Sort by: Category
- Columns: Date, Account, Category, Amount. The reader doesn't use Date and Account but I include them for debugging.
- Accounts: No Investment accounts, only Banking accounts
- Categories: All
- Payees: All
- Securities: Clear all (no securities included)
- Advanced: Transfer: Exclude all ; Subcategories: Show all

2. Run the report. Output should look like this:
INSERT FIGURE HERE

3. Export the report as XLSX and save in the same directory as the index.html file. Must save the XLSX as quicken_export.xlsx as this filename is hard coded.

# WARNING: Quicken add additional header and footer rows for titles and summary data to the XLSX that must be MANUALLY removed. The final spreadsheet should only have one workbook, Row 1 as headers, Rows2-N as single transactions. 

4. Save the modified XLSX.
5. Start a local server to readh index.html. I use VSC with the Live Server plug-in.


