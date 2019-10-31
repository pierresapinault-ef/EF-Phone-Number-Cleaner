# EF-Phone-Number-Cleaner
WPF application to clean phone numbers in order to upload them in Salesforce Marketing Cloud (SFMC).
The program takes an excel file from Poseidon (EF proprietary CRM system) or Salesforce and will parse it and output a CSV file ready for upload.

## How to use
1) Export your excel file from Poseidon or SF
2) Rename the first tab of the file with the market code so that the program knows which rules to use
3) Import the excel file into the application
4) The output file will be saved in the same folder as the import file
5) Upload the file to SFMC 
