# Power-Azure-VM-From-Excel
This PowerShell script starts, restarts, or stops Azure VMs listed in an Excel spreadsheet.



Modules Setup 

To install the required modules to run this script, follow the instructions below.

1. Open PowerShell as Administrator
2. Run the command to install the Azure PowerShell module -
    Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force
    
3. Run the command to install the ImportExcel module- 
    Install-Module -Name ImportExcel -Scope CurrentUser
    


Azure Requirements

This script requires an Azure account to connect to, with privileges that allow administration of all virtual machines in the Excel file. Currently it can start, restart, or stop virtual machines within the same subscription that you are currently working in. 

You will be prompted to enter which operation to perform on the virtual machines in lowercase - start, restart, or stop.



Excel Requirements

This script requires the user to input the following:

1. A valid file path for an Excel file - .xlsx file extension
2. A column which lists Azure Virtual Machine names
3. The number of rows in the column - starting at row 1 - that list Azure Virtual Machine names
   
   e.g. a file with VMs listed in column A, and have values in the first two rows in column A:
   
   C:\Users\Mike\Downloads\vmfile.xlsx
   A
   2
   
   
