# This script will start, restart, or stop Azure virtual machines listed in an Excel file, based on user selection

Import-Module -Name Az
Import-Module -Name ImportExcel


#Connect to an Azure Account
try{
      Connect-AzAccount
  }
  catch{
      Write-Host "Could not connect to an Azure account."
      exit
  }

#Wait for Azure account connection to process
Start-Sleep -Second 8

#Input the Excel file path
while ($true){
    $excelPath = Read-Host "Enter the path of the Excel file with Azure VMs listed "
    Write-Host "`nConfirmed file path is $excelPath"

    #Check if file path exists on local machine, and is of type .xlsx
    $isPath = Test-Path -Path $excelPath
    if ($isPath -and $excelPath -like "*.xlsx"){
        break
    }
    else{
        Write-Host "Please enter a valid Excel file path"
    }
}

#Input the Excel column to process
$column = Read-Host "`nEnter the column to process "
Write-Host "Confirmed column to process is $column"

#Input the number of Excel sheet rows to iterate through
$rowCount = Read-Host "`nEnter the number of rows to process "
Write-Host "Confirmed number VMs to process is $rowCount"


#Select VM Operation - start, restart, stop
$operationList = 'start','restart','stop'
while ($true){
    $vmOperation = Read-Host "`nWould you like to start, restart, or stop the VMs?`nEnter start, restart, or stop "
    Write-Host "`nConfirmed operation is to $vmOperation the VMs`n"

    if($operationList -contains $vmOperation ){
        break
    }
    else{
        Write-Host "Please enter one of the following VM operations - start, restart, stop"
    }
}

#Print contents of Excel file
Import-Excel -Path $excelPath

#Open Excel workbook, sheet
$excelBook = Open-ExcelPackage -Path $excelPath
$excelSheet = $excelBook.Workbook.Worksheets['Sheet1']

#Iterate through each row of Excel column previously selected
for($i = 1; $i -le $rowCount; $i++){
    $cell = $column + $i
    Write-Host "Processing $cell ..."
    $cellValue = $excelSheet.Cells[$cell].Value

    Write-Host "`nPulling VM information...`n"
    
    #Check if VM exists in connected Azure subscription
    try{
        $vm = Get-AzVM -Name $cellValue
        $vmStartStateHash = Get-AzVM -Name $cellValue -Status | Select-Object powerstate
        $vmStartStateValue = $vmStartStateHash.PowerState

        Write-Host "VM Name: "  $vm.Name
        Write-Host "Resource Group: "  $vm.ResourceGroupName
        Write-Host "OS: "  $vm.StorageProfile.ImageReference.Offer "- " $vm.StorageProfile.ImageReference.Sku 
        Write-Host "Status: " $vmStartStateValue
        Write-Host "Tags: "  $vm.Tags
    }
    catch{
        Write-Host "`nCould not find $cellValue in this Azure environment"
    }

    Write-Host "`n`nPerforming $vmOperation on $cellValue ..."
    
    #Perform selected VM operation
    try{
        if($vmOperation -eq "start"){
            Start-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
        }
        elseif ($vmOperation -eq "restart"){
            Restart-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
        }
        else{
            Stop-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
        }
    }
    catch{
        Write-Host "`nCould not $vmOperation this VM"
    }

    #Add segment lines between operations
    Write-Host "`n----------------------------------------------------------------------`n"
}


#Save Excel file changes 
Close-ExcelPackage $excelBook


