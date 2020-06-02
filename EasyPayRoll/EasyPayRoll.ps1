### Enter the email address you will be sending from.
            
            
    $defaultemail = ""


### This is the beginning of the script ###

Write-Host "Checking to see if the directory exists.  If it doesn't, it will be created in C:\EasyPayroll\`n"

# Check to see if the EasyPayroll directory exists.  If not, create it.
$path = "C:\EasyPayroll"
If(!(test-path $path))
{
    New-Item -ItemType Directory -Force -Path $path
    Write-Host "Successfully created EasyPayroll folder.`n"
}
else
{
    Write-Host "Confirmed the EasyPayroll directory exists.`n"
}

# Check to see if the Payroll directory exists.  If not, create it.
$path = "C:\EasyPayroll\Payroll"
If(!(test-path $path))
{
    New-Item -ItemType Directory -Force -Path $path
    Write-Host "Successfully created Payroll folder.`n"

}
else
{
    Write-Host("Confirmed the Payroll directory exists.`n")
}

# Check to see if the Archive directory exists.  If not, create it.
$path = "C:\EasyPayroll\Archive"
If(!(test-path $path))
{
    New-Item -ItemType Directory -Force -Path $path
    Write-Host "Successfully created Archive folder. Directories successfully created`n"
    Read-Host -Prompt "Press Enter to exit."
    exit
}
else
{
    Write-Host "Confirmed the Archive directory exists.`n"

}


# If any file exists in the Payroll directory, move it to the Archive folder
try
{
  
    $finddocs = Get-ChildItem -Path "C:\EasyPayroll\Payroll\*.csv" 

 
    if($finddocs -ne $null)
    {
        $finddocs | Move-Item  -Destination "C:\EasyPayroll\Archive\$VariableName" -Force #-EA SilentlyContinue 
        Write-Host "Archiving all files in the Payroll folder...`n" 
    }
    else
    {
        Write-Host "No files found to archive.`n"
    }

}
catch
{
    Write-Host "This broke while trying to move items to the Archive folder.`n"
}

# Import and organize data
    Write-Host "Checking for contractor payroll information to import...`n"
try 
{
    $downloadpath = Get-ChildItem -Path $env:USERPROFILE -Recurse | Where-Object -FilterScript {($_.Name -eq 'Downloads')}
    $info = import-csv $downloadpath + 'EasyPayRoll - Contractors.csv' -Delimiter ','
    Write-Host "Importing contractor contact information...`n"
    Write-Host "Importing pay rates...`n"
}   
catch [System.IO.FileNotFoundException]
{ 
    Write-Host "Check to make sure the csv file is in .CSV format and was downloaded to the default download directory.`n"
    # Cancel and kill the script
    Read-Host -Prompt "Press Enter to exit"
    exit
 }

# Concatinate the dates to make one variable
$UpdatedName = $info[0]."Pay Period Start" + " - " + $info[0]."Pay Period End" 

# Rename file to $UpdatedName
Write-Host "Renaming File...`n"
Rename-Item -Path "C:\Users\pdevine\Downloads\EasyPayRoll - Contractors.csv" -NewName "$UpdatedName.csv" 

# Move file from Download directory to Payroll Directory
Write-Host "Moving file to the Payroll folder...`n"
move-item -path "C:\Users\pdevine\Downloads\$UpdatedName.csv" -destination "C:\EasyPayroll\Payroll"

# Data is imported as an array
$info = import-csv "C:\EasyPayroll\Payroll\$UpdatedName.csv" -Delimiter ',' 

# Format and Display the information going out to each team member
$Display = $info | Format-Table
Write-Host "Displaying payroll information for review...`n"
$Display 

# Pause and confirm paystubs are correct
$decision = Read-Host "Are you sure you want to proceed? [Y] for YES, [N] for NO."
if ($decision -eq 'y') 
    {
    # Continue with emailing
    Write-Host "`nPaystubs confirmed...`n" "`nPreparing to send emails..."

        # Create credential variable so user is only prompted once
        $Credentials = Get-Credential -Credential $defaultemail
        Write-Host "`nSending emails...`n"

        # Loop through the data in the spreadsheet to send the emails
        $Continue = For($i = 0; $i -lt $info.Count; $i++)            
            { 
                $From = "philip@devineandcompany.com"
                $To = $info.Email[$i]
                $Subject = $info.name[$i] + ", here is your paystub for" + " $UpdatedName" + "." 
                $Body = "Hi " + $info.name[$i] + ", here is your paystub information. Please keep this email for your records.`n" + "`n" + "Pay Period: " + $UpdatedName + "`n" + "Quantity: " + $info.Quantity[$i] + "`n" + "Pay Rate: " + $info.Rate[$i] + "`n" + "Extra: " + $info.Extra[$i] + "`n" + "Total Pay: " + $info.'Total Amount'[$i] + "`n" + "Notes: " + $info.Notes[$i] + "`n" + "`n" + "Let me know if you have any questions.  Thanks for your work! `n" + "`n" + "Best, `n" + "Philip"
                $SMTPServer = "smtp-relay.gmail.com"
                $SMTPPort = "587"
                Send-MailMessage -From $From -to $To -Subject $Subject `
                -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
                -Credential $Credentials
            }

        Read-Host -Prompt "Payroll Emails sent.  Press Enter to exit"


    }
    else 
    {
        # Cancel and kill the script
        Write-Host "Operation Cancelled. No emails were sent. `n"
        Read-Host -Prompt "Press Enter to exit"
    }