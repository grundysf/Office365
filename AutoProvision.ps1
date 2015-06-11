function logMsg($msg)
{
    write-output $msg >> C:\temp\0365_Provisioning.log
    write-host $msg
}

function mail {

$smtpServer = "outlook.xxx.com"
$smtpFrom = "xxx@xxx.com"
$smtpTo = "xxxs@xxx.com"
$messageSubject = "Office 365 Daily Provisioning Log"

$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true

$message.Body = "(View Log on \\locahost\temp\0365_Provisioning.log ) - Office 365 Provisioning report for " + $date

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
#$attachment = "c:\temp\0365_Provisioning4.log"
#$message.Attachments.Add( $attachment )
$smtp.Send($message)
	

}


function Connect {
Get-PSSession | Remove-PSSession 

$online = "xxx@xxx.onmicrosoft.com"
$encrypted2 = Get-Content c:\temp\encrypted.txt | ConvertTo-SecureString
$credentialLive = New-Object System.Management.Automation.PsCredential($online, $encrypted2)

#Import-PSSession $session -AllowClobber
Import-Module LyncOnlineConnector
Import-Module MSOnline
Connect-MsolService -Credential $credentialLive
}

$sessionActive = Get-PSSession | Select ComputerName

if($sessionActive.ComputerName -eq 'xxx.xxx.net') 
{
Write-Host "active"
}
else 
{
Connect
}

$date = Get-Date


logMsg("###Provisioned at---" + $date + "#####")
logMsg("############################################") 
logMsg("############################################") 
$Group = "xxxcom"
       
$Groupname = (Get-ADGroup $Group)

$users = Get-ADGroupMember $Groupname | Select UserPrincipalName, employeeType, samAccountName, co

#E3 AccountSku
$AccountSkuIdE3 = "xxx:ENTERPRISEPACK"
#E2 AccountSku
$AccountSkuIdE2 = "xxx:MCOSTANDARD"

foreach ($user in $users){
    
       
    Get-ADUser -Identity $user.SamAccountName -Properties SamAccountName, whencreated, UserPrincipalName, co, employeeType, enabled, cn | Where-Object { ($_.whencreated -gt (get-date).adddays(-10)) -and ($_.enabled -eq $true)} | Select SamAccountName, whencreated, UserPrincipalName, co, employeeType, enabled, cn |

    foreach {

    $msol = Get-MsolUser -UserPrincipalName $_.UserPrincipalName | Select DisplayName, Licenses, isLicensed
    
    if($msol.isLicensed -eq $false) {
    Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $_.co
    logMsg("...........................................")
    logMsg("User : " + $msol.DisplayName + " is not Licensed licensing...")
    logMsg("IsLicensed : " + $msol.IsLicensed)
    logMSg("Usage Location : " + $_.co)
    logMsg("Employee Type : " + $_.employeeType)
    logMsg("...........................................")

     if ($_.employeeType -eq "Employee") {
        
        $LA2 = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:ENTERPRISEPACK")} | Select ConsumedUnits, ActiveUnits

        if ($LA2.ConsumedUnits -lt $LA2.ActiveUnits) {
        $O365Licences = New-MsolLicenseOptions -AccountSkuId XXX:ENTERPRISEPACK -DisabledPlans @("RMS_S_ENTERPRISE", "EXCHANGE_S_ENTERPRISE", "SHAREPOINTWAC", "OFFICESUBSCRIPTION", "SHAREPOINTENTERPRISE", "YAMMER_ENTERPRISE")
       
        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE3
        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -LicenseOptions $O365Licences
        logMsg "..........................................."
        logMsg("You are using " + $LA2.ConsumedUnits + " out of " + $LA2.ActiveUnits)
        
        } else {

        logMsg("Error" + $LA2.ActiveUnits + " available")


        }

                     
    }
    else {
        
      $LA = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:MCOSTANDARD")} | Select ConsumedUnits, ActiveUnits

      if ($LA.CosumedUnits -lt $LA.ActiveUnits)  {
      $O365Licences2 = New-MsolLicenseOptions -AccountSkuId XXX:ENTERPRISEPACK -DisabledPlans @("RMS_S_ENTERPRISE", "EXCHANGE_S_ENTERPRISE", "SHAREPOINTWAC", "OFFICESUBSCRIPTION", "SHAREPOINTENTERPRISE", "YAMMER_ENTERPRISE")
        logMsg("You are using " + $LA.ConsumedUnits + " out of " + $LA.ActiveUnits + " licenses assigning E3")
             Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE3
             Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -LicenseOptions $O365Licences2

            } else

        {
           logMsg("Error" + $LA.ActiveUnits + " available")
            Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE2
       }
        
                
      }
    
    }
    else {
    logMsg("User : " + $msol.DisplayName + " is Already Licensed skipping...")
    logMsg("---------------end job----------------")
   
    }
    
    }
    
}

logMsg("####End Provisioning #######################")
logMsg("############################################")
logMsg("############################################") 
logMsg("############################################")
logMsg(" ")       
mail
