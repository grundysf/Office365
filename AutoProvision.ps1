function logMsg([string]$logString)
{
    write-output $logString >> C:\temp\0365_Provisioning4.log
    
    #$Global:emailBody = $Global:emailBody + "$logString`r`n"
}



function Connect {
Get-PSSession | Remove-PSSession 

$online = "account@xxx.onmicrosoft.com"
$encrypted2 = Get-Content c:\temp\encrypted.txt | ConvertTo-SecureString
$credentialLive = New-Object System.Management.Automation.PsCredential($online, $encrypted2)

#Import-PSSession $session -AllowClobber
Import-Module LyncOnlineConnector
Import-Module MSOnline
Connect-MsolService -Credential $credentialLive
}

$sessionActive = Get-PSSession | Select ComputerName

if($sessionActive.ComputerName -eq '$lynconlineserver') 
{
Write-Host "active"
}
else 
{
Connect
}

$date = Get-Date


logMsg("###Provisioned at---" + $date + "#####`n")
logMsg("############################################`n") 
logMsg("############################################`n") 
$Group = "$groupname"
       
$Groupname = (Get-ADGroup $Group)

$users = Get-ADGroupMember $Groupname | Select UserPrincipalName, employeeType, samAccountName, co

#XXX E3 AccountSku
$AccountSkuIdE3 = "XXX:ENTERPRISEPACK"
#XXX E2 AccountSku
$AccountSkuIdE2 = "XXX:MCOSTANDARD"

foreach ($user in $users){
    
       
    Get-ADUser -Identity $user.SamAccountName -Properties SamAccountName, whencreated, UserPrincipalName, co, employeeType, enabled, cn | Where-Object { ($_.whencreated -gt (get-date).adddays(-5)) -and ($_.enabled -eq $true)} | Select SamAccountName, whencreated, UserPrincipalName, co, employeeType, enabled, cn |

    foreach {

    $msol = Get-MsolUser -UserPrincipalName $_.UserPrincipalName | Select DisplayName, Licenses, isLicensed
    
    if($msol.isLicensed -eq $false) {
    Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation $_.co
    logMsg("...........................................`n")
    logMsg("User : " + $msol.DisplayName + " is not Licensed licensing...`n")
    logMsg("IsLicensed : " + $msol.IsLicensed)
    logMSg("Usage Location : " + $_.co)
    logMsg("Employee Type : " + $_.employeeType)
    logMsg("...........................................`n")

     if ($_.employeeType -eq "Employee") {
        
        $LA2 = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:ENTERPRISEPACK")} | Select ConsumedUnits, ActiveUnits

        if ($LA2.ConsumedUnits -lt $LA2.ActiveUnits) {
        $O365Licences = New-MsolLicenseOptions -AccountSkuId XXX:ENTERPRISEPACK -DisabledPlans @("INTUNE_O365", "SWAY" ,"RMS_S_ENTERPRISE", "EXCHANGE_S_ENTERPRISE", "SHAREPOINTWAC", "SHAREPOINTENTERPRISE", "YAMMER_ENTERPRISE")
       
        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE3
        Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -LicenseOptions $O365Licences
        logMsg "...........................................`n"
        logMsg("You are using " + $LA2.ConsumedUnits + " out of `n" + $LA2.ActiveUnits)
        
        } else {

        logMsg("Error" + $LA2.ActiveUnits + " available`n")


        }

                     
    }
    else {
        
      $LA = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:ENTERPRISEPACK")} | Select ConsumedUnits, ActiveUnits

      if ($LA.CosumedUnits -lt $LA.ActiveUnits)  {
      $O365Licences2 = New-MsolLicenseOptions -AccountSkuId XXX:ENTERPRISEPACK -DisabledPlans @("INTUNE_O365", "SWAY" ,"RMS_S_ENTERPRISE", "EXCHANGE_S_ENTERPRISE", "SHAREPOINTWAC", "SHAREPOINTENTERPRISE", "YAMMER_ENTERPRISE")
        logMsg("You are using " + $LA.ConsumedUnits + " out of " + $LA.ActiveUnits + " licenses assigning E2`n")
             Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE3
             Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -LicenseOptions $O365Licences2

            } else

        {
           logMsg("Error" + $LA.ActiveUnits + " available`n")
            #Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses $AccountSkuIdE2
       }
        
                
      }
    
    }
    else {
    logMsg("User : " + $msol.DisplayName + " is Already Licensed skipping...`n")
    logMsg("---------------end job----------------`n")
   
    }
    
    }
    
}

logMsg("####End Provisioning #######################`n")
logMsg("############################################`n")
logMsg("############################################`n") 
logMsg("############################################`n")
logMsg(" ")       
#############################
#MAILING REPORT HTML
#############################

 $Total = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:MCOSTANDARD")} | Select ConsumedUnits, ActiveUnits
 $Total2 = Get-MsolAccountSku | Where-Object {($_.AccountSkuId -eq "XXX:ENTERPRISEPACK")} | Select ConsumedUnits, ActiveUnits
 $E2Left = $Total.ActiveUnits - $Total.ConsumedUnits
 $E3Left = $Total2.ActiveUnits - $Total2.ConsumedUnits

function mail {

$smtpServer = "mail.server.com"
$smtpFrom = "test@test.com"
$smtpTo = "test@test.com"
$messageSubject = "Office 365 Auto-Provisioning Report"

$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true

$message.Body = $message.Body + "New Employee Provision report of Last 5 days as of " + $date  + "<br><br>"
$message.Body = $message.Body + "<h2>Employees</h2> <br><br>" + "You are using " + $Total2.ConsumedUnits + " out of `n" + $Total2.ActiveUnits + " Licenses for E3<br><br>" + "$E3Left Availabe " + "<br>" + $Global:emailBody + "<br><br>"
$message.Body = $message.Body + "<h2>Non-Employees</h2> <br><br>" + "You are using " + $Total2.ConsumedUnits + " out of `n" + $Total2.ActiveUnits + " Licenses for E3<br><br>" + "$E3Left Availabe " + "<br>" + $Global:emailBody2 + "<br><br>"

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
#$attachment = "c:\temp\0365_Provisioning4.log"
#$message.Attachments.Add( $logFile )
$smtp.Send($message)
	

}


$style = "<style>BODY{font-family: Arial; font-size: 12pt;}"
$style = $style + "b.red{color: red; }"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #008C99; padding: 5px; color: #fff; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"

$style2 = "<style>BODY{font-family: Arial; font-size: 12pt;}"
$style2 = $style2 + "b.red{color: red; }"
$style2 = $style2 + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style2 = $style2 + "TH{border: 1px solid black; background: #008C99; padding: 5px; color: #fff; }"
$style2 = $style2 + "TD{border: 1px solid black; padding: 5px; }"
$style2 = $style2 + "</style>"

function employee {

$Group = "$groupname"
       
$Groupname = (Get-ADGroup $Group)

$users = Get-ADGroupMember $Groupname | Select samAccountName

foreach ($user in $users) {


$userUPN = Get-ADUser -Identity $user.samAccountName -Properties UserPrincipalName, employeeType, Enabled, Division, samAccountName, whenCreated `
|  Where-Object { ($_.whencreated -gt (get-date).adddays(-5)) -and ($_.enabled -eq $true) -and ($_.employeeType -eq "Employee")} |Select UserPrincipalName, employeeType, Enabled, Division, samAccountName, whenCreated

foreach ($i in $userUPN) {

$result = Get-MsolUser -UserPrincipalName $i.UserPrincipalName | Select-Object DisplayName, isLicensed, SignInName, UsageLocation,`
  @{Name="Skype";Expression={$_.Licenses[0].ServiceStatus[6].ProvisioningStatus}}, `
  @{Name="ProPlus";Expression={$_.Licenses[0].ServiceStatus[5].ProvisioningStatus}},
  @{Name="Licensed";Expression={$_.isLicensed}},
  @{Name="EmployeeType";Expression={$i.employeeType}}

   
  $result | Select DisplayName, 'EmployeeType', 'Licensed', 'Skype', 'ProPlus', SignInName, UsageLocation

  } 

 }
 }



function nonemployee {

$Group2 = "$groupname"
       
$Groupname2 = (Get-ADGroup $Group2)

$users2 = Get-ADGroupMember $Groupname2 | Select samAccountName

foreach ($u in $users) {


$userUPN2 = Get-ADUser -Identity $u.samAccountName -Properties UserPrincipalName, employeeType, Enabled, Division, samAccountName, whenCreated `
|  Where-Object { ($_.whencreated -gt (get-date).adddays(-5)) -and ($_.enabled -eq $true) -and ($_.employeeType -eq "Non-Employee")} |Select UserPrincipalName, employeeType, Enabled, Division, samAccountName, whenCreated

foreach ($it in $userUPN2) {

#This command will work better if ServicePlan Number changes
#$_.Licenses[0].ServiceStatus[6].ProvisioningStatus | where-object {$_.ServicePlan.ServiceName -eq "MCOSTANDARD"}
$result2 = Get-MsolUser -UserPrincipalName $it.UserPrincipalName | Select-Object DisplayName, isLicensed, SignInName, UsageLocation,`
  @{Name="Skype";Expression={$_.Licenses[0].ServiceStatus[6].ProvisioningStatus}}, `
  @{Name="ProPlus";Expression={$_.Licenses[0].ServiceStatus[5].ProvisioningStatus}},
  @{Name="Licensed";Expression={$_.isLicensed}},
  @{Name="EmployeeType";Expression={$it.employeeType}}

   
  $result2 | Select DisplayName, 'EmployeeType', 'Licensed', 'Skype', 'ProPlus', SignInName, UsageLocation

  } 

 }


 }

$Global:emailBody = employee  | ConvertTo-Html -Head $style |Out-String
$Global:emailBody2 = nonemployee | ConvertTo-Html -Head $style2 | Out-String
mail
