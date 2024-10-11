
[CmdletBinding()]
param(
    [String] $Group = "USOilGas",
    [Parameter(Mandatory = $true)]
    [ValidateSet(1, 2, 3, 4)]
    [int] $Runtype
)


$ScriptStartInitalizeTime = Get-Date

$FileTimeName = $ScriptStartInitalizeTime | Get-Date -Format yyyyMMddhhmm

# Log File Names
$ScriptExecutionPath = $PSScriptRoot

$LogName = "$Group-Log-$FileTimeName.log"
$UpdateLogName = "$Group-Update-$FileTimeName.log"
$recoveryJsonName = "$Group-Recovery-$FileTimeName.json"
$LogNamePath = "$ScriptExecutionPath\Logs\$LogName"
$recoveryJsonNamePath = "$ScriptExecutionPath\Recovery\$recoveryJsonName"
$UpdateLogNamePath = "$ScriptExecutionPath\Updates\$UpdateLogName"

# Un-needed, part of old code
switch ($Group) {
    'USOilGas' {
        $DeleteLogFileDays = 60
        $DeleteRecoveryFileDays = 90 
        $CommonDataURL = "" # Private Sharepoint list URL
        $SPEmployeeList = "Employees" # Employee List name
        $SPBranchesList = "Branches" # Company Site list name
        $SPLegalEntitiesList = "Legal Entities" # Company Entities List name
        $SearchBase = "" # Search Filter Active Directory OU
        # Sharepoint Column names
        $iCIMSColumn = 'iCIMS'
        $SAPIDColumn = 'SAP'
        $ApplusUPNColumn = 'ApplusGlobalUPN'
        $AxaptaIDColumn = 'Axapta'
        $ReportsToColumn = 'ReportsTo2'
        $DisplayNameColumn = 'DisplayName2'
        $JobTitleColumn = 'JobTitle'
        $StateColumn = 'State'
        $CostCenterColumn = 'CostCenter'
        $CostCenterDescriptionColumn = 'CostCenterDescription'
        $ADPPhoneNumberColumn = 'PhoneNumber'
        $ADPEmployeeMobilePhoneColumn = 'EmployeeMobilePhone'
        $VoIPPhoneNumberColumn = 'VoIPPhone'
        $ApplusMobilePhoneNumberColumn = 'ApplusMobilePhone'
        $ExcludeEmail = 'applusgca.com'
        # $DomainUsersRTDGroup = 'Domain Users RTD'
        # $USOilGasPasswordNotificationGroup = "GLB-pADS.PWDNotify-RTD"
        $AlwaysADUserGroup = @('Domain Users RTD', 'VPN_ClientAccess_USA')
        # Script Parameters
        $CompareCommonDataToAD = $true
        $CompareADToCommonData = $true
        $UpdateManager = $true
        $UpdateJobTitle = $true
        $UpdateDepartment = $true
        $UpdateAddress = $false
        $UpdateCompany = $true
        $UpdatePhoneNumbers = $true
        $DayOfWeekToNotifyHelpdesk = @("Monday" , "Tuesday", "Wednesday", "Thursday", "Friday")
        # 'BCC' Field on Sending Email
        $NotifyEmailAccounts = @("", "")
        # 'To' Field on Sending Email
        $SendEmailTo = @("")# IT help Desk Account ####
        
    }
}
# Global Column Names
$FirstNameColumn = 'FirstName'
$LastNameColumn = 'LastName'
$ApplusUPNColumn = 'ApplusGlobalUPN'
$EmployeeStatusColumn = 'EmployeeStatus'
$HireDateColumn = 'HireDate'
$CommonDataURLGlobal = ""# Private Sharepoint list URL
# Log Color
$SearchColor = "Yellow"
$PositiveResultColor = "Green"
$WarningResultColor = "Yellow"
$NegativeResultColor = "Red"
$UpdatingColor = "Cyan"



#Get Web and List, Sharepoint List Creds
$Username = ''
$Password = ''
$EncryptedPassword = convertto-securestring -String $Password -AsPlainText -Force
$Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $Username, $EncryptedPassword


$ADUsers = [System.Collections.ArrayList]::new()
$ListItems = @()
# Connects to Common Data to get the company and legal entitys
Connect-PnPOnline -Url $CommonDataURLGlobal -Credentials $Credentials
$CompanyList = Get-PnPListItem -List $SPLegalEntitiesList -Connection $CommonData | Select-Object -ExpandProperty FieldValues
# Connects to Common Data to get Employees
Connect-PnPOnline -Url $CommonDataURLGlobal -Credentials $Credentials

$LogData = [System.Collections.ArrayList]::new()
$LogData_Updates = [System.Collections.ArrayList]::new()
# Logging Functions
Function Write-Log {
    param(
        [Parameter(Mandatory = $true)][String]$msg,
        [switch]$ReturnLog
    )
    if ($ReturnLog) { Return $global:LogData -join "`n" }
    $null = $global:LogData.Add($msg)
    Add-Content $LogNamePath $msg
}
Function Write-Log_Updates {
    param(
        [Parameter(Mandatory = $true)][String]$msg,
        [switch]$ReturnLog
    )
    if ($ReturnLog) { Return $global:LogData_Updates -join "`n" }
    $null = $global:LogData_Updates.Add($msg)

    Add-Content $UpdateLogNamePath $msg
}


# Create the object for Search_Dataset
$Search_Dataset = @{
    USOilGas = @{}
    CAOilGas = @{}
}
$ADUsers = [System.Collections.ArrayList]::new()
$ListItems = @()
$DuplicateData = [System.Collections.ArrayList]::new()

# Creates the Object structure for the Search_Dataset
function get-Searchable_Data {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet("AD", "SP", "SPCompany")]
        $Source,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Column,

        $Value
    )
    # $Search_Dataset = $global:Seach_Dataset
    # $Group = $global:Group
    switch ($Group) {
        USOilGas {}
        CAOilGas {}
        Default { return }
    }
    if ($Group -notin $Search_Dataset.Keys) {
        $Search_Dataset.Add(@{$Group = @{} })
    }
    switch ($Source) {
        "AD" { $VaribleSource = $ADUsers }
        "SP" { $VaribleSource = $ListItems }
        "SPCompany" { $VaribleSource = $CompanyList }
        Default { return }
    }
    if ($Source -notin $Search_Dataset[$Group].Keys) {
        $Search_Dataset[$Group].Add($Source, @{})
        #$Search_Dataset[$Group][$Source] = @{"_Sources" = "$($Source)_Sources" }
    }
    if (($Column -notin $Search_Dataset[$Group][$Source].Keys)) {
        $Search_Dataset[$Group][$Source].Add($Column, @{})
    }
    foreach ($item in $VaribleSource) {
        $ColumnValueFlag = ($item.PSOBJECT.Properties.Name -contains $Column -and $null -ne $item.$Column)
        if ($ColumnValueFlag) {
            try {
                $Search_Dataset[$Group][$Source][$Column].Add([string](($item.$Column.Trim())), $item)
            }
            catch {
                $null = $DuplicateData.Add(@{"Column" = $Column; "Data" = $item; "Data2" = $Search_Dataset[$Group][$Source][$Column][[string]($item.$Column)] })
            }
                
        }
        else {
            $Search_Dataset[$Group][$Source][$Column][[string]($item.$Column)] = $item
        }
    }
    # }
    # catch {
    #     Write-Host "Failed to Gen Data"
    #     return
    # }
    return $Search_Dataset
}
# Refreshes the Search_Dataset
function Update-Searchable_Data {
    $global:Search_Dataset = @{
        USOilGas = @{}
        CAOilGas = @{}
    }
    switch ($Group) {
        "USOilGas" {
            $global:ListItems = (Get-PnPListItem -List $SPEmployeeList) | Select-Object -ExpandProperty FieldValues
        }
        "CAOilGas" {  
            $global:ListItems = (Get-PnPListItem -List $SPEmployeeList) | Select-Object -ExpandProperty FieldValues | Where-Object { $_.Email.length -gt 1 }
        }
        Default {}
    }
    #$global:ListItems = (Get-PnPListItem -List $SPEmployeeList) | Select-Object -ExpandProperty FieldValues #-Query "<View><Query><Where><Eq><FieldRef Name='EmployeeStatus'/><Value Type='String'>Active</Value></Eq></Where></Query></View>"
    Write-Host "Downloading Active Directory Users List from $SearchBase"
    Write-Log "Downloading Active Directory Users List from $SearchBase"
    # Removed extensionAttribute12 - Describing the user account type
    #$global:ADUsers = Get-Aduser -Filter { extensionAttribute12 -eq 'EMPLOYEE' -and (mail -notlike "*@applusgca.com") } -Properties employeeID, mail, company, department, extensionAttribute12, Title, department, Company, Manager, Office, StreetAddress, City, State, PostalCode, MemberOf, telephoneNumber, mobile, mobilePhone, OfficePhone, HomePhone  -SearchBase $SearchBase -Credential $Credentials
    $ADUsersRaw = Get-Aduser -Filter * -Properties employeeID, mail, company, department, Title, department, Company, Manager, Office, StreetAddress, City, State, PostalCode, MemberOf, telephoneNumber, mobile, mobilePhone, OfficePhone, HomePhone  -SearchBase $SearchBase -Credential $Credentials
    $TempAdUsers = [System.Collections.ArrayList]::new()
    foreach ($AduserRaw in $ADUsersRaw) {
        $AddtoADListFlag = $true
        #Mail Filter
        if ($Aduser.mail -like "*@gca.com") {
            $AddtoADListFlag = $false
        }
        if ($AddtoADListFlag -eq $true) {
            $null = $TempAdUsers.Add($AduserRaw) 
        }
    }
    $global:ADUsers = $TempAdUsers
    $global:Search_Dataset = get-Searchable_Data -Source SP -Column "Title"
    $global:Search_Dataset = get-Searchable_Data -Source SP -Column "ID"
    $global:Search_Dataset = get-Searchable_Data -Source SP -Column $SAPIDColumn
    $global:Search_Dataset = get-Searchable_Data -Source SP -Column $DisplayNameColumn
    $global:Search_Dataset = get-Searchable_Data -Source SP -Column $ApplusUPNColumn
    $global:Search_Dataset = get-Searchable_Data -Source AD -Column "EmployeeID"
    $global:Search_Dataset = get-Searchable_Data -Source SPCompany -Column "Title"
    #return @{"ListItems"=$ListItems;"ADUsers"=$ADUsers;"Search_Dataset"=$Search_Dataset}
}

Update-Searchable_Data

# $Search_Dataset[$Group]
# $Search_Dataset[$Group]['AD']
# $Search_Dataset[$Group]['SP']

# Send Email, had to workaround firewall setup Powerautomate 
function Send-Email {
    param(
        $EmailList,
        # Desides how the Email will look
        [ValidateSet("HTML", "StringTable")]
        $Emailtype,
        [ValidateSet("HelpDesk", "Testing")]
        $Recipient = "Testing",
        [Array]$Attachments,
        [Array]$CC
    )
    $tableData = $EmailList
    switch ($Emailtype) {
        "StringTable" {
            $noMatchEmployeelistString = ""
            $HashTableList = [System.Collections.ArrayList]::new()
    
            Foreach ($noMatchEmployee in $tableData) {
                $TempObj = @{}
                $noMatchEmployee.psobject.properties | ForEach-Object { $TempObj[$_.Name] = $_.Value }
                $null = $HashTableList.Add($TempObj)
            }
            Foreach ($noMatchEmployee in $HashTableList) {
                $objstring = "`n<br>"
                Foreach ($key in $noMatchEmployee.Keys) {
                    $TempString = "`n<br>{0} : {1}" -f @($key, $noMatchEmployee[[string]$key])
                    $objstring = $objstring + $TempString
                }
                $noMatchEmployeelistString = $noMatchEmployeelistString + $objstring
            }
            $body = @{
                EmailText = $noMatchEmployeelistString
                APIKEY    = ""
                EmailBody = ""
                style     = ""
            }
        }
        Default {
            $Style = @"
<style type="text/css">
    .my-table  {border-collapse:collapse;border-spacing:0;}
    .my-table td{border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;word-break:normal;}
    .my-table th{border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
    .my-table .tg-0lax{text-align:left;vertical-align:top}
</style>
"@
            # TABLE
            # Generate the HTML table with the custom CSS style
            $table = $tableData | ConvertTo-Html -As Table -Fragment #-PreContent $style 

            # Add the "my-table" class to the opening <table> tag
            $table = $table -replace '<table>', '<table class="my-table">'
            $body = @{
                EmailText = ""
                APIKEY    = ""
                EmailBody = $table -join ""
                style     = $style
            }
        }
    }
    
    switch ($Recipient) {
        "HelpDesk" {
            $body["TO"] = $SendEmailTo 
            $body["Subject"] = "$Group-Validate-ActiveDirectoryUsers"
            $body["Bcc"] = $NotifyEmailAccounts -join ";"
            $body["Group"] = $Group
        }
        "Testing" {
            $body["TO"] = "" #testing account
            $body["Subject"] = "$Group-Validate-CommonDataEmloyees_AD_Accounts"
            $body["CC"] = $CC -join ";"
            $body["Bcc"] = $NotifyEmailAccounts -join ";"
            $body["Group"] = $Group
        }
        Default {}
    }
    if ($Attachments.Count -gt 0) {
        $body["Attachments"] = $Attachments
    }
    else { $body["Attachments"] = @() }
    # Output the HTML table
    $EmailAPIHash = @{
        Uri         = ""# Path to PowerAutomate, Workaround firewall
        Headers     = @{}
        Body        = $body | ConvertTo-Json
        Method      = 'Post'
        ContentType = 'application/json'

    }
    $null = Invoke-RestMethod @EmailAPIHash 
}


function Search-AD_EmployeeID {
    param(
        # Parameter help description
        [Parameter(Mandatory = $true)]
        [PSCustomObject]
        $SharePointEmployee,
        [switch]$Manager = $null
    )
    switch ($Manager) {
        $true { $LogUser = " Manager " }
        Default { $LogUser = " " }
    }
    $Employee = $SharePointEmployee
    $iCIMSID = [string]$Employee.$iCIMSColumn
    $AxaptaID = [string]$Employee.$AxaptaIDColumn
    $SAPID = [string]$Employee.$SAPIDColumn
    $FirstName = $Employee.$FirstNameColumn
    $LastName = $Employee.$LastNameColumn
    $ADUser = $null
    If (($SAPID.length -gt 2) -and ($null -eq $ADUser)) {
        Write-Host "Searching AD for$LogUser[$($FirstName) $($LastName)] using SAP ID [$SAPID]" -ForegroundColor $SearchColor
        Write-Log "Searching for [$($FirstName) $($LastName)] using SAP ID [$SAPID]"
        $ADuser = $Search_Dataset[$Group]["AD"]["EmployeeID"][$SAPID]
    }
    If (($iCIMSID.length -gt 2) -and ($null -eq $ADUser)) {
        Write-Host "Searching AD for$LogUser[$($FirstName) $($LastName)] using iCIMS ID [$iCIMSID]" -ForegroundColor $SearchColor
        Write-Log "Searching for [$($FirstName) $($LastName)] using iCIMS ID [$iCIMSID]"
        $ADuser = $Search_Dataset[$Group]["AD"]["EmployeeID"][$iCIMSID]
        # Checks a No "i" iCIMS Number
        if ($null -eq $ADuser ) {
            $iCIMSID_temp = ($iCIMSID.ToLower()).Replace("i", $null)
            $ADuser = $Search_Dataset[$Group]["AD"]["EmployeeID"][$iCIMSID_temp]
            Write-Host "Searching AD for$LogUser[$($FirstName) $($LastName)] using iCIMS ID [$iCIMSID_temp]" -ForegroundColor $SearchColor
            Write-Log "Searching for [$($FirstName) $($LastName)] using iCIMS ID [$iCIMSID_temp]"
            $iCIMSID_temp = $Null
        }
    } 
    If (($AxaptaID.length -gt 2) -and ($null -eq $ADUser)) {
        Write-Host "Searching AD for$LogUser[$($FirstName) $($LastName)] using Axapta ID [$AxaptaID]" -ForegroundColor $SearchColor
        Write-Log "Searching for [$($FirstName) $($LastName)] using Axapta ID [$AxaptaID]"
        $ADuser = $Search_Dataset[$Group]["AD"]["EmployeeID"][$AxaptaID]
    }
    if ($null -ne $ADUser) {

    }
    return $ADuser
}


# Phone Number Cleaning Function to get a string into (+1 ### ### ####) format
function Get-CleanPhoneNumber {
    param (
        [Parameter(Mandatory = $true)]
        [String] $PhoneString = "#####"
    )
    # $PhoneString = "+1 (661) 399-8497 Ext:123"
    $PhoneString = $PhoneString.ToLower()
    # Clean Array[0] to basic number
    $PhoneString = ($PhoneString.Trim())
    # Reformate Number if has Ext in Number
    if ($PhoneString.ToLower().Contains("ext")) {
        $PhoneNumbers = (($PhoneString -split "ext")[0]).Replace("+1", $null) -replace "[^0-9]" , ''
        $PhoneExtention = " Ext" + ($PhoneString -split "ext")[1]
        $PhoneString = ("{0:$OutPattern}" -f [int64]$PhoneNumbers) + $PhoneExtention
    }
    else {
        $PhoneString = ($PhoneString).Replace("+1", $null) -replace "[^0-9]" , ''
        # Reformat Array[0] to PhoneNumber String
        $PhoneString = "{0:$OutPattern}" -f [int64]$PhoneString
    }
    return $PhoneString
}

# Spits out CSV of Latest recovery users.
function Get-RecoveryList {
    param ()
    $JsonFiles = Get-ChildItem -Path .\*.json
    $RecoveryFile = $JsonFiles | Sort-Object LastAccessTime -Descending | Select-Object -First 1
    $users = Get-Content -Path $RecoveryFile | ConvertFrom-Json 
    $users.Identity | Format-Table > quack.csv
}


###############################################################
###############################################################
###############################################################
###############################################################
# $Employee = $ListItems | Where-Object { $_.ID -eq 3022 }
# $TestEmployee = @($TestEmployee)
$UpdateActionSP = $false
$UpdateActionAD = @{"WhatIf" = $True }
$userinput = $null
$userinput = $Runtype
switch ($userinput) {
    1 { Write-Host "Continuing" }
    2 { $UpdateActionSP = $True; $UpdateActionAD = @{"WhatIf" = $false } }
    3 { Write-Host "Exiting"; exit }
    4 { Write-Host "Continuing`nNot Sending Email" }
}
$ScriptStartTime = Get-Date
# Compare CommonData to ActiveDirectory



if ($CompareCommonDataToAD) {
    $EmailList = [System.Collections.ArrayList]::new()
    $AdGroupUpdate = @{}
    $Sharepoint_Batch_List = [System.Collections.ArrayList]::new()
    $RecoveryArray = [System.Collections.ArrayList]::new()
    #$Employee_Search_error = [System.Collections.ArrayList]::new()
    # $ActiveListItems = $ListItems | Where-Object {$_.$EmployeeStatusColumn -eq "Active"}
    Foreach ($Employee in $ListItems) {
        
        #$Employee = $ListItems[0]
        $iCIMSID = ([string]$Employee.$iCIMSColumn).Replace(" ", $Null)
        $AxaptaID = ([string]$Employee.$AxaptaIDColumn).Replace(" ", $Null)
        $SAPID = ([string]$Employee.$SAPIDColumn).Replace(" ", $Null)
        $EmployeeStatus = [string]$Employee.$EmployeeStatusColumn
        $FirstName = $Employee.$FirstNameColumn
        $LastName = $Employee.$LastNameColumn
        $ADUser = $null
        $SkipFlag = $false
        switch ($Group) {
            USOilGas {
                if ($SAPID.StartsWith("77") -or ($iCIMSID.length -gt 8)) {
                    Write-Log "=== Start Validation of [$($FirstName) $($LastName)] - $($Employee.ID) ==="
                    Write-Log "[Warning]Skipping Due to Exeptions in users Employee ID's"
                    Write-Log -msg ("SAPID.Startswith(77): {0} iCIMSID<8: {1} Status: {2}" -f $SAPID, $iCIMSID, $EmployeeStatus)
                    $SkipFlag = $true
                    continue
                }  
            }
            Default {}
        }
        if ($SkipFlag) { continue }
        $CurrentIndex = $ListItems.IndexOf($Employee)

        Write-Log "=== Start Validation of [$($FirstName) $($LastName)] - SPID$($Employee.ID) -Index $CurrentIndex ==="
        Write-Host "=== Start Validation of [$($FirstName) $($LastName)] ==="
        # Find Employee AD User 
        $ADUser = Search-AD_EmployeeID -SharePointEmployee $Employee
        #  If User found based on Employee ID found in AD
        If ($null -ne $ADUser -and $EmployeeStatus -eq 'Active') {
            $SamAccountName = "$($ADuser.SamAccountName)"
            $UpdateingFlag = $false
            $UpdateAdObject = @{}
            Write-Host "ADUser found with SamAccountName [$SamAccountName]" -ForegroundColor $PositiveResultColor
            Write-Log "ADUser found with SamAccountName [$SamAccountName]"
            if ($Employee.$ApplusUPNColumn -notlike $ADUser.UserPrincipalName) {
                Write-Log "Updating $SamAccountName UserPrincipalName From:$($Employee.$ApplusUPNColumn) To:$($ADUser.UserPrincipalName)"
                Write-Log_Updates "Updating $SamAccountName UserPrincipalName From:$($Employee.$ApplusUPNColumn) To:$($ADUser.UserPrincipalName)"
                $Null = $Sharepoint_Batch_List.Add(@{"ID" = $Employee.ID; "ApplusGlobalUPN" = $ADUser.UserPrincipalName; "Values" = @{ApplusGlobalUPN = $ADUser.UserPrincipalName } })
                Write-Host "Updating SharePoint Employees List with UserPrincipalName" -ForegroundColor $UpdatingColor
                #$UpdateingFlag = $true
            }
            #Set AD.EmployeeID To Employee.SAPID
            If ($SAPID.length -gt 2) {
                if ($ADuser.employeeID -notlike $SAPID) {
                    #Write-Host "Updating $SamAccountName employeeID to $SAPID from $($ADuser.employeeID)" -ForegroundColor $UpdatingColor
                    Write-Log "Updating $SamAccountName employeeID to $SAPID"
                    #Set-Aduser -Identity $SamAccountName -Replace @{employeeID=$SAPID} -Credential $Credentials
                    $UpdateingFlag = $true
                    $UpdateAdObject["employeeID"] = $SAPID
                }
            }
            #Set AD.Title To $Employee.$JobTitleColumn
            if ($UpdateJobTitle) {
                $JobTitle = $Employee.$JobTitleColumn
                If ($JobTitle -gt 2) {
                    if ($ADuser.Title -notlike $JobTitle) {
                        Write-Log "Updating $SamAccountName Title to $JobTitle"
                        $UpdateingFlag = $true
                        $UpdateAdObject["title"] = $JobTitle
                    }
                }
            }
            #Set $ADUser.Department To $Employee.$CostCenterDescriptionColumn
            if ($UpdateDepartment) {
                $Department = $Employee.$CostCenterDescriptionColumn
                If ($Department -gt 2) {
                    if ($ADuser.Department -notlike $Department) {
                        #Write-Host "Updating $SamAccountName Department to $Department from $(if($null -eq $ADUser.Department){"null"}else{$ADUser.Department})" -ForegroundColor $UpdatingColor
                        Write-Log "Updating $SamAccountName Department to $Department"
                        #Set-Aduser -Identity $SamAccountName -Replace @{department=$Department} -Credential $Credentials
                        $UpdateingFlag = $true
                        $UpdateAdObject["department"] = $Department
                    }
                }
            }
            #Set $ADUser.Company To [String]($Search_Dataset[$Group]["SPCompany"]["Title"][[String]($SAPCompanyCode)].LegalEntityName)
            if ($UpdateCompany) {
                $SAPCompanyCode = [String]($Employee.$CostCenterColumn.substring(0, 4))
                $LegalEntityName = [String]($Search_Dataset[$Group]["SPCompany"]["Title"][[String]($SAPCompanyCode)].LegalEntityName)
                if ($LegalEntityName.Length -ge 2 -and $LegalEntityName -notlike $ADUser.Company) {
                    #Write-Host "Updating $SamAccountName Company to $LegalEntityName from $(if($null -eq $ADUser.Company){"null"}else{$ADUser.Company})" -ForegroundColor $UpdatingColor
                    Write-Log "Updating $SamAccountName Company to $LegalEntityName"
                    #Set-Aduser -Identity $SamAccountName -Company $LegalEntityName -Credential $Credentials
                    $UpdateingFlag = $true
                    $UpdateAdObject["Company"] = $LegalEntityName
                }
            }
            # Update Manager
            if ($UpdateManager) {
                $ManagerADUser = $null
                $ManagerSamAccountName = $Null
                $ManagerName = $Employee.$ReportsToColumn
                if ($ManagerName.Length -ge 2) {
                    Write-Host "Searching for Manager [$ManagerName] using Dataset" -ForegroundColor $SearchColor
                    $Manager = $Search_Dataset[$Group]['SP']["$DisplayNameColumn"]["$ManagerName"]
                    If ($null -ne $Manager) {
                        $ManagerADUser = $null
                        Write-Host "Manager Sharepoint found with $DisplayNameColumn [$ManagerName]" -ForegroundColor $PositiveResultColor
                        Write-Log "Manager Sharepoint found with $DisplayNameColumn [$ManagerName]" 
                        # Find Manager AD User
                        $ManagerADuser = Search-AD_EmployeeID -SharePointEmployee $Manager -Manager
                    }
                    # Set $SamAccountName.Manager to  $ManagerADUser.SamAccountName
                    If ($null -ne $ManagerADUser) {
                        $ManagerSamAccountName = $ManagerADUser.SamAccountName
                        Write-Host "Manager ADUser found with SamAccountName [$ManagerSamAccountName]" -ForegroundColor $PositiveResultColor
                        Write-Log "Manager ADUser found with SamAccountName [$ManagerSamAccountName]"
                        #Write-Host "Updating $SamAccountName Manager to $ManagerName" -ForegroundColor $UpdatingColor
                        #Set-Aduser -Identity $SamAccountName -Manager $ManagerSamAccountName -Credential $Credentials
                        if ($Aduser.Manager -ne $ManagerADuser.DistinguishedName) {
                            $UpdateingFlag = $true
                            $UpdateAdObject["Manager"] = $ManagerADUser.DistinguishedName
                        }
                    }
                    else {
                        Write-Host "Manager [$ManagerName] not found with SamAccountName [$ManagerSamAccountName]" -ForegroundColor $NegativeResultColor
                    }
                }
            }
            # Update Address
            if ($UpdateAddress) {
                $Branch = $null
                $BranchName = $null
                $BranchAddress = $null
                $BranchCity = $null
                $BranchState = $null
                $BranchPostalCode = $null
                $HRBranch = $Employee.$EmployeeLocationColumn
                $AddressID = (Get-PnPListItem -List $SPBranchesList -Query "<View><ViewFields><FieldRef Name='$BranchesHRLocationColumn'/></ViewFields><Query><Where><Eq><FieldRef Name='$BranchesHRLocationColumn'/><Value Type='Text'>$HRBranch</Value></Eq></Where></Query></View>").Id
                if ($AddressID -gt 0) {
                    $Branch = (Get-PnPListItem -List $SPBranchesList -ID $AddressID).FieldValues
                    $BranchName = $Branch.Title
                    Write-Host "Branch Identified as [$BranchName]" -ForegroundColor $PositiveResultColor
                    Write-Log "Branch Identified as [$BranchName]"
                    $BranchAddress = $Branch.Address
                    $BranchCity = $Branch.City
                    $BranchState = $Branch.$StateColumn
                    $BranchPostalCode = $Branch.PostalCode

                    if ($BranchAddress.Contains("`n")) {
                        $BranchAddress = $BranchAddress -replace "`n", "`r`n"
                    }
                    Write-Host "Updating $SamAccountName Branch Information to [$BranchName, $BranchAddress, $BranchCity, $BranchState, $BranchPostalCode]" -ForegroundColor $UpdatingColor
                    Write-Log "Updating $SamAccountName Branch Information to [$BranchName, $BranchAddress, $BranchCity, $BranchState, $BranchPostalCode]"
                    #Set-ADUser -Identity $SamAccountName -Office $BranchName -StreetAddress "$BranchAddress" -City "$BranchCity" -State "$BranchState" -PostalCode "$BranchPostalCode" -Credential $Credentials
                }
            }
            if ($UpdatePhoneNumbers) {
                <#
                Contact Card|1st|2nd|3rd
                WorkPhone|8x8|Verizon|ADP
                Mobile|Verizon|ADP
                Home|ADP

                |Source|Sharepoint|AD|ContactCard|
                |8x8|VoIPPhone|telephoneNumber/WorkPhone|WorkPhone|
                |ADP|PhoneNumber|mobile/mobilePhone|Mobile|
                |Verizon|ApplusMobilePhone|Home/HomePhoneNumber|Home|
                #>
                $OutPattern = '+1 ### ### ####'
                $VoIPPhoneNumberColumn = 'VoIPPhone'
                $ApplusMobilePhoneNumberColumn = 'ApplusMobilePhone'
                $ADPPhoneNumberColumn = 'PhoneNumber'
                $ADPEmployeeMobilePhoneColumn = 'EmployeeMobilePhone'
                $PhoneAdProperties = @("telephoneNumber", "mobile", "HomePhone")
                
                $PhoneNumberArray = [System.Collections.ArrayList]@(
                    $Employee.$VoIPPhoneNumberColumn,
                    $Employee.$ApplusMobilePhoneNumberColumn,
                    $Employee.$ADPPhoneNumberColumn,
                    $Employee.$ADPEmployeeMobilePhoneColumn
                )
                # Clean PhoneNumberArray of Null Values
                # Clean PhoneNumberArray of Null Values
                $PhoneNumberArray.Remove($null)
                $PhoneNumberArray.Remove('')
                # Format Phone Array for Cleaning and Unique Numbers
                if ($PhoneNumberArray.Count -gt 0) {
                    foreach ($SPPhoneNumber in $PhoneNumberArray.Clone()) {
                        $SPNumIndx = $PhoneNumberArray.IndexOf($SPPhoneNumber)
                        if ($Null -eq $SPPhoneNumber) {
                            $PhoneNumberArray.Remove($SPPhoneNumber)
                        }
                        else {
                            $PhoneNumberArray[$SPNumIndx] = Get-CleanPhoneNumber -PhoneString $SPPhoneNumber
                        }
                    }
                    switch (($PhoneNumberArray | Select-Object -Unique).count) {
                        0 {}
                        1 { 
                            $PhoneNumberArrayClone = ($PhoneNumberArray | Select-Object -Unique)
                            $PhoneNumberArray.Clear()
                            $null = $PhoneNumberArray.Add($PhoneNumberArrayClone)
                        }
                        Default { [System.Collections.ArrayList]$PhoneNumberArray = ($PhoneNumberArray | Select-Object -Unique) }
                    }
                }
                
                try {
                    # $key = "telephoneNumber"
                    # $key = "mobile"
                    # $key = "HomePhone"
                    foreach ($key in $PhoneAdProperties) {
                        if ($PhoneNumberArray.Count -eq 0) { break }
                        $PhoneNumSetFlag = $false
                        do {
                            # Check Input Number Array to see if there are any left to set
                            if ($PhoneNumberArray.Count -ge 1) {
                                # Check Array[0] for proper Number
                                if ($PhoneNumberArray[0].length -ge 10) {
                                    # Add Array[0] value to Current AD Phone Number
                                    if ($ADuser.$key -notlike $PhoneNumberArray[0]) {
                                        if ($null -ne $ADuser.$key) {
                                            $ADUnique = (Get-CleanPhoneNumber ($ADuser.$key))
                                            if (!$PhoneNumberArray.Contains($ADUnique)) {
                                                try {
                                                    $null = $PhoneNumberArray.Add($ADUnique)
                                                }
                                                catch {
                                                    Write-Host "Error Catch adding Unique number"
                                                    Write-Host $Error[0]
                                                    exit
                                                    <#Do this if a terminating exception happens#>
                                                }
                                            }
                                        }
                                        Write-Log "Updating $SamAccountName New $key from $($ADuser[$key]) to $($PhoneNumberArray[0])"
                                        Write-Host "Updating $SamAccountName New $key from $($ADUser[$key]) to $($PhoneNumberArray[0])"
                                        $UpdateingFlag = $true
                                        $UpdateAdObject[$key] = $PhoneNumberArray[0]
                                    }
                                    
                                    # Remove Array[0] so it wont be resused
                                    $PhoneNumberArray.RemoveAt(0)
                                    # Set Flag to end Do-While
                                    $PhoneNumSetFlag = $true
                                }
                                else { $PhoneNumberArray.RemoveAt(0) }
                            }
                            else { $PhoneNumSetFlag = $true }
                        } while ($PhoneNumSetFlag -eq $false)
                    }
                }
                catch {
                    Write-Host $Error[0]
                    Write-Host "Flag = $PhoneNumSetFlag"
                    Write-Host "PhoneNumberArray = $PhoneNumberArray"
                    Write-Host "key = $key"
                    Write-Host "Catch Error on Phone Main Loop/Logic"
                    exit
                }
            }
            
            # Update AD User with Updated values
            if ($UpdateingFlag) {
                [Hashtable]$RecoveryObject = @{"Identity" = $SamAccountName; "PropertiesN" = $UpdateAdObject; "PropertiesO" = @{} }
                foreach ($key in $UpdateAdObject.Keys) {
                    $RecoveryObject["PropertiesO"].Add([string]$Key , $ADUser.$Key) 
                    switch ($key) {
                        Manager { 
                            $ADManagerValue = $ADUser.Manager
                            if ($null -ne $ADManagerValue) {
                                $ADManagerValue = $((($ADUser["Manager"]).split(",").split("=")[1]).Replace("\", $Null))
                            }
                            
                            Write-Host "Updating $SamAccountName $key from $ADManagerValue to $($Manager.$DisplayNameColumn)" -ForegroundColor Green;
                            Write-Log_Updates -msg "Updating $SamAccountName $key from $ADManagerValue to $($Manager.$DisplayNameColumn)" 
                        }
                        Default { 
                            Write-Host "Updating $SamAccountName $key from $($ADUser[$key]) to $($UpdateAdObject.$key)" -ForegroundColor Green 
                            Write-Log_Updates -msg "Updating $SamAccountName $key from $($ADUser[$key]) to $($UpdateAdObject.$key)"
                        }
                    }
                }
                #Set-Aduser -Identity $SamAccountName -Replace $UpdateAdObject -Credential $ADCredentials @UpdateActionAD
                Set-Aduser -Identity $SamAccountName -Replace $UpdateAdObject -Credential $Credentials @UpdateActionAD
                $null = $RecoveryArray.Add($RecoveryObject)
            }
            else {
                Write-Log -msg "User Matches Sharepoint"
            }
            # Add AD User to List of AD Groups the user is missing
            $Memberof = $ADUser.MemberOf | ForEach-Object { $_.split(",").split("=")[1] }
            Foreach ($AdGroup in $AlwaysADUserGroup) {
                #Write-Host "$AdGroup"
                if ($AdGroup -notin $Memberof) {
                    if ($AdGroup -notin $AdGroupUpdate.Keys) {
                        $AdGroupUpdate["$AdGroup"] = [System.Collections.ArrayList]::new()
                    }
                    Write-Log -msg "$($ADUser.SamAccountName) Added to $AdGroup"
                    Write-Log_Updates -msg "$($ADUser.SamAccountName) Added to $AdGroup"
                    $null = $AdGroupUpdate[$AdGroup].Add($ADUser.SamAccountName)
                }
            }
        }
        else {
            # If No User found based on Employee ID found in AD
            If ($EmployeeStatus -eq 'Active') {
                # $CurrentDate -ge $employee.$HireDateColumn
                # if (!($employee.$HireDateColumn)) {
                #     $HireFlag = $False
                # }
                # elseif ($CurrentDate -ge $employee.$HireDateColumn) {
                #     $HireFlag = $true
                # }
                # Else { $HireFlag = $False }
                $null = $EmailList.Add(
                    [pscustomobject]@{
                        Index            = $Employee.ID
                        FirstName        = "$($FirstName)"
                        LastName         = "$($LastName)"
                        ID               = $Employee.Title
                        UPN              = $Employee.$ApplusUPNColumn
                        HireDate         = if ($employee.$HireDateColumn.length -gt 1) { ($employee.$HireDateColumn | Get-Date -Format "MM/dd/yyyy") };
                        CommonDataStatus = $EmployeeStatus
                    }
                )
                Write-Host "Added To Email List" -ForegroundColor DarkBlue
                Write-Log "Added To Email List"
            }
            else {
                Write-Host "User Terminated" -ForegroundColor Magenta
                Write-Log "User Terminated"
            }
        }
        Write-Log "=== End Validation of [$($FirstName) $($LastName)] ==="
        Write-Host "=== End Validation of [$($FirstName) $($LastName)] ==="
        #$EmailList.Count
    }
    
    # Update AD User's AD Groups
    Foreach ($AdGroup in $AdGroupUpdate.Keys) {
        $key = $ADGroup
        $ADGroup = $AdGroupUpdate[$key]
        Add-ADGroupMember -Identity $key -Members $ADGroup -Credential $Credentials @UpdateActionAD
    }
}
$ScriptEndTime = Get-Date
switch ($runtype) {
    2 { if ($RecoveryArray.Count -ge 1) { $RecoveryArray | ConvertTo-Json > $recoveryJsonNamePath } }
    4 { 
        "$(($ScriptEndTime - $ScriptStartInitalizeTime).Minutes) : $(($ScriptEndTime - $ScriptStartInitalizeTime).Seconds)"
        "$(($ScriptEndTime - $ScriptStartTime).Minutes) : $(($ScriptEndTime - $ScriptStartTime).Seconds)"
        if ($RecoveryArray.Count -ge 1) { $RecoveryArray | ConvertTo-Json > $recoveryJsonNamePath }
        exit 
    }
} 

# Runtimes,varible methods
####100###
# 0 : 10
# 0 : 2
# # # # #
# 0 : 14
# 0 : 6
####100###
###All###
# 1 : 26
# 1 : 19
# # # # #
# 3 : 39
# 3 : 31
###All###
# 0 : 13
# 0 : 5


Write-Host "Done"
Write-host @" 
    Emails to Sent: $($EmailList.Count)
    Total Manager Error: $($ManagerError.count)
    Total Batch List: $($Sharepoint_Batch_List.Count)
"@

# Email Attachment for CD employees with no AD account.
$EmailAttachments = [System.Collections.ArrayList]::new()


# Update SharePoint
$Batch = New-PnPBatch
$SharepointBatchSize = 10
$Sharepoint_Batch_List_Invoke = [System.Collections.ArrayList]::new()
foreach ($item in $Sharepoint_Batch_List) {
    $CurrentUpn = $Null
    $CurrentEmployee = ($Search_Dataset[$Group]['SP']['ID'][[string]($item.ID)])
    $CurrentUpn = $CurrentEmployee.ApplusGlobalUPN
    if ($CurrentUpn -ne $item.Values.ApplusGlobalUPN) {
        #Set-PnPListItem -List $SPEmployeeList -Identity $Item.ID -Values $Item.Values -Batch $Batch
        $null = $Sharepoint_Batch_List_Invoke.Add(@{"ID" = $Item.ID; "ItemUPN" = $Item.Values.ApplusGlobalUPN; "SearchUPN" = $CurrentUpn; "Verify" = $CurrentUpn })
        Write-Host "Updating:$($item.ApplusGlobalUPN)"
        foreach ($key in $item.Values.Keys) {
            Write-Host "From:$($CurrentEmployee.$key) To:$($item.Values.$key)"
        }
        Set-PnPListItem -List $SPEmployeeList -Identity $Item.ID -Values $Item.Values -Batch $Batch
    }
    if ($Batch.RequestCount -ge $SharepointBatchSize) {
        Write-Host "Invoke Batch IN"
        $Sharepoint_Batch_List_Invoke = [System.Collections.ArrayList]::new()
        if ($UpdateActionSP) {
            Invoke-PnPBatch -Batch $batch
        }
        continue
    }
}
if ($Sharepoint_Batch_List_Invoke) {
    Write-Host "Invoke Batch Out"
    $Sharepoint_Batch_List_Invoke = [System.Collections.ArrayList]::new()
    if ($UpdateActionSP) {
        Invoke-PnPBatch -Batch $batch
    }
    # Set-PnPListItem -List $SPEmployeeList -Identity $Item.ID -Values $Item.Values -Batch $Batch
}



# Reload the Search_Dataset
Update-Searchable_Data
# Compare Active Directory to CommonData
if ($CompareADToCommonData) {
    $UnmatchedADUsers = [System.Collections.ArrayList]::new()
    foreach ($ADUser in $ADUsers) {
        $ADUserUPN = $null
        $EmployeeRecordID = $null
        $ADUserUPN = $ADUser.UserPrincipalName
        
        if ($ADUserUPN) {
            $EmployeeRecordID = $Null
            $EmployeeRecordEmployeeID = $null
            # Search Common Data For the Current AD User using $ADUser.UserPrincipalName
            $EmployeeRecordID = $Search_Dataset[$Group]["SP"][$ApplusUPNColumn]["$($ADUser.UserPrincipalName)"]
            # Search Common Data For the Current AD User using $ADUser.employeeID
            $EmployeeRecordEmployeeID = $Search_Dataset[$Group]["SP"]["Title"]["$($ADUser.employeeID)"]
            if ($EmployeeRecordID -or $EmployeeRecordEmployeeID) {
                Write-Host "Employee Record for ADUser $ADUserUPN Found" -ForegroundColor $PositiveResultColor
                Write-Log "Employee Record for ADUser $ADUserUPN Found"
            }
            else {
                if ($ADUser.mail -like "*@$ExcludeEmail") {
                    Write-Host "Employee Record for ADUser $ADUserUPN is externally managed" -ForegroundColor $PositiveResultColor
                    Write-Log "Employee Record for ADUser $ADUserUPN is externally managed"
                }
                else {
                    Write-Host "[WARNING] Employee Record for ADUser $ADUserUPN Not Found" -ForegroundColor $WarningResultColor
                    Write-Log "[WARNING] Employee Record for ADUser $ADUserUPN Not Found"
                    $null = $UnmatchedADUsers.Add(
                            ($ADUser | Select-Object -Property SamAccountName, UserPrincipalName, Enabled, employeeID, Office, DistinguishedName) 
                    )
                }
            }
        }       
    }
    $EmailAttachments_Temp = $UnmatchedADUsers | ConvertTo-Csv -NoTypeInformation
    $EmailAttachments_Temp = $EmailAttachments_Temp -join [Environment]::NewLine
    $null = $EmailAttachments.Add(
        @{
            "Name"         = "AD_NoMatch_Commondata.csv"
            "ContentBytes" = [convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($EmailAttachments_Temp))
        }
    )
    $UnmatchedADUsers | ConvertTo-Json  | Out-File -FilePath "$($env:USERPROFILE)\Downloads\AD_To_CommonData_UPN.json" -Encoding ASCII

}



$EmailListAttachment = $EmailList | ConvertTo-Csv -NoTypeInformation
$EmailListAttachment = $EmailListAttachment -join [Environment]::NewLine

if ($EmailListAttachment) {
    $null = $EmailAttachments.Add(
        @{
            "Name"         = "CommonData_TO_AD.csv"
            "ContentBytes" = [convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($EmailListAttachment))
        }
    )
}
if ((Write-log -ReturnLog -msg 1)) {
    $null = $EmailAttachments.Add(@{
            "Name"         = "Log.txt"
            "ContentBytes" = [convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((Write-log -ReturnLog -msg 1)))
        })
}
if ((Write-Log_Updates -ReturnLog -msg 1)) {
    $null = $EmailAttachments.Add(@{
            "Name"         = "UpdatesMade.txt"
            "ContentBytes" = [convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes((Write-Log_Updates -ReturnLog -msg 1)))
        })
}


# Check Logs to see if its run in the last 24 Hours.
# Gets the 2nd Log and not the current Log 
$LastCreatedLog = (Get-ChildItem -Path $ScriptExecutionPath -Exclude @("*.ps1", $LogName, $UpdateLogName, $recoveryJsonName) | Sort-Object -Property LastWriteTime -Descending)
# Check if theres 2 logs
switch ($LastCreatedLog.count) {
    0 { $LastCreatedLog = @{"CreationTime" = ((Get-Date).AddHours( - (24))) } }
    Default { $LastCreatedLog = $LastCreatedLog[0] }
}

if (((Get-date).DayOfWeek) -in $DayOfWeekToNotifyHelpdesk -and ($EmailList.Count -ge 1) -and (((Get-Date).AddHours(-12)) -gt $LastCreatedLog.CreationTime)) {
    Write-Log -msg "Email Sent"
    Write-Host "Email Sent" -ForegroundColor $PositiveResultColor
    switch ($Runtype) {
        1 {
            Send-Email -EmailList $EmailList -Emailtype HTML -Recipient Testing -Attachments $EmailAttachments
        }
        2 {
            #Send-Email -EmailList $EmailList -Emailtype StringTable -Recipient HelpDesk -Attachments $EmailAttachments
            Send-Email -EmailList $EmailList -Emailtype HTML -Recipient Testing -Attachments $EmailAttachments
        }
        Default {}
    }
}
else {

    $LogMsg = @"
    No Email Sent
Day of Week : $(((Get-date).DayOfWeek) -in $DayOfWeekToNotifyHelpdesk)
EmailList : $(($EmailList.Count -ge 1))
LastWrittenLog : $(((Get-Date).AddHours(-12)) -gt $LastCreatedLog.CreationTime)
DayOfWeek = $(((Get-date).DayOfWeek))
EmailList = $($EmailList.Count)
LastWrittenLog = $($LastCreatedLog.CreationTime)
"@
    Write-Host $LogMsg
    Write-Log -msg $LogMsg
}


# General Cleaning of Files
# $DeleteLogFileDays = -1
# $DeleteRecoveryFileDays = -1
# Set Dates of if lastWrite le DeleteDate remove
$DeleteLogFileDate = (Get-date).AddDays(-$DeleteLogFileDays)
$DeleteRecoveryFile = (Get-date).AddDays(-$DeleteRecoveryFileDays) 
$RemoveFileBool = -not $UpdateActionAD.Values

# Gets all files excluding log,update,recovery and ps1 files sort by last Write Time
$FilesForCleaning = (Get-ChildItem -Recurse -Path $ScriptExecutionPath -Exclude @("EmployeeFix", "*.ps1", $LogName, $UpdateLogName, $recoveryJsonName)) | Sort-Object LastWriteTime -Descending
# Clean Log Files
$FilesForCleaning | Where-Object { $_.Name -like "$group-Log-*.log" -and $_.LastWriteTime -le $DeleteLogFileDate } | ForEach-Object {
    Write-Log "Removing=$RemoveFileBool : $_"
    Remove-Item $_ @UpdateActionAD
}
$FilesForCleaning | Where-Object { $_.Name -like "$group-Update-*.log" -and $_.LastWriteTime -le $DeleteLogFileDate } | ForEach-Object {
    Write-Log "Removing=$RemoveFileBool : $_"
    Remove-Item $_ @UpdateActionAD
}
# Clean Recovery Files
$FilesForCleaning | Where-Object { $_.Name -like "$group-Recovery-*.json" -and $_.LastWriteTime -le $DeleteRecoveryFile } | ForEach-Object {
    Write-Log "Removing=$RemoveFileBool : $_"
    Remove-Item $_ @UpdateActionAD
}

exit

