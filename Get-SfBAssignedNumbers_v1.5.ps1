#region INFO
<# 
.SYNOPSIS
 
    Get-SfBAssignedNumbers.ps1 collects assigned phone numbers from all Skype for Business Server objects.
 
.DESCRIPTION
    Author: Andrew Morpeth
    Contact: https://ucgeek.co/
    
    This script queries Skype for Business Server for all assigned numbers and displays in a formatted table with the option to export to CSV. 
    During processing LineURI's are run against a regex pattern to extract the DDI/DID and the extension to a separate column.
    
    This script collects Skype for Business Server objects including:
    LineURI, Private Line, Analouge Lines, Common Area Phones, RGS Workflows, Exchange UM Contacts, Trusted Applications, Conferencing Numbers, 
    Meeting Rooms, Hybrid Application Endpoint (On-premises Resource Accounts)
    
    This script is provided as-is, no warrenty is provided or implied.The author is NOT responsible for any damages or data loss that may occur
    through the use of this script.  Always test before using in a production environment. This script is free to use for both personal and 
    business use, however, it may not be sold or included as part of a package that is for sale. A Service Provider may include this script 
    as part of their service offering/best practices provided they only charge for their time to implement and support.

.RUN INSTRUCTIONS 
    Update settings "Settings" at the top of the script
    Run: .\Get-SfBAssignedNumbers.ps1

.NOTES
    v1.0 - Initial release
    v1.4 - Added Get-CsMeetingRoom to data collection
    v1.5 - Added Get-CsHybridApplicationEndpoint to data collection
       
#>
#endregion INFO

# Settings
$FileName = "SfBAssignedNumbers_" + (Get-Date -Format s).replace(":","-") +".csv"
$FilePath = "C:\LXLSUPPORT\Scripts\$FileName"
$OutputType = "CSV" #OPTIONS: CSV - Outputs CSV to specified FilePath, CONSOLE - Outputs to console

Import-Module SkypeForBusiness

#$Regex1 = ‘tel:\+(\d+)(?:;ext=(\d+))?(?:;(\w+))?’
#$Regex1 = '^tel:\+(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'
$Regex1 = '^(?:tel:)?(?:\+)?(\d+)(?:;ext=(\d+))?(?:;([\w-]+))?$'

$Array1 = @()

#Get Users with LineURI
$UsersLineURI = Get-CsUser -Filter {LineURI -ne $Null}
if($UsersLineURI -ne $null)
{
    foreach($item in $UsersLineURI)
    {                  
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value $Item.FirstName
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value $Item.LastName
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "User"
        $Array1 += $myObject1          
    }
}

#Get Users with Private Line
$UsersPrivateLine = Get-CsUser -Filter {PrivateLine -ne $Null} 
if($UsersPrivateLine -ne $null)
{
    foreach($item in $UsersPrivateLine)
    {                   
        $Matches = @()
        $Item.PrivateLine -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.PrivateLine
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "FirstName" -Value $Item.FirstName
        $myObject1 | Add-Member -type NoteProperty -name "LastName" -Value $Item.LastName
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "UserPrivateLine"
        $Array1 += $myObject1          
    }
}

#Get analouge lines
$AnalougeLineURI = Get-CsAnalogDevice -Filter {LineURI -ne $Null}  
if($AnalougeLineURI -ne $null)
{
    foreach($item in $AnalougeLineURI)
    {                  
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "AnalougeLine"
        $Array1 += $myObject1          
    }
}

#Get common area phones
$CommonAreaLineURI = Get-CsCommonAreaPhone -Filter {LineURI -ne $Null} 
if($CommonAreaLineURI -ne $null)
{
    foreach($item in $CommonAreaLineURI)
    {                    
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "CommonArea"
        $Array1 += $myObject1          
    }
}

#Get RGS workflows
$WorkflowLineURI = Get-CsRgsWorkflow
if($WorkflowLineURI -ne $null)
{
    foreach($item in $WorkflowLineURI)
    {                 
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "RGSWorkflow"
        $Array1 += $myObject1          
    }
}

#Get Exchange UM Contacts
$ExUmContactLineURI = Get-CsExUmContact -Filter {LineURI -ne $Null}
if($ExUmContactLineURI -ne $null)
{
    foreach($item in $ExUmContactLineURI)
    {                   
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "ExUmContact"
        $Array1 += $myObject1          
    }
}

#Get trusted applications
$TrustedApplicationLineURI = Get-CsTrustedApplicationEndpoint -Filter {LineURI -ne $Null}
if($TrustedApplicationLineURI -ne $null)
{
    foreach($item in $TrustedApplicationLineURI)
    {                   
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.Name
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "TrustedApplication"
        $Array1 += $myObject1          
    }
}

#Get conferencing numbers
$DialInConfLineURI = Get-CsDialInConferencingAccessNumber -Filter {LineURI -ne $Null}
if($DialInConfLineURI -ne $null)
{
    foreach($Item in $DialInConfLineURI)
    {                 
        $Matches = @()
        $Item.LineURI -match $Regex1 | out-null
            
        $myObject1 = New-Object System.Object
        $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
        $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
        $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
        $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.DisplayName
        $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "DialInConf"
        $Array1 += $myObject1          
    }
}

    #Get meeting room numbers
    $MeetingRoomLineURI = Get-CsMeetingRoom -Filter {LineURI -ne $Null}
    if($MeetingRoomLineURI -ne $null)
    {
	    Write-Verbose "Processing Meeting Room Numbers"
        foreach($Item in $MeetingRoomLineURI)
        {                 
            $Matches = @()
            $Item.LineURI -match $Regex1 | out-null
            
            $myObject1 = New-Object System.Object
            $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
            $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
            $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
            $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.DisplayName
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value "MeetingRoom"
            $Array1 += $myObject1         
        }
    }

    #Get on-prem resource accounts
    $HybridApplicationEndpointLineURI = Get-CsHybridApplicationEndpoint -Filter {LineURI -ne $Null}
    if($HybridApplicationEndpointLineURI -ne $null)
    {
	    Write-Verbose "Processing Hybrid Application Endpoint (On-premises Resource Accounts) Numbers"
        foreach($Item in $HybridApplicationEndpointLineURI)
        {                 
            $Matches = @()
            $Item.LineURI -match $Regex1 | out-null
            
            $myObject1 = New-Object System.Object
            $myObject1 | Add-Member -type NoteProperty -name "LineURI" -Value $Item.LineURI
            $myObject1 | Add-Member -type NoteProperty -name "DDI" -Value $Matches[1]
            $myObject1 | Add-Member -type NoteProperty -name "Ext" -Value $Matches[2]
            $myObject1 | Add-Member -type NoteProperty -name "Name" -Value $Item.DisplayName
            $myObject1 | Add-Member -type NoteProperty -name "Type" -Value $(if ($item.ApplicationId -eq "ce933385-9390-45d1-9512-c8d228074e07") {"Auto Attendant On-premises Resource Account"} elseif ($item.ApplicationId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") {"Call Queue On-premises Resource Account"} else {"Unknown On-premises Resource Account"})
            $Array1 += $myObject1         
        }
    }

if($OutputType -eq "CSV")
{
    $Array1 | export-csv $FilePath -NoTypeInformation
    Write-Host "ALL DONE!! Your file has been saved to $FilePath. Press any key to quit"
}
elseif($OutputType -eq "CONSOLE")
{
    $Array1 | FT -AutoSize -Property LineURI,DDI,Ext,Name,Type
    Write-Host "ALL DONE!! Press any key to quit"
}
else
{
    $Array1 | FT -AutoSize -Property LineURI,DDI,Ext,Name,Type
    Write-Host "WARNING: Valid output type not set, defaulted to console. Press any key to quit"
}
cmd /c pause | out-null