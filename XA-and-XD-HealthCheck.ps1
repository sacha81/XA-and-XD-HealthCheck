#==============================================================================================
# Created on: 11.2014 modfied 01.2022 Version: 1.4.6
# Created by: Sacha / sachathomet.ch & Contributers (see changelog at EOF)
# File name: XA-and-XD-HealthCheck.ps1
#
# Description: This script checks a Citrix XenDesktop and/or XenApp 7.x Farm
# It generates a HTML output File which will be sent as Email.
#
# Initial versions tested on XenApp/XenDesktop 7.6 and XenDesktop 5.6 
# Newest version tested on XenApp/XenDesktop 7.11-7.13
#
# Prerequisite: Config file, a XenDesktop Controller with according privileges necessary 
# Config file:  In order for the script to work properly, it needs a configuration file.
#               This has the same name as the script, with extension _Parameters.
#               The script name can't contain any another point, even with a version.
#               Example: Script = "XA and XD HealthCheck.ps1", Config = "XA and XD HealthCheck_Parameters.xml"
#
# Call by :     Manual or by Scheduled Task, e.g. once a day
#               !! If you run it as scheduled task you need to add with argument "non interactive"
#               or your user has interactive persmission!
#
#
#==============================================================================================

# Don't change below here if you don't know what you are doing ... 
#==============================================================================================

#==============================================================================================
#Define variable to count script execution time and clear screen
$scriptstart = Get-Date

Clear-Host

#==============================================================================================

# Load only the snap-ins, which are used

if ($null -eq (Get-PSSnapin "Citrix.*" -EA silentlycontinue)) {
try { Add-PSSnapin Citrix.* -ErrorAction Stop }
catch { write-error "Error Get-PSSnapin Citrix.Broker.Admin.* Powershell snapin"; Return }
}

#==============================================================================================
# Import Variables from XML:

If (![string]::IsNullOrEmpty($hostinvocation)) {
	[string]$Global:ScriptPath = [System.IO.Path]::GetDirectoryName([System.Windows.Forms.Application]::ExecutablePath)
	[string]$Global:ScriptFile = [System.IO.Path]::GetFileName([System.Windows.Forms.Application]::ExecutablePath)
	[string]$global:ScriptName = [System.IO.Path]::GetFileNameWithoutExtension([System.Windows.Forms.Application]::ExecutablePath)
} ElseIf ($Host.Version.Major -lt 3) {
	[string]$Global:ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition
	[string]$Global:ScriptFile = Split-Path -Leaf $script:MyInvocation.MyCommand.Path
	[string]$global:ScriptName = $ScriptFile.Split('.')[0].Trim()
} Else {
	[string]$Global:ScriptPath = $PSScriptRoot
	[string]$Global:ScriptFile = Split-Path -Leaf $PSCommandPath
	[string]$global:ScriptName = $ScriptFile.Split('.')[0].Trim()
}

Set-StrictMode -Version Latest

# Import parameter file
$Global:ParameterFile = $ScriptName + "_Parameters.xml"
$Global:ParameterFilePath = $ScriptPath
[xml]$cfg = Get-Content ($ParameterFilePath + "\" + $ParameterFile) # Read content of XML file

# Import variables
Function New-XMLVariables {
	# Create a variable reference to the XML file
	$cfg.Settings.Variables.Variable | ForEach-Object {
		# Set Variables contained in XML file
		$VarValue = $_.Value
		$CreateVariable = $True # Default value to create XML content as Variable
		switch ($_.Type) {
			# Format data types for each variable 
			'[string]' { $VarValue = [string]$VarValue } # Fixed-length string of Unicode characters
			'[char]' { $VarValue = [char]$VarValue } # A Unicode 16-bit character
			'[byte]' { $VarValue = [byte]$VarValue } # An 8-bit unsigned character
            '[bool]' { If ($VarValue.ToLower() -eq 'false'){$VarValue = [bool]$False} ElseIf ($VarValue.ToLower() -eq 'true'){$VarValue = [bool]$True} } # An boolean True/False value
			'[int]' { $VarValue = [int]$VarValue } # 32-bit signed integer
			'[long]' { $VarValue = [long]$VarValue } # 64-bit signed integer
			'[decimal]' { $VarValue = [decimal]$VarValue } # A 128-bit decimal value
			'[single]' { $VarValue = [single]$VarValue } # Single-precision 32-bit floating point number
			'[double]' { $VarValue = [double]$VarValue } # Double-precision 64-bit floating point number
			'[DateTime]' { $VarValue = [DateTime]$VarValue } # Date and Time
			'[Array]' { $VarValue = [Array]$VarValue.Split(',') } # Array
			'[Command]' { $VarValue = Invoke-Expression $VarValue; $CreateVariable = $False } # Command
		}
		If ($CreateVariable) { New-Variable -Name $_.Name -Value $VarValue -Scope $_.Scope -Force }
	}
}

New-XMLVariables

$PvsWriteMaxSizeInGB = $PvsWriteMaxSize * 1Gb

ForEach ($DeliveryController in $DeliveryControllers){
    If ($DeliveryController -ieq "LocalHost"){
        $DeliveryController = [System.Net.DNS]::GetHostByName('').HostName
    }
    If (Test-Connection $DeliveryController) {
        $AdminAddress = $DeliveryController
        break
    }
}

$ReportDate = (Get-Date -UFormat "%A, %d. %B %Y %R")


$currentDir = Split-Path $MyInvocation.MyCommand.Path
$outputpath = Join-Path $currentDir "" #add here a custom output folder if you wont have it on the same directory
$logfile = Join-Path $outputpath ("CTXXDHealthCheck.log")
$resultsHTM = Join-Path $outputpath ("CTXXDHealthCheck.htm") #add $outputdate in filename if you like
  
#Header for Table "XD/XA Controllers" Get-BrokerController
$XDControllerFirstheaderName = "ControllerServer"
$XDControllerHeaderNames = "Ping", 	"State","DesktopsRegistered", 	"ActiveSiteServices"
$XDControllerHeaderWidths = "2",	"2", 	"2", 					"10"				
$XDControllerTableWidth= 1200
foreach ($disk in $diskLettersControllers)
{
    $XDControllerHeaderNames += "$($disk)Freespace"
    $XDControllerHeaderWidths += "4"
}
$XDControllerHeaderNames +=  	"AvgCPU", 	"MemUsg", 	"Uptime"
$XDControllerHeaderWidths +=    "4",		"4",		"4"

#Header for Table "Fail Rates" FUTURE
#$CTXFailureFirstheaderName = "Checks"
#$CTXFailureHeaderNames = "#","in Percentage", "CauseServiceInterruption","CausePartialServiceInterruption"
#$CTXFailureHeaderWidths = "2", "2", "2", 	"2"
#$CTXFailureTableWidth= 900

#Header for Table "CTX Licenses" Get-BrokerController
$CTXLicFirstheaderName = "LicenseName"
$CTXLicHeaderNames = "LicenseServer", 	"Count","InUse", 	"Available"
$CTXLicHeaderWidths = "4",	"2", 	"2", 					"2"
$CTXLicTableWidth= 900
  
#Header for Table "MachineCatalogs" Get-BrokerCatalog
$CatalogHeaderName = "CatalogName"
$CatalogHeaderNames = 	"AssignedToUser", 	"AssignedToDG", "NotToUserAssigned","ProvisioningType", "AllocationType", "MinimumFunctionalLevel", "UsedMCSSnapshot"
$CatalogWidths = 		"4",				"8", 			"8", 				"8", 				"8", 				"4", 				"4"
$CatalogTablewidth = 900
  
#Header for Table "DeliveryGroups" Get-BrokerDesktopGroup
$AssigmentFirstheaderName = "DeliveryGroup"
$vAssigmentHeaderNames = 	"PublishedName","DesktopKind", "SessionSupport", "ShutdownAfterUse", 	"TotalMachines","DesktopsAvailable","DesktopsUnregistered", "DesktopsInUse","DesktopsFree", "MaintenanceMode", "MinimumFunctionalLevel"
$vAssigmentHeaderWidths = 	"4", 			"4", 			"4", 	"4", 		"4", 				"4", 					"4", 			"4", 			"2", 			"2", 			"2"
$Assigmenttablewidth = 900
  
#Header for Table "VDI Checks" Get-BrokerMachine
$VDIfirstheaderName = "virtualDesktops"

$VDIHeaderNames = "CatalogName","DeliveryGroup","PowerState", "Ping", "MaintMode", 	"Uptime","LastConnect", 	"RegState","VDAVersion","AssociatedUserNames",  "WriteCacheType", "WriteCacheSize", "Tags", "HostedOn", "displaymode", "EDT_MTU", "OSBuild", "MCSVDIImageOutOfDate"
$VDIHeaderWidths = "4", "4",		"4","4", 	"4", 				"4", 		"4","4", 				"4",			  "4",			  "4",			  "4",			  "4", "4", "4", 		"4", "4", "4"

$VDItablewidth = 1200
  
#Header for Table "XenApp Checks" Get-BrokerMachine
$XenAppfirstheaderName = "virtualApp-Servers"
$XenAppHeaderNames = "CatalogName", "DeliveryGroup", "Serverload", 	"Ping", "MaintMode","Uptime", 	"RegState", "VDAVersion", "Spooler",  	"CitrixPrint", "OSBuild", "MCSImageOutOfDate"
$XenAppHeaderWidths = "4", 			"4", 				"4", 			"4", 	"4", 		"4", 		"4", 		"4", 		"4", 		 	"4", 		"4", 		"4"
foreach ($disk in $diskLettersWorkers)
{
    $XenAppHeaderNames += "$($disk)Freespace"
    $XenAppHeaderWidths += "4"
}

if ($ShowConnectedXenAppUsers -eq "1") { 

	$XenAppHeaderNames += "AvgCPU", 	"MemUsg", 	"ActiveSessions",  "WriteCacheType", "WriteCacheSize", "ConnectedUsers" , "Tags","HostedOn"
	$XenAppHeaderWidths +="4",		"4",			  "4",			"4",			"4",			"4",			"4","4"
}
else { 
	$XenAppHeaderNames += "AvgCPU", 	"MemUsg", 	"ActiveSessions", "WriteCacheType", "WriteCacheSize", "Tags","HostedOn"
	$XenAppHeaderWidths +="4",		"4",		"4",			  "4",			"4",			"4","4"

}

$XenApptablewidth = 1200
  
#==============================================================================================
#log function
function LogMe() {
Param(
[parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
[switch]$display,
[switch]$error,
[switch]$warning,
[switch]$progress
)
  
if ($error) { $logEntry = "[ERROR] $logEntry" ; Write-Host "$logEntry" -Foregroundcolor Red }
elseif ($warning) { Write-Warning "$logEntry" ; $logEntry = "[WARNING] $logEntry" }
elseif ($progress) { Write-Host "$logEntry" -Foregroundcolor Green }
elseif ($display) { Write-Host "$logEntry" }
  
#$logEntry = ((Get-Date -uformat "%D %T") + " - " + $logEntry)
$logEntry | Out-File $logFile -Append
}
  
#==============================================================================================
function Ping([string]$hostname, [int]$timeout = 200) {
$ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
try { $result = $ping.send($hostname, $timeout).Status.ToString() }
catch { $result = "Failure" }
return $result
}
#==============================================================================================
# The function will check the processor counter and check for the CPU usage. Takes an average CPU usage for 5 seconds. It check the current CPU usage for 5 secs.
Function CheckCpuUsage() 
{ 
	param ($hostname)
	Try { $CpuUsage=(Get-WmiObject -computer $hostname -class win32_processor | Measure-Object -property LoadPercentage -Average | Select-Object -ExpandProperty Average)
    $CpuUsage = [math]::round($CpuUsage, 1); return $CpuUsage


	} Catch { "Error returned while checking the CPU usage. Perfmon Counters may be fault" | LogMe -error; return 101 } 
}
#============================================================================================== 
# The function check the memory usage and report the usage value in percentage
Function CheckMemoryUsage()
{ 
	param ($hostname)
    Try 
	{   $SystemInfo = (Get-WmiObject -computername $hostname -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object TotalVisibleMemorySize, FreePhysicalMemory)
    	$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB 
    	$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB 
    	$UsedRAM = $TotalRAM - $FreeRAM 
    	$RAMPercentUsed = ($UsedRAM / $TotalRAM) * 100 
    	$RAMPercentUsed = [math]::round($RAMPercentUsed, 2);
    	return $RAMPercentUsed
	} Catch { "Error returned while checking the Memory usage. Perfmon Counters may be fault" | LogMe -error; return 101 } 
}
#==============================================================================================

# The function check the HardDrive usage and report the usage value in percentage and free space
Function CheckHardDiskUsage() 
{ 
	param ($hostname, $deviceID)
    Try 
	{   
    	$HardDisk = $null
		$HardDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $hostname -Filter "DeviceID='$deviceID'" -ErrorAction Stop | Select-Object Size,FreeSpace
        if ($null -ne $HardDisk)
		{
		$DiskTotalSize = $HardDisk.Size 
        $DiskFreeSpace = $HardDisk.FreeSpace 
        $frSpace=[Math]::Round(($DiskFreeSpace/1073741824),2)
		$PercentageDS = (($DiskFreeSpace / $DiskTotalSize ) * 100); $PercentageDS = [math]::round($PercentageDS, 2)
		
		Add-Member -InputObject $HardDisk -MemberType NoteProperty -Name PercentageDS -Value $PercentageDS
		Add-Member -InputObject $HardDisk -MemberType NoteProperty -Name frSpace -Value $frSpace
		} 
		
    	return $HardDisk
	} Catch { "Error returned while checking the Hard Disk usage. Perfmon Counters may be fault" | LogMe -error; return $null } 
}


#==============================================================================================
Function writeHtmlHeader
{
param($title, $fileName)
$date = $ReportDate
$head = @"
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<title>$title</title>
<STYLE TYPE="text/css">
<!--
td {
font-family: Tahoma;
font-size: 11px;
border-top: 1px solid #999999;
border-right: 1px solid #999999;
border-bottom: 1px solid #999999;
border-left: 1px solid #999999;
padding-top: 0px;
padding-right: 0px;
padding-bottom: 0px;
padding-left: 0px;
overflow: hidden;
}
body {
margin-left: 5px;
margin-top: 5px;
margin-right: 0px;
margin-bottom: 10px;
table {
table-layout:fixed;
border: thin solid #000000;
}
-->
</style>
</head>
<body>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='48' align='center' valign="middle">
<font face='tahoma' color='#003399' size='4'>
<strong>$title - $date</strong></font>
</td>
</tr>
</table>
"@
$head | Out-File $fileName
}
  
# ==============================================================================================
Function writeTableHeader
{
param($fileName, $firstheaderName, $headerNames, $headerWidths, $tablewidth)
$tableHeader = @"
  
<table width='$tablewidth'><tbody>
<tr bgcolor=#CCCCCC>
<td width='6%' align='center'><strong>$firstheaderName</strong></td>
"@
  
$i = 0
while ($i -lt $headerNames.count) {
$headerName = $headerNames[$i]
$headerWidth = $headerWidths[$i]
$tableHeader += "<td width='" + $headerWidth + "%' align='center'><strong>$headerName</strong></td>"
$i++
}
  
$tableHeader += "</tr>"
  
$tableHeader | Out-File $fileName -append
}
  
# ==============================================================================================
Function writeTableFooter
{
param($fileName)
"</table><br/>"| Out-File $fileName -append
}
  
#==============================================================================================
Function writeData
{
param($data, $fileName, $headerNames)

$tableEntry  =""  
$data.Keys | Sort-Object | ForEach-Object {
$tableEntry += "<tr>"
$computerName = $_
$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")
#$data.$_.Keys | foreach {
$headerNames | ForEach-Object {
#"$computerName : $_" | LogMe -display
try {
if ($data.$computerName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
elseif ($data.$computerName.$_[0] -eq "WARNING") { $bgcolor = "#FF7700"; $fontColor = "#FFFFFF" }
elseif ($data.$computerName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF" }
else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
$testResult = $data.$computerName.$_[1]
}
catch {
$bgcolor = "#CCCCCC"; $fontColor = "#003399"
$testResult = ""
}
$tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")
}
$tableEntry += "</tr>"
}
$tableEntry | Out-File $fileName -append
}
  
# ==============================================================================================
Function writeHtmlFooter
{
param($fileName)
@"
</table>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#000000' size='2'>

<strong>Uptime Threshold: </strong> $maxUpTimeDays days <br>
<strong>Database: </strong> $dbinfo <br>
<strong>LicenseServerName: </strong> $lsname <strong>LicenseServerPort: </strong> $lsport <br>
<strong>ConnectionLeasingEnabled: </strong> $CLeasing <br>
<strong>LocalHostCacheEnabled: </strong> $LHC <br>
<strong>HypervisorConnectionstate: </strong> $HVCS <br>

</font>
</td>
</table>
</body>
</html>
"@ | Out-File $FileName -append
}

# ==============================================================================================
Function ToHumanReadable()
{
  param($timespan)
  
  If ($timespan.TotalHours -lt 1) {
    return $timespan.Minutes + "minutes"
  }

  $sb = New-Object System.Text.StringBuilder
  If ($timespan.Days -gt 0) {
    [void]$sb.Append($timespan.Days)
    [void]$sb.Append(" days")
    [void]$sb.Append(", ")    
  }
  If ($timespan.Hours -gt 0) {
    [void]$sb.Append($timespan.Hours)
    [void]$sb.Append(" hours")
  }
  If ($timespan.Minutes -gt 0) {
    [void]$sb.Append(" and ")
    [void]$sb.Append($timespan.Minutes)
    [void]$sb.Append(" minutes")
  }
  return $sb.ToString()
}


# if enabled for Citrix Cloud set the credential profile: 
# Help from https://www.citrix.com/blogs/2016/07/01/introducing-remote-powershell-sdk-v2-for-citrix-cloud/ and 
# from https://hallspalmer.wordpress.com/2019/02/19/manage-citrix-cloud-using-powershell/ 
if ( $CitrixCloudCheck -eq "1" ) {
Set-XDCredentials -CustomerId $CustomerID -SecureClientFile $SecureClientFile -ProfileType CloudApi -StoreAs default
}

# ==============================================================================================
<#
	.SYNOPSIS
		Get information about user that set maintenance mode.
	
	.DESCRIPTION
		Over the Citrix XenDeesktop or XenApp log database, you can finde the user that
		set the maintenance mode of an worker.
		This is version 1.0.
	
	.PARAMETER AdminAddress
		Specifies the address of the Delivery Controller to which the PowerShell module will connect. This can be provided as a host name or an IP address.
	
	.PARAMETER Credential
		Specifies a user account that has permission to perform this action. The default is the current user.
	
	.EXAMPLE
		Get-CitrixMaintenanceInfo
		Get the informations on an delivery controller with nedded credentials.
	
	.EXAMPLE
		Get-CitrixMaintenanceInfo -AdminAddress server.domain.tld -Credential (Get-Credential)
		Use sever.domain.tld to get the log informations and use credentials.

	.LINK
		http://www.beckmann.ch/blog/2016/11/01/get-user-who-set-maintenance-mode-for-a-server-or-client/
#>
function Get-CitrixMaintenanceInfo {
	[CmdletBinding()]
	[OutputType([System.Management.Automation.PSCustomObject])]
	param
	(
		[Parameter(Mandatory = $false,
				   ValueFromPipeline = $true,
				   Position = 0)]
		[System.String[]]$AdminAddress = 'localhost',
		[Parameter(Mandatory = $false,
				   ValueFromPipeline = $true,
				   Position = 1)]
		[System.Management.Automation.PSCredential]$Credential
	) # Param
	
	Try {
		$PSSessionParam = @{ }
		If ($null -ne $Credential) { $PSSessionParam['Credential'] = $Credential } #Splatting
		If ($null -ne $AdminAddress) { $PSSessionParam['ComputerName'] = $AdminAddress } #Splatting
		
		# Create Session
		$Session = New-PSSession -ErrorAction Stop @PSSessionParam
		
		# Create script block for invoke command
		$ScriptBlock = {
			if ($null -eq (Get-PSSnapin "Get-PSSnapin Citrix.ConfigurationLogging.Admin.*" -ErrorAction silentlycontinue)) {
				try { Add-PSSnapin Citrix.ConfigurationLogging.Admin.* -ErrorAction Stop } catch { write-error "Error Get-PSSnapin Citrix.ConfigurationLogging.Admin.* Powershell snapin"; Return }
			} #If
			
			$Date = Get-Date
			$StartDate = $Date.AddDays(-7) # Hard coded value for how many days back
			$EndDate = $Date
			
			# Command to get the informations from log
			$LogEntrys = Get-LogLowLevelOperation -MaxRecordCount 1000000 -Filter { StartTime -ge $StartDate -and EndTime -le $EndDate } | Where-Object { $_.Details.PropertyName -eq 'MAINTENANCEMODE' } | Sort-Object EndTime -Descending
			
			# Build an object with the data for the output
			[array]$arrMaintenance = @()
			ForEach ($LogEntry in $LogEntrys) {
				$TempObj = New-Object -TypeName psobject -Property @{
					User = $LogEntry.User
					TargetName = $LogEntry.Details.TargetName
					NewValue = $LogEntry.Details.NewValue
					PreviousValue = $LogEntry.Details.PreviousValue
					StartTime = $LogEntry.Details.StartTime
					EndTime = $LogEntry.Details.EndTime
				} #TempObj
				$arrMaintenance += $TempObj
			} #ForEach				
			$arrMaintenance
		} # ScriptBlock
		
		# Run the script block with invoke-command, return the values and close the session
		$MaintLogs = Invoke-Command -Session $Session -ScriptBlock $ScriptBlock -ErrorAction Stop
		Write-Output $MaintLogs
		Remove-PSSession -Session $Session -ErrorAction SilentlyContinue
		
	} Catch {
		Write-Warning "Error occurs: $_"
	} # Try/Catch
} # Get-CitrixMaintenanceInfo

#==============================================================================================

$wmiOSBlock = {param($computer)
  try { $wmi=Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop }
  catch { $wmi = $null }
  return $wmi
}

#==============================================================================================
# == MAIN SCRIPT ==
#==============================================================================================
Remove-Item $logfile -force -EA SilentlyContinue
Remove-Item $resultsHTM -force -EA SilentlyContinue

"#### Begin with Citrix XenDestop / XenApp HealthCheck ######################################################################" | LogMe -display -progress
  
" " | LogMe -display -progress

# get some farm infos, which will be presented in footer 
$dbinfo = Get-BrokerDBConnection -AdminAddress $AdminAddress
$brokersiteinfos = Get-BrokerSite
$lsname = $brokersiteinfos.LicenseServerName
$lsport = $brokersiteinfos.LicenseServerPort
$CLeasing = $brokersiteinfos.ConnectionLeasingEnabled
$LHC =$brokersiteinfos.LocalHostCacheEnabled

$BrkrHvsCon = Get-Brokerhypervisorconnection
$HVCS =$BrkrHvsCon.State

# Log the loaded Citrix PS Snapins
(Get-PSSnapin "Citrix.*" -EA silentlycontinue).Name | ForEach-Object {"PSSnapIn: " + $_ | LogMe -display -progress}
  
#== Controller Check ============================================================================================
"Check Controllers #############################################################################" | LogMe -display -progress
  
" " | LogMe -display -progress
  
$ControllerResults = @{}
$Controllers = Get-BrokerController -AdminAddress $AdminAddress

# Get first DDC version (should be all the same unless an upgrade is in progress)
$ControllerVersion = $Controllers[0].ControllerVersion
"Version: $controllerversion " | LogMe -display -progress
  
if ($ControllerVersion -lt 7 ) {
  "XenDesktop/XenApp Version below 7.x ($controllerversion) - only DesktopCheck will be performed" | LogMe -display -progress
  #$ShowXenAppTable = 0 #doesent work with XML variables
  Set-Variable -Name ShowXenAppTable -Value 0
} else { 
  "XenDesktop/XenApp Version above 7.x ($controllerversion) - XenApp and DesktopCheck will be performed" | LogMe -display -progress
}

foreach ($Controller in $Controllers) {
$tests = @{}
  
#Name of $Controller
$ControllerDNS = $Controller | ForEach-Object{ $_.DNSName }
"Controller: $ControllerDNS" | LogMe -display -progress
  
#Ping $Controller
$result = Ping $ControllerDNS 100
if ($result -ne "SUCCESS") { $tests.Ping = "Error", $result }
else { $tests.Ping = "SUCCESS", $result 

#Now when Ping is ok also check this:
  
#State of this controller
$ControllerState = $Controller | ForEach-Object{ $_.State }
"State: $ControllerState" | LogMe -display -progress
if ($ControllerState -ne "Active") { $tests.State = "ERROR", $ControllerState }
else { $tests.State = "SUCCESS", $ControllerState }


  
#DesktopsRegistered on this controller
$ControllerDesktopsRegistered = $Controller | ForEach-Object{ $_.DesktopsRegistered }
"Registered: $ControllerDesktopsRegistered" | LogMe -display -progress
$tests.DesktopsRegistered = "NEUTRAL", $ControllerDesktopsRegistered
  
#ActiveSiteServices on this controller
$ActiveSiteServices = $Controller | ForEach-Object{ $_.ActiveSiteServices }
"ActiveSiteServices $ActiveSiteServices" | LogMe -display -progress
$tests.ActiveSiteServices = "NEUTRAL", $ActiveSiteServices


#==============================================================================================
#               CHECK CPU AND MEMORY USAGE 
#==============================================================================================

        # Check the AvgCPU value for 5 seconds
        $AvgCPUval = CheckCpuUsage ($ControllerDNS)
		#$VDtests.LoadBalancingAlgorithm = "SUCCESS", "LB is set to BEST EFFORT"} 
			
        if( [int] $AvgCPUval -lt 75) { "CPU usage is normal [ $AvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$AvgCPUval %" }
		elseif([int] $AvgCPUval -lt 85) { "CPU usage is medium [ $AvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$AvgCPUval %" }   	
		elseif([int] $AvgCPUval -lt 95) { "CPU usage is high [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }
		elseif([int] $AvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
        else { "CPU usage is Critical [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }   
		$AvgCPUval = 0

        # Check the Physical Memory usage       
        $UsedMemory = CheckMemoryUsage ($ControllerDNS)
        if( $UsedMemory -lt 75) { "Memory usage is normal [ $UsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$UsedMemory %" }
		elseif( [int] $UsedMemory -lt 85) { "Memory usage is medium [ $UsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$UsedMemory %" }   	
		elseif( [int] $UsedMemory -lt 95) { "Memory usage is high [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" }
		elseif( [int] $UsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" }
        else { "Memory usage is Critical [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" }   
		$UsedMemory = 0  

        foreach ($disk in $diskLettersControllers)
        {
            # Check Disk Usage 
		    $HardDisk = CheckHardDiskUsage -hostname $ControllerDNS -deviceID "$($disk):"
		    if ($null -ne $HardDisk) {	
			    $XAPercentageDS = $HardDisk.PercentageDS
			    $frSpace = $HardDisk.frSpace
			
	            If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
			    ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "Err" }
			    ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" } 
			    ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" }     
	            Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" }  
        
			    $XAPercentageDS = 0
			    $frSpace = 0
			    $HardDisk = $null
		    }
        }
		
    # Check uptime (Query over WMI)
    $tests.WMI = "ERROR","Error"
    try { $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $ControllerDNS }
    catch { $wmi = $null }

    # Perform WMI related checks
    if ($null -ne $wmi) {
        $tests.WMI = "SUCCESS", "Success"
        $LBTime=$wmi.ConvertToDateTime($wmi.Lastbootuptime)
        [TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)

        if ($uptime.days -lt $minUpTimeDaysDDC){
            "reboot warning, last reboot: {0:D}" -f $LBTime | LogMe -display -warning
            $tests.Uptime = "WARNING", (ToHumanReadable($uptime))
        }
        else { $tests.Uptime = "SUCCESS", (ToHumanReadable($uptime)) }
    }
    else { "WMI connection failed - check WMI for corruption" | LogMe -display -error }
}


  
" --- " | LogMe -display -progress
#Fill $tests into array

$ControllerResults.$ControllerDNS = $tests
}
  
#== Catalog Check ============================================================================================
"Check Catalog #################################################################################" | LogMe -display -progress
" " | LogMe -display -progress
  
$CatalogResults = @{}
$Catalogs = Get-BrokerCatalog -AdminAddress $AdminAddress
  
foreach ($Catalog in $Catalogs) {
  $tests = @{}
  
  #Name of MachineCatalog
  $CatalogName = $Catalog | ForEach-Object{ $_.Name }
  "Catalog: $CatalogName" | LogMe -display -progress

  if ($ExcludedCatalogs -contains $CatalogName) {
    "Excluded Catalog, skipping" | LogMe -display -progress
  } else {
    #CatalogAssignedCount
    $CatalogAssignedCount = $Catalog | ForEach-Object{ $_.AssignedCount }
    "Assigned: $CatalogAssignedCount" | LogMe -display -progress
    $tests.AssignedToUser = "NEUTRAL", $CatalogAssignedCount
  
    #CatalogUnassignedCount
    $CatalogUnAssignedCount = $Catalog | ForEach-Object{ $_.UnassignedCount }
    "Unassigned: $CatalogUnAssignedCount" | LogMe -display -progress
    $tests.NotToUserAssigned = "NEUTRAL", $CatalogUnAssignedCount
  
    # Assigned to DeliveryGroup
    $CatalogUsedCountCount = $Catalog | ForEach-Object{ $_.UsedCount }
    "Used: $CatalogUsedCountCount" | LogMe -display -progress
    $tests.AssignedToDG = "NEUTRAL", $CatalogUsedCountCount

    #MinimumFunctionalLevel
	$MinimumFunctionalLevel = $Catalog | ForEach-Object{ $_.MinimumFunctionalLevel }
	"MinimumFunctionalLevel: $MinimumFunctionalLevel" | LogMe -display -progress
    $tests.MinimumFunctionalLevel = "NEUTRAL", $MinimumFunctionalLevel
  
     #ProvisioningType
     $CatalogProvisioningType = $Catalog | ForEach-Object{ $_.ProvisioningType }
     "ProvisioningType: $CatalogProvisioningType" | LogMe -display -progress
     $tests.ProvisioningType = "NEUTRAL", $CatalogProvisioningType
  
     #AllocationType
     $CatalogAllocationType = $Catalog | ForEach-Object{ $_.AllocationType }
     "AllocationType: $CatalogAllocationType" | LogMe -display -progress
     $tests.AllocationType = "NEUTRAL", $CatalogAllocationType


     #UsedMcsSnapshot 
     $UsedMcsSnapshot = ""
     $MCSInfo.MasterImageVM = ""

     $CatalogProvisioningSchemeId = $Catalog | ForEach-Object{ $_.ProvisioningSchemeId }

     "ProvisioningSchemeId: $CatalogProvisioningSchemeId " | LogMe -display -progress
     $MCSInfo = (Get-ProvScheme -ProvisioningSchemeUid $CatalogProvisioningSchemeId)

     $UsedMcsSnapshot = $MCSInfo.MasterImageVM


     #"MasterImageVM: $MCSInfo.MasterImageVM"
     $UsedMcsSnapshot = $UsedMcsSnapshot.trimstart("XDHyp:\HostingUnits\") = $MCSInfo.MasterImageVM
     $UsedMcsSnapshot = $UsedMcsSnapshot.trimend(".template")
     "UsedMcsSnapshot: = $UsedMcsSnapshot"


     $tests.UsedMcsSnapshot  = "NEUTRAL", $UsedMcsSnapshot 


  
    "", ""
    $CatalogResults.$CatalogName = $tests
  }  
  " --- " | LogMe -display -progress
}
  
#== DeliveryGroups Check ============================================================================================
"Check Assigments #############################################################################" | LogMe -display -progress
  
" " | LogMe -display -progress
  
$AssigmentsResults = @{}
$Assigments = Get-BrokerDesktopGroup -AdminAddress $AdminAddress
  
foreach ($Assigment in $Assigments) {
  $tests = @{}
  
  #Name of DeliveryGroup
  $DeliveryGroup = $Assigment | ForEach-Object{ $_.Name }
  "DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
  
  if ($ExcludedCatalogs -contains $DeliveryGroup) {
    "Excluded Delivery Group, skipping" | LogMe -display -progress
  } else {
  
    #PublishedName
    $AssigmentDesktopPublishedName = $Assigment | ForEach-Object{ $_.PublishedName }
    "PublishedName: $AssigmentDesktopPublishedName" | LogMe -display -progress
    $tests.PublishedName = "NEUTRAL", $AssigmentDesktopPublishedName
  
    #DesktopsTotal
    $TotalDesktops = $Assigment | ForEach-Object{ $_.TotalDesktops }
    "DesktopsAvailable: $TotalDesktops" | LogMe -display -progress
    $tests.TotalMachines = "NEUTRAL", $TotalDesktops
  
    #DesktopsAvailable
    $AssigmentDesktopsAvailable = $Assigment | ForEach-Object{ $_.DesktopsAvailable }
    "DesktopsAvailable: $AssigmentDesktopsAvailable" | LogMe -display -progress
    $tests.DesktopsAvailable = "NEUTRAL", $AssigmentDesktopsAvailable
  
    #DesktopKind
    $AssigmentDesktopsKind = $Assigment | ForEach-Object{ $_.DesktopKind }
    "DesktopKind: $AssigmentDesktopsKind" | LogMe -display -progress
    $tests.DesktopKind = "NEUTRAL", $AssigmentDesktopsKind
	
	#SessionSupport
	$SessionSupport = $Assigment | ForEach-Object{ $_.SessionSupport }
	"SessionSupport: $SessionSupport" | LogMe -display -progress
    $tests.SessionSupport = "NEUTRAL", $SessionSupport
	
	#ShutdownAfterUse
	$ShutdownDesktopsAfterUse = $Assigment | ForEach-Object{ $_.ShutdownDesktopsAfterUse }
	"ShutdownDesktopsAfterUse: $ShutdownDesktopsAfterUse" | LogMe -display -progress
    
	if ($SessionSupport -eq "MultiSession" -and $ShutdownDesktopsAfterUse -eq "$true" ) { 
	$tests.ShutdownAfterUse = "ERROR", $ShutdownDesktopsAfterUse
	}
	else { 
	 $tests.ShutdownAfterUse = "NEUTRAL", $ShutdownDesktopsAfterUse
	}
	

    #MinimumFunctionalLevel
	$MinimumFunctionalLevel = $Assigment | ForEach-Object{ $_.MinimumFunctionalLevel }
	"MinimumFunctionalLevel: $MinimumFunctionalLevel" | LogMe -display -progress
    $tests.MinimumFunctionalLevel = "NEUTRAL", $MinimumFunctionalLevel
	
	if ($SessionSupport -eq "MultiSession" ) { 
	
	$tests.DesktopsFree = "NEUTRAL", "N/A"
	$tests.DesktopsInUse = "NEUTRAL", "N/A"
		
	}
    else { 
			#DesktopsInUse
			$AssigmentDesktopsInUse = $Assigment | ForEach-Object{ $_.DesktopsInUse }
			"DesktopsInUse: $AssigmentDesktopsInUse" | LogMe -display -progress
			$tests.DesktopsInUse = "NEUTRAL", $AssigmentDesktopsInUse
	
			#DesktopFree
			$AssigmentDesktopsFree = $AssigmentDesktopsAvailable - $AssigmentDesktopsInUse
			"DesktopsFree: $AssigmentDesktopsFree" | LogMe -display -progress
  
			if ($AssigmentDesktopsKind -eq "shared") {
			if ($AssigmentDesktopsFree -gt 0 ) {
				"DesktopsFree < 1 ! ($AssigmentDesktopsFree)" | LogMe -display -progress
				$tests.DesktopsFree = "SUCCESS", $AssigmentDesktopsFree
			} elseif ($AssigmentDesktopsFree -lt 0 ) {
				"DesktopsFree < 1 ! ($AssigmentDesktopsFree)" | LogMe -display -progress
				$tests.DesktopsFree = "SUCCESS", "N/A"
			} else {
				$tests.DesktopsFree = "WARNING", $AssigmentDesktopsFree
				"DesktopsFree > 0 ! ($AssigmentDesktopsFree)" | LogMe -display -progress
			}
			} else {
			$tests.DesktopsFree = "NEUTRAL", "N/A"
			}
	
	
	}
		
  
    #inMaintenanceMode
    $AssigmentDesktopsinMaintenanceMode = $Assigment | ForEach-Object{ $_.inMaintenanceMode }
    "inMaintenanceMode: $AssigmentDesktopsinMaintenanceMode" | LogMe -display -progress
    if ($AssigmentDesktopsinMaintenanceMode) { $tests.MaintenanceMode = "WARNING", "ON" }
    else { $tests.MaintenanceMode = "SUCCESS", "OFF" }
  
    #DesktopsUnregistered
    $AssigmentDesktopsUnregistered = $Assigment | ForEach-Object{ $_.DesktopsUnregistered }
    "DesktopsUnregistered: $AssigmentDesktopsUnregistered" | LogMe -display -progress    
    if ($AssigmentDesktopsUnregistered -gt 0 ) {
      "DesktopsUnregistered > 0 ! ($AssigmentDesktopsUnregistered)" | LogMe -display -progress
      $tests.DesktopsUnregistered = "WARNING", $AssigmentDesktopsUnregistered
    } else {
      $tests.DesktopsUnregistered = "SUCCESS", $AssigmentDesktopsUnregistered
      "DesktopsUnregistered <= 0 ! ($AssigmentDesktopsUnregistered)" | LogMe -display -progress
    }
  
    
      
    #Fill $tests into array
    $AssigmentsResults.$DeliveryGroup = $tests
  }
  " --- " | LogMe -display -progress
}
  
# ======= License Check ========
if($ShowCTXLicense -eq 1 ){

    $myCollection = @()
    try 
	{
        $LicWMIQuery = get-wmiobject -namespace "ROOT\CitrixLicensing" -computer $lsname -query "select * from Citrix_GT_License_Pool" -ErrorAction Stop | ? {$_.PLD -in $CTXLicenseMode}
        
	
        foreach ($group in $($LicWMIQuery | group pld))
        {
            $lics = $group | Select-Object -ExpandProperty group
            $i = 1

            $myArray_Count = 0
		    $myArray_InUse = 0
		    $myArray_Available = 0
		
		    foreach ($lic in @($lics))
		    {
		    $myArray = "" | Select-Object LicenseServer,LicenceName,Count,InUse,Available
		    $myArray.LicenseServer = $lsname
		    $myArray.LicenceName = "$($lics.pld) ($i) Licence"
		    $myArray.Count = $Lic.count - $Lic.Overdraft
		    if ($Lic.inusecount -gt $myArray.Count) {$myArray.InUse = $myArray.Count} else {$myArray.InUse = $Lic.inusecount}
		    $myArray.Available = $myArray.count - $myArray.InUse
		    $myCollection += $myArray
		
		    $myArray = "" | Select-Object LicenseServer,LicenceName,Count,InUse,Available
		    $myArray.LicenseServer = $lsname
		    $myArray.LicenceName = "$($lics.pld) ($i) Overdraft"
		    $myArray.Count = $Lic.Overdraft
		    if ($Lic.inusecount -gt $($Lic.count - $Lic.Overdraft)) {$myArray.InUse = $Lic.inusecount - $($Lic.count - $Lic.Overdraft)} else {$myArray.InUse = 0}
		    $myArray.Available = $myArray.count - $myArray.InUse
		    $myCollection += $myArray
		
		    $myArray_Count += $Lic.count
		    $myArray_InUse += $Lic.inusecount
		    $myArray_Available += $Lic.pooledavailable
				
		    $i++
		    }
		
		    $myArray = "" | Select-Object LicenseServer,LicenceName,Count,InUse,Available
		    $myArray.LicenseServer = $lsname
		    $myArray.LicenceName = "$($lics.pld) - Total"
		    $myArray.Count = $myArray_Count
		    $myArray.InUse = $myArray_InUse
		    $myArray.Available = $myArray_Available
		    $myCollection += $myArray

    }
    }
    catch
    {
            $myArray = "" | Select-Object LicenseServer,LicenceName,Count,InUse,Available
		    $myArray.LicenseServer = $lsname
		    $myArray.LicenceName = "n/a"
		    $myArray.Count = "n/a"
		    $myArray.InUse = "n/a"
		    $myArray.Available = "n/a"
		    $myCollection += $myArray 
    }
    
    $CTXLicResults = @{}

    foreach ($line in $myCollection)
    {
        $tests = @{}


        if ($line.LicenceName -eq "n/a")
        {
            $tests.LicenseServer ="error", $line.LicenseServer
            $tests.Count ="error", $line.Count
		    $tests.InUse ="error", $line.InUse
		    $tests.Available ="error", $line.Available
        }
        else
        {
            $tests.LicenseServer ="NEUTRAL", $line.LicenseServer
            $tests.Count ="NEUTRAL", $line.Count
		    $tests.InUse ="NEUTRAL", $line.InUse
		    $tests.Available ="NEUTRAL", $line.Available}
            $CTXLicResults.($line.LicenceName) =  $tests
        }

}
else {"CTX License Check skipped because ShowCTXLicense = 0 " | LogMe -display -progress }
  

# ======= Desktop Check ========
"Check virtual Desktops ####################################################################################" | LogMe -display -progress
" " | LogMe -display -progress
  
if($ShowDesktopTable -eq 1 ) {
  
$allResults = @{}
  
$machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress| Where-Object {$_.SessionSupport -eq "SingleSession" -and @(Compare-Object $_.tags $ExcludedTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0}
  
# SessionSupport only availiable in XD 7.x - for this reason only distinguish in Version above 7 if Desktop or XenApp
if($controllerversion -lt 7 ) { $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -and @(Compare-Object $_.tags $ExcludedTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0}
else { $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress| Where-Object {$_.SessionSupport -eq "SingleSession" -and @(Compare-Object $_.tags $ExcludedTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0} }

$Maintenance = Get-CitrixMaintenanceInfo -AdminAddress $AdminAddress

foreach($machine in $machines) {
$tests = @{}
  
$ErrorVDI = 0
  
# Column Name of Desktop
$machineDNS = $machine | ForEach-Object{ $_.DNSName }
"Machine: $machineDNS" | LogMe -display -progress
  
# Column CatalogName
$CatalogName = $machine | ForEach-Object{ $_.CatalogName }
"Catalog: $CatalogName" | LogMe -display -progress
$tests.CatalogName = "NEUTRAL", $CatalogName

# Column DeliveryGroup
$DeliveryGroup = $machine | ForEach-Object{ $_.DesktopGroupName }
"DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
$tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup

# Column Powerstate
$Powered = $machine | ForEach-Object{ $_.PowerState }
"PowerState: $Powered" | LogMe -display -progress
$tests.PowerState = "NEUTRAL", $Powered

if ($Powered -eq "Off" -OR $Powered -eq "Unknown") {
$tests.PowerState = "NEUTRAL", $Powered
}

if ($Powered -eq "On") {
$tests.PowerState = "SUCCESS", $Powered
}

if ($Powered -eq "On" -OR $Powered -eq "Unknown" -OR $Powered -eq "Unmanaged") {



# Column Ping Desktop
$result = Ping $machineDNS 100
if ($result -eq "SUCCESS") {
  $tests.Ping = "SUCCESS", $result
  
  #==============================================================================================
  # Column Uptime (Query over WMI - only if Ping successfull)
  $tests.WMI = "ERROR","Error"
  $job = Start-Job -ScriptBlock $wmiOSBlock -ArgumentList $machineDNS
  $wmi = Wait-job $job -Timeout 15 | Receive-Job

  # Perform WMI related checks
  if ($null -ne $wmi) {
    $tests.WMI = "SUCCESS", "Success"
    $LBTime=[Management.ManagementDateTimeConverter]::ToDateTime($wmi.Lastbootuptime)
    [TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)
  
    if ($uptime.days -gt $maxUpTimeDays) {
      "reboot warning, last reboot: {0:D}" -f $LBTime | LogMe -display -warning
      $tests.Uptime = "WARNING", $uptime.days
      $ErrorVDI = $ErrorVDI + 1
    } else { 
      $tests.Uptime = "SUCCESS", $uptime.days 
    }
  } else { 
    "WMI connection failed - check WMI for corruption" | LogMe -display -error
    stop-job $job
  }

  #-----------------
# Column WriteCacheSize (only if Ping is successful)
################ PVS SECTION ###############
if (test-path \\$machineDNS\c$\Personality.ini) {
# Test if PVS cache is of type "device's hard drive"
$PvsWriteCacheUNC = Join-Path "\\$machineDNS" ($PvsWriteCacheDrive+"$"+"\.vdiskcache")
$CacheDiskOnHD = Test-Path $PvsWriteCacheUNC

if ($CacheDiskOnHD -eq $True) {
  $CacheDiskExists = $True
  $CachePVSType = "Device HD"
} else {
  # Test if PVS cache is of type "device RAM with overflow to hard drive"
  $PvsWriteCacheUNC = Join-Path "\\$machineDNS" ($PvsWriteCacheDrive+"$"+"\vdiskdif.vhdx")
  $CacheDiskRAMwithOverflow = Test-Path $PvsWriteCacheUNC
  if ($CacheDiskRAMwithOverflow -eq $True) {
    $CacheDiskExists = $True
    $CachePVSType = "Device RAM with overflow to disk"
  } else {
    $CacheDiskExists = $False
    $CachePVSType = ""
  }
}

if ($CacheDiskExists -eq $True) {
$CacheDisk = [long] ((get-childitem $PvsWriteCacheUNC -force).length)
$CacheDiskGB = "{0:n2}GB" -f($CacheDisk / 1GB)
"PVS Cache file size: {0:n2}GB" -f($CacheDisk / 1GB) | LogMe
#"PVS Cache max size: {0:n2}GB" -f($PvsWriteMaxSizeInGB / 1GB) | LogMe -display
$tests.WriteCacheType = "NEUTRAL", $CachePVSType
if ($CacheDisk -lt ($PvsWriteMaxSizeInGB * 0.5)) {
"WriteCache file size is low" | LogMe
$tests.WriteCacheSize = "SUCCESS", $CacheDiskGB
}
elseif ($CacheDisk -lt ($PvsWriteMaxSizeInGB * 0.8)) {
"WriteCache file size moderate" | LogMe -display -warning
$tests.WriteCacheSize = "WARNING", $CacheDiskGB
}
else {
"WriteCache file size is high" | LogMe -display -error
$tests.WriteCacheSize = "ERROR", $CacheDiskGB
}
}
$Cachedisk = 0
}
else { $tests.WriteCacheSize = "SUCCESS", "N/A" }
############## END PVS SECTION #############

# Column OSBuild 
$MachineOSVersion = "N/A"
$MachineOSVersion = (Get-ItemProperty -Path "\\$machineDNS\C$\WINDOWS\System32\hal.dll" -ErrorAction SilentlyContinue).VersionInfo.FileVersion.Split()[0]
$tests.OSBuild = "NEUTRAL", $MachineOSVersion


#---------------------
  
  }
else {
$tests.Ping = "Neutral", $result
$ErrorVDI = $ErrorVDI + 0 # Ping is no definitve indicator for a problem
}
#END of Ping-Section

# Column RegistrationState
$RegistrationState = $machine | ForEach-Object{ $_.RegistrationState }
"State: $RegistrationState" | LogMe -display -progress
if ($RegistrationState -ne "Registered") {
$tests.RegState = "ERROR", $RegistrationState
$ErrorVDI = $ErrorVDI + 1
}
else { $tests.RegState = "SUCCESS", $RegistrationState }

} 
 
# Column MaintenanceMode
$MaintenanceMode = $machine | ForEach-Object{ $_.InMaintenanceMode }
"MaintenanceMode: $MaintenanceMode" | LogMe -display -progress
if ($MaintenanceMode) {
	$objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $machine.MachineName.ToUpper() } | Select-Object -First 1
	If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
	"MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
	$tests.MaintMode = "WARNING", $MaintenanceModeOn
	$ErrorVDI = $ErrorVDI + 1
}
else { $tests.MaintMode = "SUCCESS", "OFF" }
  
# Column HostedOn 
$HostedOn = $machine | ForEach-Object{ $_.HostingServerName }
"HostedOn: $HostedOn" | LogMe -display -progress
$tests.HostedOn = "NEUTRAL", $HostedOn

# Column VDAVersion AgentVersion
$VDAVersion = $machine | ForEach-Object{ $_.AgentVersion }
"VDAVersion: $VDAVersion" | LogMe -display -progress
$tests.VDAVersion = "NEUTRAL", $VDAVersion

# Column AssociatedUserNames
$AssociatedUserNames = $machine | ForEach-Object{ $_.AssociatedUserNames }
"Assigned to $AssociatedUserNames" | LogMe -display -progress
$tests.AssociatedUserNames = "NEUTRAL", $AssociatedUserNames

# Column Tags 
$Tags = $machine | ForEach-Object{ $_.Tags }
"Tags: $Tags" | LogMe -display -progress
$tests.Tags = "NEUTRAL", $Tags

# Column MCSVDIImageOutOfDate
$MCSVDIImageOutOfDate = $machine | ForEach-Object{ $_.ImageOutOfDate }
"ImageOutOfDate: $MCSVDIImageOutOfDate" | LogMe -display -progress
if ($MCSVDIImageOutOfDate -eq $true) { $tests.MCSImageOutOfDate = "ERROR", $MCSVDIImageOutOfDate }
elseif ($MCSVDIImageOutOfDate -eq $false) { $tests.MCSImageOutOfDate = "SUCCESS", $MCSVDIImageOutOfDate  }
else { $tests.MCSVDIImageOutOfDate = "NEUTRAL", $MCSVDIImageOutOfDate }


## Column LastConnect
$yellow =((Get-Date).AddMonths(-1).ToString('yyyy-MM-dd HH:mm:s'))
$red =((Get-Date).AddMonths(-3).ToString('yyyy-MM-dd HH:mm:s'))

$machineLastConnect = $machine | ForEach-Object{ $_.LastConnectionTime }

if ([string]::IsNullOrWhiteSpace($machineLastConnect))
	{
		$tests.LastConnect = "NEUTRAL", "NO DATA"
	}
elseif ($machineLastConnect -lt $red)
	{
		"LastConnect: $machineLastConnect" | LogMe -display -ERROR
		$tests.LastConnect = "ERROR", $machineLastConnect
	} 	
elseif ($machineLastConnect -lt $yellow)
	{
		"LastConnect: $machineLastConnect" | LogMe -display -WARNING
		$tests.LastConnect = "WARNING", $machineLastConnect
	}
else 
	{
		$tests.LastConnect = "SUCCESS", $machineLastConnect
		"LastConnect: $machineLastConnect" | LogMe -display -progress
	}
## End Column LastConnect

## EDT MTU (set by default.ica or MTUDiscovery)
$EDTMTU = Invoke-Command -ComputerName $machineDNS -ScriptBlock {(ctxsession -v | findstr "EDT MTU:" | select -Last 1).split(":")[1].trimstart()}
$tests.EDT_MTU = "NEUTRAL", $EDTMTU
"EDT MTU Size is set to $EDTMTU" | LogMe -display -progress



# Column displaymode when a User has a Session
$sessionUser = $machine | ForEach-Object{ $_.SessionUserName }

$displaymode = "N/A"
if ( $ShowGraphicsMode -eq "1" ) {

if ($sessionUser -notlike "" )
{

$displaymode = "unknown"
$displaymodeTable = @{}


#H264
$displaymodeTable.H264Active = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr IsActive=*

    # H.264 Pure
    #Component_Encoder=DeepCompressionV2Encoder	
	$displaymodeTable.Component_Encoder_DeepCompressionEncoder = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=DeepCompressionEncoder
	if ($displaymodeTable.Component_Encoder_DeepCompressionEncoder -eq "Component_Encoder=DeepCompressionEncoder")
	{
	$Displaymode = "Pure H.264"
	}
	
	# Thinwire H.264 + Lossless (true native H264)
    #Component_Encoder=DeepCompressionV2Encoder
	$displaymodeTable.Component_Encoder_DeepCompressionV2Encoder = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=DeepCompressionV2Encoder
	if ($displaymodeTable.Component_Encoder_DeepCompressionV2Encoder -eq "Component_Encoder=DeepCompressionV2Encoder")
	{
	$Displaymode = "H.264 + Lossless"
	}
	
	#H.264 Compatibility Mode (ThinWire +)
    #Component_Encoder=CompatibilityEncoder
	$displaymodeTable.Component_Encoder_CompatibilityEncoder = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=CompatibilityEncoder
	if ($displaymodeTable.Component_Encoder_CompatibilityEncoder -eq "Component_Encoder=CompatibilityEncoder")
	{
	$Displaymode = "H.264 Compatibility Mode (ThinWire +)"
	}
		
	# Selective H.264 Is configured
	$displaymodeTable.Component_Encoder_Deprecated = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=Deprecated
	#Component_Encoder=Deprecated
	
		#fall back to H.264 Compatibility Mode (ThinWire +)
		# Auf Receiver selective nicht geht:
		$displaymodeTable.Component_VideoCodecUse_None = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_VideoCodecUse=None
		
		if ($displaymodeTable.Component_VideoCodecUse_None -eq "Component_VideoCodecUse=None")
		{
		$Displaymode = "Compatibility Mode (ThinWire +), selective H264 maybe not supported by Receiver)"
		}
			
		#Is used
		$displaymodeTable.Component_VideoCodecUse_Active = wmic /node:`'$machineDNS`' /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr 'Component_VideoCodecUse=For actively changing regions'			
		if ($displaymodeTable.Component_VideoCodecUse_Active -eq "Component_VideoCodecUse=For actively changing regions")
		{
		$Displaymode = "Selective H264"
		}

#Legacy Graphics
$displaymodeTable.LegacyGraphicsIsActive = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_graphics get /value | findstr IsActive=*
$displaymodeTable.Policy_LegacyGraphicsMode = wmic  /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_graphics get /value | findstr Policy_LegacyGraphicsMode=TRUE
if ($displaymodeTable.LegacyGraphicsIsActive -eq "IsActive=Active")
	{
	$Displaymode = "Legacy Graphics"
	}	

#DCR
$displaymodeTable.DcrIsActive = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_d3d get /value | findstr IsActive=*
$displaymodeTable.DcrAERO = wmic /node:`'$machineDNS`' /namespace:\\root\citrix\hdx path citrix_virtualchannel_d3d get /value | findstr Policy_AeroRedirection=*
if ($displaymodeTable.DcrAERO -eq "Policy_AeroRedirection=TRUE")
	{
	$Displaymode = "DCR"
	}
}
$tests.displaymode = "NEUTRAL", $displaymode




}
#-------------------------------------------------------------------------------------------------------------

  
  
" --- " | LogMe -display -progress
  
# Fill $tests into array if error occured OR $ShowOnlyErrorVDI = 0
# Check to see if the server is in an excluded folder path
if ($ExcludedCatalogs -contains $CatalogName) {
"$machineDNS in excluded folder - skipping" | LogMe -display -progress
}
else {
# Check if error exists on this vdi
if ($ShowOnlyErrorVDI -eq 0 ) { $allResults.$machineDNS = $tests }
else {
if ($ErrorVDI -gt 0) { $allResults.$machineDNS = $tests }
else { "$machineDNS is ok, no output into HTML-File" | LogMe -display -progress }
}
}
}
}
else { "Desktop Check skipped because ShowDesktopTable = 0 " | LogMe -display -progress }
  
# ======= XenApp Check ========
"Check XenApp Servers ####################################################################################" | LogMe -display -progress
" " | LogMe -display -progress
  
# Check XenApp only if $ShowXenAppTable is 1
#Skip2
if($ShowXenAppTable -eq 1 ) {
$allXenAppResults = @{}
$tests = @{}
$Catalogs = Get-BrokerCatalog -AdminAddress $AdminAddress
foreach ($Catalog in $Catalogs) {
  
  
  #Name of MachineCatalog
  $CatalogName = $Catalog | ForEach-Object{ $_.Name }

   if ($ExcludedCatalogs -like "*$CatalogName*" ) 
  { 
  "$CatalogName is excluded folder hence skipping" | LogMe -display -progress
    }
else 
{
"$CatalogName is available and processing" | LogMe -display -progress


  
#$XAmachines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress | Where-Object {$_.SessionSupport -eq "MultiSession" -and @(compare $_.tags $ExcludedTags -IncludeEqual | ? {$_.sideindicator -eq '=='}).count -eq 0}
$XAmachines = Get-BrokerMachine -MaxRecordCount $maxmachines -MachineName "*" -CatalogName $CatalogName -AdminAddress $AdminAddress | Where-Object {$_.SessionSupport -eq "MultiSession" -and @(Compare-Object $_.tags $ExcludedTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0}
$Maintenance = Get-CitrixMaintenanceInfo -AdminAddress $AdminAddress
  
foreach ($XAmachine in $XAmachines) {
$tests = @{}
  
# Column Name of Machine
$machineDNS = $XAmachine | ForEach-Object{ $_.DNSName }
"Machine: $machineDNS" | LogMe -display -progress
  
# Column CatalogNameName
$CatalogName = $XAmachine | ForEach-Object{ $_.CatalogName }
"Catalog: $CatalogName" | LogMe -display -progress
$tests.CatalogName = "NEUTRAL", $CatalogName

  
# Ping Machine
$result = Ping $machineDNS 100
if ($result -eq "SUCCESS") {
$tests.Ping = "SUCCESS", $result
  
#==============================================================================================
# Column Uptime (Query over WMI - only if Ping successfull)
$tests.WMI = "ERROR","Error"

$job = Start-Job -ScriptBlock $wmiOSBlock -ArgumentList $machineDNS

$wmi = Wait-job $job -Timeout 15 | Receive-Job

# Perform WMI related checks
    if ($null -ne $wmi) {
	    
        $tests.WMI = "SUCCESS", "Success"
	    
        $LBTime=[Management.ManagementDateTimeConverter]::ToDateTime($wmi.Lastbootuptime)
	    
        [TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)

	    if ($uptime.days -gt $maxUpTimeDays) {
		    "reboot warning, last reboot: {0:D}" -f $LBTime | LogMe -display -warning
		    
            $tests.Uptime = "WARNING", $uptime.days
	    }#If Uptime
        else {
		    
            $tests.Uptime = "SUCCESS", $uptime.days
	    }#Else Uptime

#==============================================================================================
#               CHECK CPU AND MEMORY USAGE 
#==============================================================================================

        # Check the AvgCPU value for 5 seconds
        $XAAvgCPUval = CheckCpuUsage ($machineDNS)
		#$VDtests.LoadBalancingAlgorithm = "SUCCESS", "LB is set to BEST EFFORT"} 
			
        if( [int] $XAAvgCPUval -lt 75) { "CPU usage is normal [ $XAAvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$XAAvgCPUval %" }
		elseif([int] $XAAvgCPUval -lt 85) { "CPU usage is medium [ $XAAvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$XAAvgCPUval %" }   	
		elseif([int] $XAAvgCPUval -lt 95) { "CPU usage is high [ $XAAvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$XAAvgCPUval %" }
		elseif([int] $XAAvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
        else { "CPU usage is Critical [ $XAAvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$XAAvgCPUval %" }   
		$XAAvgCPUval = 0

        # Check the Physical Memory usage       
        [int] $XAUsedMemory = CheckMemoryUsage ($machineDNS)
        if( [int] $XAUsedMemory -lt 75) { "Memory usage is normal [ $XAUsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$XAUsedMemory %" }
		elseif( [int] $XAUsedMemory -lt 85) { "Memory usage is medium [ $XAUsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$XAUsedMemory %" }   	
		elseif( [int] $XAUsedMemory -lt 95) { "Memory usage is high [ $XAUsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$XAUsedMemory %" }
		elseif( [int] $XAUsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" }
        else { "Memory usage is Critical [ $XAUsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$XAUsedMemory %" }   
		$XAUsedMemory = 0  

        foreach ($disk in $diskLettersWorkers)
        {
            # Check Disk Usage 
            $HardDisk = CheckHardDiskUsage -hostname $machineDNS -deviceID "$($disk):"
		    if ($null -ne $HardDisk) {	
			    $XAPercentageDS = $HardDisk.PercentageDS
			    $frSpace = $HardDisk.frSpace

			    If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
			    ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests.CFreespace = "ERROR", "Err" }
			    ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" } 
			    ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" }     
			    Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" }
			
			    $XAPercentageDS = 0
			    $frSpace = 0
			    $HardDisk = $null
		    }
		
        }

  
" --- " | LogMe -display -progress


    }#If WMI Not Null
    else {
	    
        "WMI connection failed - check WMI for corruption" | LogMe -display -error
	    
        #Original Line v1.4.4
        #stop-job $job

        #ALTERED LINES
        $JobID = $Job.Id

        Get-Job -Id $JobID | Remove-Job -Force

        $XAAvgCPUval = 101 | LogMe -error; $tests.AvgCPU = "ERROR", "Wmi Err"
    
        $XAUsedMemory = 101 | LogMe -error; $tests.MemUsg = "ERROR", "Wmi Err"
    
            foreach ($disk in $diskLettersWorkers){
        
                $XAPercentageDS = 0 | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "Wmi Err"

                $XAPercentageDS = 0
		        $frSpace = 0
		        $HardDisk = $null
                
            }#end of Foreach disk

    }#Else WMI Not Null
#----
  
# Column WriteCacheSize (only if Ping is successful)
################ PVS SECTION ###############
if (test-path "\\$machineDNS\c$\Personality.ini") {
    # Test if PVS cache is of type "device's hard drive"
    $PvsWriteCacheUNC = Join-Path "\\$machineDNS" ($PvsWriteCacheDrive+"$"+"\.vdiskcache")
    
    $CacheDiskOnHD = Test-Path $PvsWriteCacheUNC

if ($CacheDiskOnHD -eq $True) {
    
    $CacheDiskExists = $True
    
    $CachePVSType = "Device HD"
}
else{
  # Test if PVS cache is of type "device RAM with overflow to hard drive"
  $PvsWriteCacheUNC = Join-Path "\\$machineDNS" ($PvsWriteCacheDrive+"$"+"\vdiskdif.vhdx")
  
  $CacheDiskRAMwithOverflow = Test-Path $PvsWriteCacheUNC
  if ($CacheDiskRAMwithOverflow -eq $True) {
    
    $CacheDiskExists = $True
    
    $CachePVSType = "Device RAM with overflow to disk"
  }
  else {
    
    $CacheDiskExists = $False
    
    $CachePVSType = ""
  }
}

if ($CacheDiskExists -eq $True) {
    $CacheDisk = [long] ((get-childitem $PvsWriteCacheUNC -force).length)
    $CacheDiskGB = "{0:n2}GB" -f($CacheDisk / 1GB)
    "PVS Cache file size: {0:n2}GB" -f($CacheDisk / 1GB) | LogMe
    #"PVS Cache max size: {0:n2}GB" -f($PvsWriteMaxSizeInGB / 1GB) | LogMe -display
    $tests.WriteCacheType = "NEUTRAL", $CachePVSType
    if ($CacheDisk -lt ($PvsWriteMaxSizeInGB * 0.5)) {
        "WriteCache file size is low" | LogMe
        $tests.WriteCacheSize = "SUCCESS", $CacheDiskGB
    }
    elseif ($CacheDisk -lt ($PvsWriteMaxSizeInGB * 0.8)) {
        "WriteCache file size moderate" | LogMe -display -warning
        $tests.WriteCacheSize = "WARNING", $CacheDiskGB
    }
    else {
        "WriteCache file size is high" | LogMe -display -error
        $tests.WriteCacheSize = "ERROR", $CacheDiskGB
    }
}
$Cachedisk = 0
}
else {

    $tests.WriteCacheSize = "SUCCESS", "N/A"
}
############## END PVS SECTION #############
  
# Check services
$services = Get-Service -ComputerName $machineDNS
  
if (($services | Where-Object {$_.Name -eq "Spooler"}).Status -Match "Running") {
    "SPOOLER service running..." | LogMe
    $tests.Spooler = "SUCCESS","Success"
}
else {
    "SPOOLER service stopped" | LogMe -display -error
    $tests.Spooler = "ERROR","Error"
}
  
if (($services | Where-Object {$_.Name -eq "cpsvc"}).Status -Match "Running") {
    "Citrix Print Manager service running..." | LogMe
    $tests.CitrixPrint = "SUCCESS","Success"
}
else {
    "Citrix Print Manager service stopped" | LogMe -display -error
    $tests.CitrixPrint = "ERROR","Error"
}
 
 # Column OSBuild 
$MachineOSVersion = "N/A"
$MachineOSVersion = (Get-ItemProperty -Path "\\$machineDNS\C$\WINDOWS\System32\hal.dll" -ErrorAction SilentlyContinue).VersionInfo.FileVersion.Split()[0]
$tests.OSBuild = "NEUTRAL", $MachineOSVersion

  
}
else { $tests.Ping = "Neutral", $result }
#END of Ping-Section
  
# Column Serverload
$Serverload = $XAmachine | ForEach-Object{ $_.LoadIndex }
"Serverload: $Serverload" | LogMe -display -progress
if ($Serverload -ge $loadIndexError) { $tests.Serverload = "ERROR", $Serverload }
elseif ($Serverload -ge $loadIndexWarning) { $tests.Serverload = "WARNING", $Serverload }
else { $tests.Serverload = "SUCCESS", $Serverload }
  
# Column MaintMode
$MaintMode = $XAmachine | ForEach-Object{ $_.InMaintenanceMode }
"MaintenanceMode: $MaintMode" | LogMe -display -progress
if ($MaintMode) { 
	$objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $XAmachine.MachineName.ToUpper() } | Select-Object -First 1
	If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
	"MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
	$tests.MaintMode = "WARNING", $MaintenanceModeOn
	$ErrorVDI = $ErrorVDI + 1
}
else { $tests.MaintMode = "SUCCESS", "OFF" }
  
# Column RegState
$RegState = $XAmachine | ForEach-Object{ $_.RegistrationState }
"State: $RegState" | LogMe -display -progress
  
if ($RegState -ne "Registered") { $tests.RegState = "ERROR", $RegState }
else { $tests.RegState = "SUCCESS", $RegState }

# Column VDAVersion AgentVersion
$VDAVersion = $XAmachine | ForEach-Object{ $_.AgentVersion }
"VDAVersion: $VDAVersion" | LogMe -display -progress
$tests.VDAVersion = "NEUTRAL", $VDAVersion

# Column HostedOn - v1.4.4 Lines
#$HostedOn = $XAmachine | ForEach-Object{ $_.HostingServerName }
#"HostedOn: $HostedOn" | LogMe -display -progress
#$tests.HostedOn = "NEUTRAL", $HostedOn

# Column HostedOn 
$HostedOn = $XAmachine | ForEach-Object{ 
 if ($_.HostingServerName){

$_.HostingServerName.Split(".")[0]

}else{

"Not Known"

    }
}#end ForEach
"HostedOn: $HostedOn" | LogMe -display -progress
$tests.HostedOn = "NEUTRAL", $HostedOn

# Column Tags 
$Tags = $XAmachine | ForEach-Object{ $_.Tags }
"Tags: $Tags" | LogMe -display -progress
$tests.Tags = "NEUTRAL", $Tags
  
# Column ActiveSessions
$ActiveSessions = $XAmachine | ForEach-Object{ $_.SessionCount }
"Active Sessions: $ActiveSessions" | LogMe -display -progress
$tests.ActiveSessions = "NEUTRAL", $ActiveSessions

# Column ConnectedUsers
$ConnectedUsers = $XAmachine | ForEach-Object{ $_.AssociatedUserNames }
"Connected users: $ConnectedUsers" | LogMe -display -progress
$tests.ConnectedUsers = "NEUTRAL", $ConnectedUsers
  
# Column DeliveryGroup
$DeliveryGroup = $XAmachine | ForEach-Object{ $_.DesktopGroupName }
"DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
$tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup

# Column MCSImageOutOfDate
$MCSImageOutOfDate = $XAmachine | ForEach-Object{ $_.ImageOutOfDate }
"ImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -progress
if ($MCSImageOutOfDate -eq $true) { $tests.MCSImageOutOfDate = "ERROR", $MCSImageOutOfDate }
elseif ($MCSImageOutOfDate -eq $false) { $tests.MCSImageOutOfDate = "SUCCESS", $MCSImageOutOfDate  }
else { $tests.MCSImageOutOfDate = "NEUTRAL", $MCSImageOutOfDate }

  
# Check to see if the server is in an excluded folder path
if ($ExcludedCatalogs -contains $CatalogName) { "$machineDNS in excluded folder - skipping" | LogMe -display -progress }
else { $allXenAppResults.$machineDNS = $tests }
}
}
  }

  }#skip2end
  

else { "XenApp Check skipped because ShowXenAppTable = 0 or Farm is < V7.x " | LogMe -display -progress }
  
####################### Check END ####################################################################################" | LogMe -display -progress
# ======= Write all results to an html file =================================================
# Add Version of XenDesktop to EnvironmentName
$XDmajor, $XDminor = $controllerversion.Split(".")[0..1]
$XDVersion = "$XDmajor.$XDminor"
$EnvironmentNameOut = "$EnvironmentName $XDVersion"
$emailSubject = ("$EnvironmentNameOut Farm Report - " + $ReportDate)

Write-Host ("Saving results to html report: " + $resultsHTM)
writeHtmlHeader "$EnvironmentNameOut Farm Report" $resultsHTM

# Write Table with the Failures #FUTURE !!!!
#writeTableHeader $resultsHTM $CTXFailureFirstheaderName $CTXFailureHeaderNames $CTXFailureHeaderWidths $CTXFailureTableWidth
#$ControllerResults | sort-object -property XDControllerFirstheaderName | ForEach-Object{ writeData $CTXFailureResults $resultsHTM $CTXFailureFirstheaderName }
#writeTableFooter $resultsHTM

# Write Table with the Controllers
writeTableHeader $resultsHTM $XDControllerFirstheaderName $XDControllerHeaderNames $XDControllerHeaderWidths $XDControllerTableWidth
$ControllerResults | sort-object -property XDControllerFirstheaderName | ForEach-Object{ writeData $ControllerResults $resultsHTM $XDControllerHeaderNames }
writeTableFooter $resultsHTM

# Write Table with the License
writeTableHeader $resultsHTM $CTXLicFirstheaderName $CTXLicHeaderNames $CTXLicHeaderWidths $CTXLicTableWidth
$CTXLicResults | sort-object -property LicenseName | ForEach-Object{ writeData $CTXLicResults $resultsHTM $CTXLicHeaderNames }
writeTableFooter $resultsHTM
  
# Write Table with the Catalogs
writeTableHeader $resultsHTM $CatalogHeaderName $CatalogHeaderNames $CatalogWidths $CatalogTablewidth
$CatalogResults | ForEach-Object{ writeData $CatalogResults $resultsHTM $CatalogHeaderNames}
writeTableFooter $resultsHTM
  
  
# Write Table with the Assignments (Delivery Groups)
writeTableHeader $resultsHTM $AssigmentFirstheaderName $vAssigmentHeaderNames $vAssigmentHeaderWidths $Assigmenttablewidth
$AssigmentsResults | sort-object -property ReplState | ForEach-Object{ writeData $AssigmentsResults $resultsHTM $vAssigmentHeaderNames }
writeTableFooter $resultsHTM

# Write Table with all XenApp Servers
if ($ShowXenAppTable -eq 1 ) {
writeTableHeader $resultsHTM $XenAppFirstheaderName $XenAppHeaderNames $XenAppHeaderWidths $XenApptablewidth
$allXenAppResults | sort-object -property CatalogName | ForEach-Object{ writeData $allXenAppResults $resultsHTM $XenAppHeaderNames }
writeTableFooter $resultsHTM
}
else { "No XenApp output in HTML " | LogMe -display -progress }

# Write Table with all Desktops
if ($ShowDesktopTable -eq 1 ) {
writeTableHeader $resultsHTM $VDIFirstheaderName $VDIHeaderNames $VDIHeaderWidths $VDItablewidth
$allResults | sort-object -property CatalogName | ForEach-Object{ writeData $allResults $resultsHTM $VDIHeaderNames }
writeTableFooter $resultsHTM
}
else { "No XenDesktop output in HTML " | LogMe -display -progress }
  
 
writeHtmlFooter $resultsHTM

$scriptend = Get-Date
$scriptruntime =  $scriptend - $scriptstart | Select-Object TotalSeconds
$scriptruntimeInSeconds = $scriptruntime.TotalSeconds
#Write-Host $scriptruntime.TotalSeconds
"Script was running for $scriptruntimeInSeconds " | LogMe -display -progress

#Only Send Email if Variable in XML file is equal to 1
if ($CheckSendMail -eq 1){

#send email
$emailMessage = New-Object System.Net.Mail.MailMessage
$emailMessage.From = $emailFrom
$emailMessage.To.Add( $emailTo )
$emailMessage.CC.Add( $emailCC )
$emailMessage.Subject = $emailSubject 
$emailMessage.IsBodyHtml = $true
$emailMessage.Body = (Get-Content $resultsHTM) | Out-String
$emailMessage.Attachments.Add($resultsHTM)
$emailMessage.Priority = ($emailPrio)

$smtpClient = New-Object System.Net.Mail.SmtpClient( $smtpServer , $smtpServerPort )
$smtpClient.EnableSsl = $smtpEnableSSL

# If you added username an password, add this to smtpClient
If ((![string]::IsNullOrEmpty($smtpUser)) -and (![string]::IsNullOrEmpty($smtpPW))){
	$pass = $smtpPW | ConvertTo-SecureString -key $smtpKey
	$cred = New-Object System.Management.Automation.PsCredential($smtpUser,$pass)

	$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($cred.Password)
	$smtpUserName = $cred.Username
	$smtpPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)

	$smtpClient.Credentials = New-Object System.Net.NetworkCredential( $smtpUserName , $smtpPassword );
}

$smtpClient.Send( $emailMessage )



}#end of IF CheckSendMail
else{

    "XenApp Check skipped because CheckSendMail = 0" | LogMe -display -progress

}#Skip Send Mail


#=========== History ===========================================================================
# Version 0.6
# Edited on December 15th 2014 by Sebastiaan Brozius (NORISK IT Groep)
# - Added MaintenanceMode-column for XenApp-servers, Fixed some typos, Cleaned up layout
#
# Version 0.7
# Edited on December 15th 2014 by Sebastiaan Brozius (NORISK IT Groep)
# - Added ServerLoad-column for XenApp-servers
# - Added loadIndexError and loadIndexWarning-variables which are used as thresholds for the ServerLoad-column
# - Removed line $machines = Get-BrokerMachine | Where-Object {$_.SessionSupport -eq "SingleSession"} from the XenApp-section (because this is for VDI)
# - Removed Column-positions from comments
#
# Version 0.8
# Edited on January 17th 2015 by Sacha Thomet
# - Added MaintenanceMode-column for Assignments (DeliveryGroups), added MaintenanceMode-column for VDIs
#
# Version 0.9
# Edited on May 5th 2015 by Sacha Thomet
# - Changed some wording terms into account of XenDesktop/XenApp 7.x
# (Assignment to DeliveryGroup, AssignedCount to AssignedToUser, UsedCount to AssignedToDG,
# UnassignedCount to NotToUserAssigned)
# - Added some columns the MachineCatalog (ProvisioningType, AllocationType)
# - Added some columns the DeliveryGroups (PublishedName, DesktopKind)
#
# Version 0.92
# Edited on July 2015 by Sacha Thomet
# - Adjusted some typo
# - Delivery-Group-Section: XenDesktop show Total Desktops assigned to this group (FeatureRequest Luis G)
# - Delivery-Group-Section: XenDesktop Filter out shared Desktops on count for free Desktops (FeatureRequest James)
#
# Version 0.93
# Edited on August 2015 by Stefan Beckmann (Unico Data AG)
# - Added loadIndexError and loadIndexWarning as variable in the config section
# - Specifying the PowerShell SnapIn which must be loaded
# - You can now specify the delivery controller. This makes it possible to execute the script on a system which is not Delivery Controller or VDA.
# - Added the connected users in the table
# - Remove the old report bevor main script starts
# - Log Citrix PS Snapins
# - Run Citrix Cmdlet with AdminAdress to run this script remote
# - If you run script remote, now returned correctly the version of XenDesktop. In the first Get-BrokerController command I added the parameter DNSName.
# - New Way to send mail, which also allows passwords
# - To send a mail without user and password I could not test
#
# # Version 0.94
# Edited on October 2015 by Sacha Thomet
# - Removed PVS WriteCache column - if you need that check this on PVS Health Check. XenApp/XenDesktop is maybe provisioned without PVS. 
# - Add possibility to set a Mail-Priority
#
# # Version 0.95
# Edited on May 2016 by Sacha Thomet
# - Check CPU, Memory and C: of Controllers  
# - XenApp: Add values: CPU & Memory and Disk Usage 
# - XenApp: Option to toggle on/off to show Connected Users 
# - XenApp: DesktopFree set to N/A because not relevant
#
# ToDo in a further version (Correction & Feature Requests)
# - XenApp:  Add more relevant values: Number of active users per server / Logon Enabled / SessionSupport 
# - XenDesktop: show Connected User 
# - XenApp: Get-BrokerSession per Machine for unique Sessions (?)(Feature Request Majeed Attar) => need better specification of the need.
#
# # Version 0.96
# Edited on May 2016 by Sacha Thomet
#  Added D: for Controller and XenApp Server
#
# # Version 0.99
# Edited on September 2016 by Sacha Thomet
# - Show PowerState and do some checks not on powered off maschone (Ping, RegistrationState)
# - Show VDA Version of XA or XD
# - Show Host
#
# # Version 0.992
# Edited on September 2016 by Stefan Beckmann
# - Variable $EnvironmentName adjusted without version, and joined in the script later
# - Variable $PvsWriteCache and $PvsWriteMaxSize re-added, since it is still used in the script
# - Created timestamp report as variable $ReportDate
# - $EnvironmentName and $emailSubject redefined in the report generation, incl. XenDesktop version with the new variable $XDmajor, $XDminor and $XDVersion
#
# # Version 0.995
# Edited on September 2016 by Stefan Beckmann
# - Configuration via an XML file
# - Redefined display the date for the report
# - Replaced generate the date in the second place by variable
#
# # Version 0.996 - 1.00
# Edited on September 2016 by Sacha Thomet
# - minor bug fixes
#
# # Version 1.0.1
# Edited on September 2016 by Tyron Scholem
# - localization correction for systems with decimal separator of ","
#
# # Version 1.1.0
# Edited on September 2016 by Tyron Scholem
# - added uptime information for Delivery Controllers
#
# # Version 1.2.0
# Edited on September 2016 by Sacha Thomet
# - determine Graphic Mode
#
# # Version 1.2.1 / 1.2.2
# Edited on September 2016 by Sacha Thomet
# - speed improvement for Graphic Mode - just check used desktops
#
# # Version 1.2.3
# Edited on November 2016 by Stefan Beckmann
# - Add function to get informations about the user who set maintenance mode
# - WinRM needed if script run remote
# # Version 1.2.4
# Edited on December 2016 by Sacha (Input by https://github.com/sommerers)
# - Bugfix on DFreespace is on N/A even if exists
#
# # Version 1.2.5
# descending order for the XenApp XenDesktop tables (sort-object -property CatalogName)
# Info about running the Script as Scheduled task
#
# # Version 1.2.6
# Edited on February 2017 by Sacha Thomet
# - Minor changes in description of the columns (DesktopGroup is now DelvieryGroup similar like Studio)
# - new column for DeliveryGroup in Desktops-Table
# - new column for SessionSupport in DeliveryGroup-Table
# - Some new additional Farm infos in footer (DB,LHC, ConnectionLeasing, LicenseServer)
# - Fix of WriteCache in Desktop section (still not 100% ok ... Sorry!)
# 
# # Version 1.2.7
# Edited on March 2017 by mikekacz
# - Fix required to work on non-controller server, like on Studio server. -adminaddress needs to be 
#   setup only once, as it is kept in PoSh session
# - Exclusions by tags. PS1 and XML changed! 
# - Citrix License report 
#
# # Version 1.2.8
# Edited on April 2017 by Sacha Thomet
# - Bugfixes:
#   * Replace $PvsWriteMaxSize with $PvsWriteMaxSizeInGB (PvsWriteMaxSize is unique to take it from XML)
#   * Replace $EnvironmentName with $EnvironmentNameOut ($EnvironmentName is unique to take it from XML)
#   * Replace $ShowXenAppTable with $ShowXenAppTable ($ShowXenAppTable is unique to take it from XML)
#
# # Version 1.2.9
# Edited on April 2017 by mikekacz
# - Added disk to check selector in config file
#
# # Version 1.3
# Edited on June 2017 by Sacha
# - Added column with tags on VDI & VDA Server
# - Added column with MinimumFunctionalLevel on Catalogs and DeliveryGroups
# # Version 1.3.1-1.3.4
# Edited on July 2017 by Sacha
# - OS Build in a Column
# - Bugfixes
# - On Table delivery group, Error if a DG with type "MultiSession" has ShutdownDesktopsAfterUse on true
#
# # Version 1.3.5
# Edited on May 2018 by Sacha
# - improvement for the excluded Catalogs (Thank you Im-Saravana, https://github.com/Im-Saravana)
# - Added output of the Runtime (Script start - scriptend)
#
# # Version 1.3.6
# Edited on September 2018 by Stefan
# - The command changed from wmic /node:$machineDNS to wmic /node:`'$machineDNS`'. That supports dashes in hostname.
#
# # Version 1.3.7
# Version changes by M.Löffler on Oct 2018
# added column "LastConnect" to variable "$VDIHeaderNames" Report Table around line 149
# added column "4" to variable "$VDIHeaderWidths" to format new column "LastConnect" in VDI Report Table around line 150
# added Scriptblock "## Column LastConnect" around line 1108
#
#  1.3.7.1
# Version changes by S.Thomet, removed a lot of Alias according suggestion of VSCode
#
#  1.3.8
# - Implement Issue/Idea #70 show Hypervisorconnection
#
#  1.3.9
# - Changed the way how to gather the AvgCGPU for more comptaibility.
#  1.3.9.1
# - Bugfix for licesne reading.  
# 
#  1.4 Enable MCS Features 
# - Add MCSImageOutOfDate (PendingUpdate) Column for Desktops & Apps
# - 1.4.2, Bugfix for https://github.com/sacha81/XA-and-XD-HealthCheck/issues/73 
# - 1.4.4, Column EDT_MTU added für virtual Desktops 
# - 1.4.5, Enabled to work for Citrix Cloud (see also changes in XML!) 
# - 1.4.6, Ping to Desktop or AppServer is no more red because not critical
#
# == FUTURE ==
# #  1.5
# Version changes by S.Thomet
# - CREATE Proper functions
#
# #  1.5.1
# Version changes by S.Thomet
# - Implement Idea #27 from GitHub: Fail-Rate in %
#
#=========== History END ===========================================================================
