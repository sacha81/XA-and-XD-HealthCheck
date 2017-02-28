#==============================================================================================
# Created on: 11.2014 Version: 1.2.6
# Created by: Sacha / sachathomet.ch & Contributers (see changelog)
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
# Call by : Manual or by Scheduled Task, e.g. once a day
#           !! If you run it as scheduled task you need to add with argument “non interactive” 
#           or your user has interactive persmission!
#
# Code History at the end of the file
#
#==============================================================================================

#Don't change below here if you don't know what you are doing ... 
#==============================================================================================
# Load only the snap-ins, which are used
if ((Get-PSSnapin "Citrix.Broker.Admin.*" -EA silentlycontinue) -eq $null) {
try { Add-PSSnapin Citrix.Broker.Admin.* -ErrorAction Stop }
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
	$cfg.Settings.Variables.Variable | foreach {
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


$PvsWriteMaxSize = $PvsWriteMaxSize * 1Gb

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
$logfile = Join-Path $currentDir ("CTXXDHealthCheck.log")
$resultsHTM = Join-Path $currentDir ("CTXXDHealthCheck.htm")
  
#Header for Table "XD/XA Controllers" Get-BrokerController
$XDControllerFirstheaderName = "ControllerServer"
$XDControllerHeaderNames = "Ping", 	"State","DesktopsRegistered", 	"ActiveSiteServices", 	"CFreespace", 	"DFreespace", 	"AvgCPU", 	"MemUsg", 	"Uptime"
$XDControllerHeaderWidths = "2",	"2", 	"2", 					"10",					"4",			"4",			"4",		"4",		"4"
$XDControllerTableWidth= 1200
  
#Header for Table "MachineCatalogs" Get-BrokerCatalog
$CatalogHeaderName = "CatalogName"
$CatalogHeaderNames = 	"AssignedToUser", 	"AssignedToDG", "NotToUserAssigned","ProvisioningType", "AllocationType"
$CatalogWidths = 		"4",				"8", 			"8", 				"8", 				"8"
$CatalogTablewidth = 900
  
#Header for Table "DeliveryGroups" Get-BrokerDesktopGroup
$AssigmentFirstheaderName = "DeliveryGroup"
$vAssigmentHeaderNames = 	"PublishedName","DesktopKind", "SessionSupport", 	"TotalMachines","DesktopsAvailable","DesktopsUnregistered", "DesktopsInUse","DesktopsFree", "MaintenanceMode"
$vAssigmentHeaderWidths = 	"4", 			"4", 			"4", 	"4", 		"4", 				"4", 					"4", 			"4", 			"2"
$Assigmenttablewidth = 900
  
#Header for Table "VDI Checks" Get-BrokerMachine
$VDIfirstheaderName = "Desktop-Name"

$VDIHeaderNames = "CatalogName","DeliveryGroup","PowerState", "Ping", "MaintMode", 	"Uptime", 	"RegState","VDAVersion","AssociatedUserNames",  "WriteCacheType", "WriteCacheSize", "HostedOn", "displaymode"
$VDIHeaderWidths = "4", "4",		"4","4", 	"4", 				"4", 		"4", 				"4",			  "4",			  "4",			  "4",			  "4", "4"

$VDItablewidth = 1200
  
#Header for Table "XenApp Checks" Get-BrokerMachine
$XenAppfirstheaderName = "XenApp-Server"
if ($ShowConnectedXenAppUsers -eq "1") { 

	$XenAppHeaderNames = "CatalogName", "DeliveryGroup", "Serverload", 	"Ping", "MaintMode","Uptime", 	"RegState", "VDAVersion", "Spooler", 	"CitrixPrint",  "CFreespace", 	"DFreespace", 	"AvgCPU", 	"MemUsg", 	"ActiveSessions",  "WriteCacheType", "WriteCacheSize", "ConnectedUsers" , "HostedOn"
	$XenAppHeaderWidths = "4", 			"4", 				"4", 			"4", 	"4", 		"4", 		"4", 		"4", 		"4", 			"4",			"4",			"4",		"4",		"4",			  "4",			"4",			"4",			"4",			"4"
}
else { 
	$XenAppHeaderNames = "CatalogName",  "DeliveryGroup", "Serverload", 	"Ping", "MaintMode","Uptime", 	"RegState", "Spooler", 	"CitrixPrint", 	"CFreespace", 	"DFreespace", 	"AvgCPU", 	"MemUsg", 	"ActiveSessions", "VDAVersion", "WriteCacheType", "WriteCacheSize", "HostedOn"
	$XenAppHeaderWidths = "4", 			"4", 				"4", 			"4", 	"4", 		"4", 		"4", 		"4", 		"4", 			"4",			"4",			"4",		"4",		"4",			  "4",			"4",			"4",			"4"

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
	Try { $CpuUsage=(get-counter -ComputerName $hostname -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 5 -ErrorAction Stop | select -ExpandProperty countersamples | select -ExpandProperty cookedvalue | Measure-Object -Average).average
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
        if ($HardDisk -ne $null)
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
$data.Keys | sort | foreach {
$tableEntry += "<tr>"
$computerName = $_
$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")
#$data.$_.Keys | foreach {
$headerNames | foreach {
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
			if ((Get-PSSnapin "Get-PSSnapin Citrix.ConfigurationLogging.Admin.*" -ErrorAction silentlycontinue) -eq $null) {
				try { Add-PSSnapin Citrix.ConfigurationLogging.Admin.* -ErrorAction Stop } catch { write-error "Error Get-PSSnapin Citrix.ConfigurationLogging.Admin.* Powershell snapin"; Return }
			} #If
			
			$Date = Get-Date
			$StartDate = $Date.AddDays(-7) # Hard coded value for how many days back
			$EndDate = $Date
			
			# Command to get the informations from log
			$LogEntrys = Get-LogLowLevelOperation -MaxRecordCount 1000000 -Filter { StartTime -ge $StartDate -and EndTime -le $EndDate } | Where { $_.Details.PropertyName -eq 'MAINTENANCEMODE' } | Sort EndTime -Descending
			
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
  try { $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $computer }
  catch { $wmi = $null }
  return $wmi
}

#==============================================================================================
# == MAIN SCRIPT ==
#==============================================================================================
rm $logfile -force -EA SilentlyContinue
rm $resultsHTM -force -EA SilentlyContinue
  
"#### Begin with Citrix XenDestop / XenApp HealthCheck ######################################################################" | LogMe -display -progress
  
" " | LogMe -display -progress

# get some farm infos, which will be presented in footer 
$dbinfo = Get-BrokerDBConnection
$brokersiteinfos = Get-BrokerSite
$lsname = $brokersiteinfos.LicenseServerName
$lsport = $brokersiteinfos.LicenseServerPort
$CLeasing = $brokersiteinfos.ConnectionLeasingEnabled
$LHC =$brokersiteinfos.LocalHostCacheEnabled


# Log the loaded Citrix PS Snapins
(Get-PSSnapin "Citrix.*" -EA silentlycontinue).Name | ForEach {"PSSnapIn: " + $_ | LogMe -display -progress}
  
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
  $ShowXenAppTable = 0
} else { 
  "XenDesktop/XenApp Version above 7.x ($controllerversion) - XenApp and DesktopCheck will be performed" | LogMe -display -progress
}

foreach ($Controller in $Controllers) {
$tests = @{}
  
#Name of $Controller
$ControllerDNS = $Controller | %{ $_.DNSName }
"Controller: $ControllerDNS" | LogMe -display -progress
  
#Ping $Controller
$result = Ping $ControllerDNS 100
if ($result -ne "SUCCESS") { $tests.Ping = "Error", $result }
else { $tests.Ping = "SUCCESS", $result 

#Now when Ping is ok also check this:
  
#State of this controller
$ControllerState = $Controller | %{ $_.State }
"State: $ControllerState" | LogMe -display -progress
if ($ControllerState -ne "Active") { $tests.State = "ERROR", $ControllerState }
else { $tests.State = "SUCCESS", $ControllerState }
  
#DesktopsRegistered on this controller
$ControllerDesktopsRegistered = $Controller | %{ $_.DesktopsRegistered }
"Registered: $ControllerDesktopsRegistered" | LogMe -display -progress
$tests.DesktopsRegistered = "NEUTRAL", $ControllerDesktopsRegistered
  
#ActiveSiteServices on this controller
$ActiveSiteServices = $Controller | %{ $_.ActiveSiteServices }
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

        # Check C Disk Usage 
		$HardDisk = CheckHardDiskUsage -hostname $ControllerDNS -deviceID "C:"
		if ($HardDisk -ne $null) {	
			$XAPercentageDS = $HardDisk.PercentageDS
			$frSpace = $HardDisk.frSpace
			
	        If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests.CFreespace = "SUCCESS", "$frSpace GB" } 
			ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests.CFreespace = "ERROR", "Err" }
			ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests.CFreespace = "ERROR", "$frSpace GB" } 
			ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests.CFreespace = "WARNING", "$frSpace GB" }     
	        Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests.CFreespace = "ERROR", "$frSpace GB" }  
        
			$XAPercentageDS = 0
			$frSpace = 0
			$HardDisk = $null
		}

		$tests.DFreespace = "NEUTRAL", "N/A" 
		if ( $ControllerHaveD -eq "1" ) {
			# Check D Disk Usage on DeliveryController
	        $HardDiskd = CheckHardDiskUsage -hostname $ControllerDNS -deviceID "D:"
			if ($HardDiskd -ne $null)
			{
				$XAPercentageDSd = $HardDiskd.PercentageDS
				$frSpaced = $HardDiskd.frSpace

				If ( [int] $XAPercentageDSd -gt 15) { "D: Disk Free is normal [ $XAPercentageDSd % ]" | LogMe -display; $tests.DFreespace = "SUCCESS", "$frSpaced GB" } 
				ElseIf ([int] $XAPercentageDSd -eq 0) { "D: Disk Free test failed" | LogMe -error; $tests.DFreespace = "ERROR", "Err" }
				ElseIf ([int] $XAPercentageDSd -lt 5) { "D: Disk Free is Critical [ $XAPercentageDSd % ]" | LogMe -error; $tests.DFreespace = "ERROR", "$frSpaced GB" } 
				ElseIf ([int] $XAPercentageDSd -lt 15) { "D: Disk Free is Low [ $XAPercentageDSd % ]" | LogMe -warning; $tests.DFreespace = "WARNING", "$frSpaced GB" }     
				Else { "D: Disk Free is Critical [ $XAPercentageDSd % ]" | LogMe -error; $tests.DFreespace = "ERROR", "$frSpaced GB" }  
				
				$XAPercentageDSd = 0
				$frSpaced = 0
				$HardDiskd = $null
			}
		}
		
		# Check uptime (Query over WMI)
    $tests.WMI = "ERROR","Error"
    try { $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $ControllerDNS }
    catch { $wmi = $null }

    # Perform WMI related checks
    if ($wmi -ne $null) {
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
  $CatalogName = $Catalog | %{ $_.Name }
  "Catalog: $CatalogName" | LogMe -display -progress

  if ($ExcludedCatalogs -contains $CatalogName) {
    "Excluded Catalog, skipping" | LogMe -display -progress
  } else {
    #CatalogAssignedCount
    $CatalogAssignedCount = $Catalog | %{ $_.AssignedCount }
    "Assigned: $CatalogAssignedCount" | LogMe -display -progress
    $tests.AssignedToUser = "NEUTRAL", $CatalogAssignedCount
  
    #CatalogUnassignedCount
    $CatalogUnAssignedCount = $Catalog | %{ $_.UnassignedCount }
    "Unassigned: $CatalogUnAssignedCount" | LogMe -display -progress
    $tests.NotToUserAssigned = "NEUTRAL", $CatalogUnAssignedCount
  
    # Assigned to DeliveryGroup
    $CatalogUsedCountCount = $Catalog | %{ $_.UsedCount }
    "Used: $CatalogUsedCountCount" | LogMe -display -progress
    $tests.AssignedToDG = "NEUTRAL", $CatalogUsedCountCount
  
     #ProvisioningType
     $CatalogProvisioningType = $Catalog | %{ $_.ProvisioningType }
     "ProvisioningType: $CatalogProvisioningType" | LogMe -display -progress
     $tests.ProvisioningType = "NEUTRAL", $CatalogProvisioningType
  
     #AllocationType
     $CatalogAllocationType = $Catalog | %{ $_.AllocationType }
     "AllocationType: $CatalogAllocationType" | LogMe -display -progress
     $tests.AllocationType = "NEUTRAL", $CatalogAllocationType
  
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
  $DeliveryGroup = $Assigment | %{ $_.Name }
  "DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
  
  if ($ExcludedCatalogs -contains $DeliveryGroup) {
    "Excluded Delivery Group, skipping" | LogMe -display -progress
  } else {
  
    #PublishedName
    $AssigmentDesktopPublishedName = $Assigment | %{ $_.PublishedName }
    "PublishedName: $AssigmentDesktopPublishedName" | LogMe -display -progress
    $tests.PublishedName = "NEUTRAL", $AssigmentDesktopPublishedName
  
    #DesktopsTotal
    $TotalDesktops = $Assigment | %{ $_.TotalDesktops }
    "DesktopsAvailable: $TotalDesktops" | LogMe -display -progress
    $tests.TotalMachines = "NEUTRAL", $TotalDesktops
  
    #DesktopsAvailable
    $AssigmentDesktopsAvailable = $Assigment | %{ $_.DesktopsAvailable }
    "DesktopsAvailable: $AssigmentDesktopsAvailable" | LogMe -display -progress
    $tests.DesktopsAvailable = "NEUTRAL", $AssigmentDesktopsAvailable
  
    #DesktopKind
    $AssigmentDesktopsKind = $Assigment | %{ $_.DesktopKind }
    "DesktopKind: $AssigmentDesktopsKind" | LogMe -display -progress
    $tests.DesktopKind = "NEUTRAL", $AssigmentDesktopsKind
	
	#SessionSupport
	$SessionSupport = $Assigment | %{ $_.SessionSupport }
	"SessionSupport: $SessionSupport" | LogMe -display -progress
    $tests.SessionSupport = "NEUTRAL", $SessionSupport
	
	if ($SessionSupport -eq "MultiSession" ) { 
	
	$tests.DesktopsFree = "NEUTRAL", "N/A"
	$tests.DesktopsInUse = "NEUTRAL", "N/A"
		
	}
    else { 
			#DesktopsInUse
			$AssigmentDesktopsInUse = $Assigment | %{ $_.DesktopsInUse }
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
    $AssigmentDesktopsinMaintenanceMode = $Assigment | %{ $_.inMaintenanceMode }
    "inMaintenanceMode: $AssigmentDesktopsinMaintenanceMode" | LogMe -display -progress
    if ($AssigmentDesktopsinMaintenanceMode) { $tests.MaintenanceMode = "WARNING", "ON" }
    else { $tests.MaintenanceMode = "SUCCESS", "OFF" }
  
    #DesktopsUnregistered
    $AssigmentDesktopsUnregistered = $Assigment | %{ $_.DesktopsUnregistered }
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
  
# ======= Desktop Check ========
"Check virtual Desktops ####################################################################################" | LogMe -display -progress
" " | LogMe -display -progress
  
if($ShowDesktopTable -eq 1 ) {
  
$allResults = @{}
  
$machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress| Where-Object {$_.SessionSupport -eq "SingleSession"}
  
# SessionSupport only availiable in XD 7.x - for this reason only distinguish in Version above 7 if Desktop or XenApp
if($controllerversion -lt 7 ) { $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress}
else { $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress| Where-Object {$_.SessionSupport -eq "SingleSession" } }

$Maintenance = Get-CitrixMaintenanceInfo -AdminAddress $AdminAddress

foreach($machine in $machines) {
$tests = @{}
  
$ErrorVDI = 0
  
# Column Name of Desktop
$machineDNS = $machine | %{ $_.DNSName }
"Machine: $machineDNS" | LogMe -display -progress
  
# Column CatalogName
$CatalogName = $machine | %{ $_.CatalogName }
"Catalog: $CatalogName" | LogMe -display -progress
$tests.CatalogName = "NEUTRAL", $CatalogName

# Column DeliveryGroup
$DeliveryGroup = $machine | %{ $_.DesktopGroupName }
"DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
$tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup

# Column Powerstate
$Powered = $machine | %{ $_.PowerState }
"PowerState: $Powered" | LogMe -display -progress
$tests.PowerState = "NEUTRAL", $Powered

if ($Powered -eq "Off" -OR $Powered -eq "Unknown") {
$tests.PowerState = "NEUTRAL", $Powered
}

if ($Powered -eq "On") {
$tests.PowerState = "SUCCESS", $Powered
}

if ($Powered -eq "On" -OR $Powered -eq "Unknown") {


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
  if ($wmi -ne $null) {
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
#"PVS Cache max size: {0:n2}GB" -f($PvsWriteMaxSize / 1GB) | LogMe -display
$tests.WriteCacheType = "NEUTRAL", $CachePVSType
if ($CacheDisk -lt ($PvsWriteMaxSize * 0.5)) {
"WriteCache file size is low" | LogMe
$tests.WriteCacheSize = "SUCCESS", $CacheDiskGB
}
elseif ($CacheDisk -lt ($PvsWriteMaxSize * 0.8)) {
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


#---------------------
  
  }
else {
$tests.Ping = "Error", $result
$ErrorVDI = $ErrorVDI + 1
}
#END of Ping-Section

# Column RegistrationState
$RegistrationState = $machine | %{ $_.RegistrationState }
"State: $RegistrationState" | LogMe -display -progress
if ($RegistrationState -ne "Registered") {
$tests.RegState = "ERROR", $RegistrationState
$ErrorVDI = $ErrorVDI + 1
}
else { $tests.RegState = "SUCCESS", $RegistrationState }

} 
 
# Column MaintenanceMode
$MaintenanceMode = $machine | %{ $_.InMaintenanceMode }
"MaintenanceMode: $MaintenanceMode" | LogMe -display -progress
if ($MaintenanceMode) {
	$objMaintenance = $Maintenance | Where { $_.TargetName.ToUpper() -eq $machine.MachineName.ToUpper() } | Select -First 1
	If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
	"MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
	$tests.MaintMode = "WARNING", $MaintenanceModeOn
	$ErrorVDI = $ErrorVDI + 1
}
else { $tests.MaintMode = "SUCCESS", "OFF" }
  
# Column HostedOn 
$HostedOn = $machine | %{ $_.HostingServerName }
"HostedOn: $HostedOn" | LogMe -display -progress
$tests.HostedOn = "NEUTRAL", $HostedOn

# Column VDAVersion AgentVersion
$VDAVersion = $machine | %{ $_.AgentVersion }
"VDAVersion: $VDAVersion" | LogMe -display -progress
$tests.VDAVersion = "NEUTRAL", $VDAVersion

# Column AssociatedUserNames
$AssociatedUserNames = $machine | %{ $_.AssociatedUserNames }
"Assigned to $AssociatedUserNames" | LogMe -display -progress
$tests.AssociatedUserNames = "NEUTRAL", $AssociatedUserNames







# Column displaymode when a User has a Session
$sessionUser = $machine | %{ $_.SessionUserName }

$displaymode = "N/A"
if ( $ShowGraphicsMode -eq "1" ) {

if ($sessionUser -notlike "" )
{

$displaymode = "unknown"
$displaymodeTable = @{}


#H264
$displaymodeTable.H264Active = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr IsActive=*

    # H.264 Pure
    #Component_Encoder=DeepCompressionV2Encoder	
	$displaymodeTable.Component_Encoder_DeepCompressionEncoder = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=DeepCompressionEncoder
	if ($displaymodeTable.Component_Encoder_DeepCompressionEncoder -eq "Component_Encoder=DeepCompressionEncoder")
	{
	$Displaymode = "Pure H.264"
	}
	
	# Thinwire H.264 + Lossless (true native H264)
    #Component_Encoder=DeepCompressionV2Encoder
	$displaymodeTable.Component_Encoder_DeepCompressionV2Encoder = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=DeepCompressionV2Encoder
	if ($displaymodeTable.Component_Encoder_DeepCompressionV2Encoder -eq "Component_Encoder=DeepCompressionV2Encoder")
	{
	$Displaymode = "H.264 + Lossless"
	}
	
	#H.264 Compatibility Mode (ThinWire +)
    #Component_Encoder=CompatibilityEncoder
	$displaymodeTable.Component_Encoder_CompatibilityEncoder = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=CompatibilityEncoder
	if ($displaymodeTable.Component_Encoder_CompatibilityEncoder -eq "Component_Encoder=CompatibilityEncoder")
	{
	$Displaymode = "H.264 Compatibility Mode (ThinWire +)"
	}
		
	# Selective H.264 Is configured
	$displaymodeTable.Component_Encoder_Deprecated = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_Encoder=Deprecated
	#Component_Encoder=Deprecated
	
		#fall back to H.264 Compatibility Mode (ThinWire +)
		# Auf Receiver selective nicht geht:
		$displaymodeTable.Component_VideoCodecUse_None = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr Component_VideoCodecUse=None
		
		if ($displaymodeTable.Component_VideoCodecUse_None -eq "Component_VideoCodecUse=None")
		{
		$Displaymode = "Compatibility Mode (ThinWire +), selective H264 maybe not supported by Receiver)"
		}
			
		#Is used
		$displaymodeTable.Component_VideoCodecUse_Active = wmic /node:$machineDNS /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_thinwire get /value | findstr 'Component_VideoCodecUse=For actively changing regions'			
		if ($displaymodeTable.Component_VideoCodecUse_Active -eq "Component_VideoCodecUse=For actively changing regions")
		{
		$Displaymode = "Selective H264"
		}

#Legacy Graphics
$displaymodeTable.LegacyGraphicsIsActive = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_graphics get /value | findstr IsActive=*
$displaymodeTable.Policy_LegacyGraphicsMode = wmic  /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_graphics get /value | findstr Policy_LegacyGraphicsMode=TRUE
if ($displaymodeTable.LegacyGraphicsIsActive -eq "IsActive=Active")
	{
	$Displaymode = "Legacy Graphics"
	}	

#DCR
$displaymodeTable.DcrIsActive = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_d3d get /value | findstr IsActive=*
$displaymodeTable.DcrAERO = wmic /node:$machineDNS /namespace:\\root\citrix\hdx path citrix_virtualchannel_d3d get /value | findstr Policy_AeroRedirection=*
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
if($ShowXenAppTable -eq 1 ) {
$allXenAppResults = @{}
  
$XAmachines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress | Where-Object {$_.SessionSupport -eq "MultiSession"}

$Maintenance = Get-CitrixMaintenanceInfo -AdminAddress $AdminAddress
  
foreach ($XAmachine in $XAmachines) {
$tests = @{}
  
# Column Name of Machine
$machineDNS = $XAmachine | %{ $_.DNSName }
"Machine: $machineDNS" | LogMe -display -progress
  
# Column CatalogNameName
$CatalogName = $XAmachine | %{ $_.CatalogName }
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
if ($wmi -ne $null) {
	$tests.WMI = "SUCCESS", "Success"
	$LBTime=[Management.ManagementDateTimeConverter]::ToDateTime($wmi.Lastbootuptime)
	[TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)

	if ($uptime.days -gt $maxUpTimeDays) {
		"reboot warning, last reboot: {0:D}" -f $LBTime | LogMe -display -warning
		$tests.Uptime = "WARNING", $uptime.days
	} else {
		$tests.Uptime = "SUCCESS", $uptime.days
	}
} else {
	"WMI connection failed - check WMI for corruption" | LogMe -display -error
	stop-job $job
}
#----
  
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
#"PVS Cache max size: {0:n2}GB" -f($PvsWriteMaxSize / 1GB) | LogMe -display
$tests.WriteCacheType = "NEUTRAL", $CachePVSType
if ($CacheDisk -lt ($PvsWriteMaxSize * 0.5)) {
"WriteCache file size is low" | LogMe
$tests.WriteCacheSize = "SUCCESS", $CacheDiskGB
}
elseif ($CacheDisk -lt ($PvsWriteMaxSize * 0.8)) {
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
  
# Check services
$services = Get-Service -Computer $machineDNS
  
if (($services | ? {$_.Name -eq "Spooler"}).Status -Match "Running") {
"SPOOLER service running..." | LogMe
$tests.Spooler = "SUCCESS","Success"
}
else {
"SPOOLER service stopped" | LogMe -display -error
$tests.Spooler = "ERROR","Error"
}
  
if (($services | ? {$_.Name -eq "cpsvc"}).Status -Match "Running") {
"Citrix Print Manager service running..." | LogMe
$tests.CitrixPrint = "SUCCESS","Success"
}
else {
"Citrix Print Manager service stopped" | LogMe -display -error
$tests.CitrixPrint = "ERROR","Error"
}
  
}
else { $tests.Ping = "Error", $result }
#END of Ping-Section
  
# Column Serverload
$Serverload = $XAmachine | %{ $_.LoadIndex }
"Serverload: $Serverload" | LogMe -display -progress
if ($Serverload -ge $loadIndexError) { $tests.Serverload = "ERROR", $Serverload }
elseif ($Serverload -ge $loadIndexWarning) { $tests.Serverload = "WARNING", $Serverload }
else { $tests.Serverload = "SUCCESS", $Serverload }
  
# Column MaintMode
$MaintMode = $XAmachine | %{ $_.InMaintenanceMode }
"MaintenanceMode: $MaintMode" | LogMe -display -progress
if ($MaintMode) { 
	$objMaintenance = $Maintenance | Where { $_.TargetName.ToUpper() -eq $XAmachine.MachineName.ToUpper() } | Select -First 1
	If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
	"MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
	$tests.MaintMode = "WARNING", $MaintenanceModeOn
	$ErrorVDI = $ErrorVDI + 1
}
else { $tests.MaintMode = "SUCCESS", "OFF" }
  
# Column RegState
$RegState = $XAmachine | %{ $_.RegistrationState }
"State: $RegState" | LogMe -display -progress
  
if ($RegState -ne "Registered") { $tests.RegState = "ERROR", $RegState }
else { $tests.RegState = "SUCCESS", $RegState }

# Column VDAVersion AgentVersion
$VDAVersion = $XAmachine | %{ $_.AgentVersion }
"VDAVersion: $VDAVersion" | LogMe -display -progress
$tests.VDAVersion = "NEUTRAL", $VDAVersion

# Column HostedOn 
$HostedOn = $XAmachine | %{ $_.HostingServerName }
"HostedOn: $HostedOn" | LogMe -display -progress
$tests.HostedOn = "NEUTRAL", $HostedOn

  
# Column ActiveSessions
$ActiveSessions = $XAmachine | %{ $_.SessionCount }
"Active Sessions: $ActiveSessions" | LogMe -display -progress
$tests.ActiveSessions = "NEUTRAL", $ActiveSessions

# Column ConnectedUsers
$ConnectedUsers = $XAmachine | %{ $_.AssociatedUserNames }
"Connected users: $ConnectedUsers" | LogMe -display -progress
$tests.ConnectedUsers = "NEUTRAL", $ConnectedUsers
  
# Column DeliveryGroup
$DeliveryGroup = $XAmachine | %{ $_.DesktopGroupName }
"DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
$tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup


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

        # Check C Disk Usage 
        $HardDisk = CheckHardDiskUsage -hostname $machineDNS -deviceID "C:"
		if ($HardDisk -ne $null) {	
			$XAPercentageDS = $HardDisk.PercentageDS
			$frSpace = $HardDisk.frSpace

			If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests.CFreespace = "SUCCESS", "$frSpace GB" } 
			ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests.CFreespace = "ERROR", "Err" }
			ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests.CFreespace = "ERROR", "$frSpace GB" } 
			ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests.CFreespace = "WARNING", "$frSpace GB" }     
			Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests.CFreespace = "ERROR", "$frSpace GB" }
			
			$XAPercentageDS = 0
			$frSpace = 0
			$HardDisk = $null
		}
		
		$tests.DFreespace = "NEUTRAL", "N/A" 
		if ( $XAServerHaveD -eq "1" ) {
		# Check D Disk Usage 
        $HardDiskd = CheckHardDiskUsage -hostname $machineDNS -deviceID "D:"
		if ($HardDiskd -ne $null) {			
			$XAPercentageDSd = $HardDiskd.PercentageDS
			$frSpaced = $HardDiskd.frSpace

			If ( [int] $XAPercentageDSd -gt 15) { "Disk Free is normal [ $XAPercentageDSd % ]" | LogMe -display; $tests.DFreespace = "SUCCESS", "$frSpaced GB" } 
			ElseIf ([int] $XAPercentageDSd -eq 0) { "Disk Free test failed" | LogMe -error; $tests.DFreespace = "ERROR", "Err" }
			ElseIf ([int] $XAPercentageDSd -lt 5) { "Disk Free is Critical [ $XAPercentageDSd % ]" | LogMe -error; $tests.DFreespace = "ERROR", "$frSpaced GB" } 
			ElseIf ([int] $XAPercentageDSd -lt 15) { "Disk Free is Low [ $XAPercentageDSd % ]" | LogMe -warning; $tests.DFreespace = "WARNING", "$frSpaced GB" }     
			Else { "Disk Free is Critical [ $XAPercentageDSd % ]" | LogMe -error; $tests.DFreespace = "ERROR", "$frSpaced GB" }  
			
			$XAPercentageDSd = 0
			$frSpaced = 0
			$HardDiskd = $null
		}
	}




  
" --- " | LogMe -display -progress
  
# Check to see if the server is in an excluded folder path
if ($ExcludedCatalogs -contains $CatalogName) { "$machineDNS in excluded folder - skipping" | LogMe -display -progress }
else { $allXenAppResults.$machineDNS = $tests }
}
  
}
else { "XenApp Check skipped because ShowXenAppTable = 0 or Farm is < V7.x " | LogMe -display -progress }
  
"####################### Check END ####################################################################################" | LogMe -display -progress

# ======= Write all results to an html file =================================================
# Add Version of XenDesktop to EnvironmentName
$XDmajor, $XDminor = $controllerversion.Split(".")[0..1]
$XDVersion = "$XDmajor.$XDminor"
$EnvironmentName = "$EnvironmentName $XDVersion"
$emailSubject = ("$EnvironmentName Farm Report - " + $ReportDate)

Write-Host ("Saving results to html report: " + $resultsHTM)
writeHtmlHeader "$EnvironmentName Farm Report" $resultsHTM
  
# Write Table with the Controllers
writeTableHeader $resultsHTM $XDControllerFirstheaderName $XDControllerHeaderNames $XDControllerHeaderWidths $XDControllerTableWidth
$ControllerResults | sort-object -property XDControllerFirstheaderName | %{ writeData $ControllerResults $resultsHTM $XDControllerHeaderNames }
writeTableFooter $resultsHTM
  
# Write Table with the Catalogs
writeTableHeader $resultsHTM $CatalogHeaderName $CatalogHeaderNames $CatalogWidths $CatalogTablewidth
$CatalogResults | %{ writeData $CatalogResults $resultsHTM $CatalogHeaderNames}
writeTableFooter $resultsHTM
  
  
# Write Table with the Assignments (Delivery Groups)
writeTableHeader $resultsHTM $AssigmentFirstheaderName $vAssigmentHeaderNames $vAssigmentHeaderWidths $Assigmenttablewidth
$AssigmentsResults | sort-object -property ReplState | %{ writeData $AssigmentsResults $resultsHTM $vAssigmentHeaderNames }
writeTableFooter $resultsHTM

# Write Table with all XenApp Servers
if ($ShowXenAppTable -eq 1 ) {
writeTableHeader $resultsHTM $XenAppFirstheaderName $XenAppHeaderNames $XenAppHeaderWidths $XenApptablewidth
$allXenAppResults | sort-object -property CatalogName | %{ writeData $allXenAppResults $resultsHTM $XenAppHeaderNames }
writeTableFooter $resultsHTM
}
else { "No XenApp output in HTML " | LogMe -display -progress }

# Write Table with all Desktops
if ($ShowDesktopTable -eq 1 ) {
writeTableHeader $resultsHTM $VDIFirstheaderName $VDIHeaderNames $VDIHeaderWidths $VDItablewidth
$allResults | sort-object -property CatalogName | %{ writeData $allResults $resultsHTM $VDIHeaderNames }
writeTableFooter $resultsHTM
}
else { "No XenDesktop output in HTML " | LogMe -display -progress }
  
 
writeHtmlFooter $resultsHTM

#send email
$emailMessage = New-Object System.Net.Mail.MailMessage
$emailMessage.From = $emailFrom
$emailMessage.To.Add( $emailTo )
$emailMessage.Subject = $emailSubject 
$emailMessage.IsBodyHtml = $true
$emailMessage.Body = (gc $resultsHTM) | Out-String
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
#=========== History END ===========================================================================
