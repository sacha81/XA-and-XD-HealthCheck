#==============================================================================================
# Created on: 11.2014 modfied 09.2025 Version: 1.5.8
# Created by: Sacha / sachathomet.ch & Contributers (see changelog at EOF)
# File name: XA-and-XD-HealthCheck.ps1
#
# Description: This script checks a Citrix XenDesktop and/or XenApp 7.x Farm
# It generates a HTML output File which will be sent as Email.
#
# Initial versions tested on XenApp/XenDesktop 7.6 and XenDesktop 5.6 
# Newest version tested on CVAD 2203 and above, including Citrix Cloud.
#
# Prerequisite: Config file, a Delivery Controller for on-prem with according privileges necessary 
# Config file:  In order for the script to work properly, it needs a configuration file.
#               This has the same name as the script, with extension _Parameters.
#               The script name can't contain any another point, even with a version.
#               Example: Script = "XA and XD HealthCheck.ps1", Config = "XA and XD HealthCheck_Parameters.xml"
#
# Call by :     Manual or by Scheduled Task, e.g. once a day
#               !! If you run it as scheduled task you need to add with argument "non interactive"
#               or your user has interactive persmission!
#
# IMPORTANT: For this script to function correctly it requires:
#            If using on-prem Delivery Controllers, it requires TCP Ports 80 and 443 for the PowerShell SDK and XDPing test.
#            If using Citrix Cloud and you have Cloud Connectors, it requires TCP Port 80 for the XDPing test.
#            For all other tests it is important to ensure WinRM is configured and enabled on all the Citrix Infrastructure
#            servers and Session Hosts. This script requires WinRM, and will attempt to fallback to WMI (DCOM) if WinRM is not
#            available.
#            It also tries the UNC path for drive shares, such as C$ and D$, the PVS/MCSIO write-cache drive you specify,
#            and to read the Nvidia log file.
#            Whilst it does a ping test, its success is not as important as the script has been changed from version 1.4.7 to
#            proceed, even if ping times out or fails. This allows for complex environments where Session Hosts may be in
#            different security zones.
#            If you use the Windows host-based firewall, or a 3rd party product, add some rules to allow for this connectivity.
#            For example, the following Inbound Windows host-based firewall rules are requires across all session hosts:
#            - Windows Remote Management (HTTP-In) (Domain/Private)
#            - Windows Management Instrumentation (DCOM-In)
#            - Windows Management Instrumentation (WMI-In)
#            - File and Printer Sharing (SMB-In)
#            As of version 1.4.7 the script now has 3 extra columns in the VDI and XenApp/RDSH tables labeled WinRM, WMI and
#            UNC. These will either be flagged as True or False, which will help you to easily identify connectivity issues
#            experienced by the script.
#
# Example Syntax:
#   Default (assumes the XA-and-XD-HealthCheck_Parameters.xml file is present):
#     powershell -executionpolicy bypass .\XA-and-XD-HealthCheck.ps1
#   Specify a different Parameters file:
#     powershell -executionpolicy bypass .\XA-and-XD-HealthCheck.ps1 -ParamsFile:"IOC-OPS_Params.xml"
#   Process ALL parameters files in the same folder as the script:
#     powershell -executionpolicy bypass .\XA-and-XD-HealthCheck.ps1 -All
#   Process ALL parameters files in the same folder as the script in runspaces, so they run in parallel:
#     powershell -executionpolicy bypass .\XA-and-XD-HealthCheck.ps1 -All -UseRunspace
#
#==============================================================================================

# Don't change below here if you don't know what you are doing ... 
#==============================================================================================
Param (
       [String]$ParamsFile="XA-and-XD-HealthCheck_Parameters.xml",
       [Switch]$All,
       [Switch]$UseRunspace
      )

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
# Get all the loaded SnapIns that we pass into the scriptblock so that it's written to the log file.
# We force the output into an array to ensure we have consistent behaviour.
$SnapIns = @()
$SnapIns += (Get-PSSnapin "Citrix.*" -EA silentlycontinue)

#==============================================================================================

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

#==============================================================================================

# Create the runspace pool.

Function Get-LogicalProcessorCount {
  $NumberOfLogicalProcessors = 0
  $ProcessorCountArray = @()
  Try {
    $VerbosePreference = 'SilentlyContinue'
    $ProcessorCountArray = Get-CimInstance -ClassName win32_processor -ErrorAction Stop | Select-Object DeviceID, SocketDesignation, NumberOfLogicalProcessors
    $VerbosePreference = 'Continue'
    ForEach ($ProcessorCount in $ProcessorCountArray) {
      $NumberOfLogicalProcessors = $NumberOfLogicalProcessors + $ProcessorCount.NumberOfLogicalProcessors
    }
  }
  Catch [System.Exception]{
    #$($Error[0].Exception.Message)
  }
  return $NumberOfLogicalProcessors
}

If ($UseRunspace) {
  $MaxThreads = (Get-LogicalProcessorCount)
  If ($MaxThreads -gt 1) {
    $MaxThreads = $MaxThreads - 1
  } Else {
    $MaxThreads = 4
  }
  write-verbose "$(Get-Date): Runspace pool configuration:" -verbose
  write-verbose "$(Get-Date): - minimum number of opened/concurrent runspaces for the pool: 1" -verbose
  write-verbose "$(Get-Date): - maximum number of opened/concurrent runspaces for the pool: $($MaxThreads)" -verbose
  $RunspacePool = [runspacefactory]::CreateRunspacePool(
      1, # minimum number of opened/concurrent runspaces for the pool
      $MaxThreads # maximum number of opened/concurrent runspaces for the pool
    )
  $RunspacePool.open()
  $runspaces = New-Object System.Collections.ArrayList
}

#==============================================================================================

# Get all XML Files

# Import parameter file
$Global:ParameterFiles = @()
$Global:ParameterFilePath = $ScriptPath
$Global:CountParameterFiles = 0
If ($All -eq $False) {
  $Global:ParameterFile = $ScriptName + "_Parameters.xml"
  If (![string]::IsNullOrEmpty($ParamsFile)) {
    $Global:ParameterFile = $ParamsFile
  }
  $Global:ParameterFiles += $ParameterFile
} Else {
  # Get the XML files
  # Using the Get-ChildItem –Filter parameter provides the fastest outcome, but we then need to pipe the
  # output to Where-Object for further filtering so we don't pick up any .xmlold files, as an example.
  $Global:ParameterFiles += Get-ChildItem -Path "$($ParameterFilePath + '\')" -Filter '*.xml' | Where-Object {$_.Extension -eq ".xml"} | Select-Object -ExpandProperty Name
  $CountParameterFiles = ($ParameterFiles | Measure-Object).Count
  Write-Verbose "$(Get-Date): There are $CountParameterFiles XML files to process" -verbose
}

$currentDir = Split-Path $MyInvocation.MyCommand.Path
$outputpath = Join-Path $currentDir "" #add here a custom output folder if you wont have it on the same directory

ForEach ($ParameterFile in $ParameterFiles) {

  Write-Verbose "$(Get-Date): Processing $($ParameterFilePath + '\' + $ParameterFile)" -verbose

#==============================================================================================

# Scriptblock

$scriptBlockExecute = {
  param(
        [PSObject[]]$paramBundle,
        [string]$ParameterFilePath,
        [string]$ParameterFile,
        [string]$outputpath,
        [PSObject[]]$SnapIns
       )

  If ($null -ne $paramBundle) {
    ForEach ($obj in $paramBundle) {
      $obj.PSObject.Properties | ForEach-Object {
        If ($_.Name -eq "ParameterFilePath") {
          [string]$ParameterFilePath = $_.Value
        }
        If ($_.Name -eq "ParameterFile") {
          [string]$ParameterFile = $_.Value
        }
        If ($_.Name -eq "outputpath") {
          [string]$outputpath = $_.Value
        }
        If ($_.Name -eq "SnapIns") {
          [PSObject[]]$SnapIns = $_.Value
        }
      }
    }
  }

#==============================================================================================

# Import variables
Function New-XMLVariables {
  param(
        [xml]$cfg,
        [switch]$output
       )
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
    If ($CreateVariable) {
      If ($output) {
        write-verbose "Creating variable $($_.Name) with the value $($VarValue)" -verbose
      }
      New-Variable -Name $_.Name -Value $VarValue -Scope $_.Scope -Force
    }
  }
}

[xml]$cfg = Get-Content ($ParameterFilePath + '\' + $ParameterFile) # Read content of XML file

New-XMLVariables -cfg:$cfg

#==============================================================================================

# These variables should remain constant, so they do not need to be in the XML params file.

# The prefix for the Structured Data ID (SD-ID), which is the bit that comes before the @ symbol. The Private Enterprise Number (PEN) comes
# after the @ symbol to make up the Structured Data ID.
$StructuredDataIDPrefix = "citrixhealthcheck"

# The MSGID field provides a unique identifier for the type of message being sent. This field helps in categorizing and filtering log messages.
$SyslogMsgId = "citrix-health-check"

#==============================================================================================

# These variables are derived from either the XML or constant variables.

$citrixCloudRegionBaseUrl = "https://" + $CCRegion
$SyslogMessageStart = "$LogonDurationInSeconds second logon duration exceeded"
$SyslogAppName = ($EnvironmentName -replace ' ','')
$StructuredDataID = $StructuredDataIDPrefix + "@" + $PrivateEnterpriseNumber

#==============================================================================================

$ReportDate = (Get-Date -UFormat "%A, %d. %B %Y %R")

$logfile = Join-Path $outputpath ("CTXXDHealthCheck.log")
If (![string]::IsNullOrEmpty($OutputLog)) {
  $logfile = Join-Path $outputpath ($OutputLog)
}
$resultsHTM = Join-Path $outputpath ("CTXXDHealthCheck.htm") #add $outputdate in filename if you like
If (![string]::IsNullOrEmpty($OutputHTML)) {
  $resultsHTM = Join-Path $outputpath ($OutputHTML) #add $outputdate in filename if you like
}

$resultsSyslog = Join-Path $outputpath ("CTXXDHealthCheckSyslog.log") #add $outputdate in filename if you like
If (![string]::IsNullOrEmpty($OutputSyslog)) {
  $resultsSyslog = Join-Path $outputpath ($OutputSyslog) #add $outputdate in filename if you like
}

Remove-Item $logfile -force -EA SilentlyContinue
Remove-Item $resultsHTM -force -EA SilentlyContinue
Remove-Item $resultsSyslog -force -EA SilentlyContinue

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
  if ($error) { $logEntry = "[ERROR] $logEntry" ; Write-Host "$logEntry" -Foregroundcolor Red	}
  elseif ($warning) { Write-Warning "$logEntry" ; $logEntry = "[WARNING] $logEntry" }
  elseif ($progress) { Write-Host "$logEntry" -Foregroundcolor Green }
  elseif ($display) { Write-Host "$logEntry" }

  #$logEntry = ((Get-Date -uformat "%D %T") + " - " + $logEntry)
  $logEntry | Out-File $logFile -Append
}

#==============================================================================================

Function Test-OpenPort {
<#
  .SYNOPSIS
  Test-OpenPort is an advanced Powershell function. Test-OpenPort acts like a port scanner. 
  .DESCRIPTION
  Uses Test-NetConnection. Define multiple targets and multiple ports. 
  .PARAMETER
  Target
  Define the target by hostname or IP-Address. Separate them by comma. Default: localhost 
  .PARAMETER
  Port
  Mandatory. Define the TCP port. Separate them by comma. 
  .EXAMPLE
  Test-OpenPort -Target sid-500.com,cnn.com,10.0.0.1 -Port 80,443 
  .NOTES
  Author: Patrick Gruenauer
  Web: https://sid-500.com
  Modified by Jeremy Saunders (jeremy@jhouseconsulting.com)
  .LINK
  None. 
  .INPUTS
  None. 
  .OUTPUTS
  None.
#>
[CmdletBinding()]
param (
       [string[]]$Targets='localhost',
       [Parameter(Mandatory=$true, Helpmessage = 'Enter Port Numbers. Separate them by comma.')]
       [string[]]$Ports,
       [switch]$DisableProgressBar,
       [switch]$Output
      )
  $result=@()
  foreach ($t in $Targets) {
    If ($Output) { write-verbose "Target: $t" -verbose }
    foreach ($p in $Ports) {
      If ($Output) { write-verbose "Testing Port: $p" -verbose }
      If ($DisableProgressBar) {
        $OriginalProgressPreference = $Global:ProgressPreference
        $Global:ProgressPreference = 'SilentlyContinue'
      }
      $a = Test-NetConnection -ComputerName $t -Port $p -WarningAction SilentlyContinue
      If ($Output) { write-verbose "Success: $($a.tcpTestSucceeded)" -verbose }
      $result += New-Object -TypeName PSObject -Property ([ordered]@{
        'Target' = $a.ComputerName;
        'RemoteAddress' = $a.RemoteAddress;
        'Port' = $a.RemotePort;
        'Status' = $a.tcpTestSucceeded
        })
      If ($DisableProgressBar) {
        $Global:ProgressPreference = $OriginalProgressPreference
      }
    }
  }
  return $result
}

# ==============================================================================================

Function Get-BearerToken {
  param (
         [Parameter(Mandatory=$true)][string]
         $clientId,
         [Parameter(Mandatory=$true)][string]
         $clientSecret
        )
  [string]$bearerToken = $null
  [bool]$success = $false
  [hashtable]$body = @{
    'grant_type' = 'client_credentials'
    'client_id' = $clientId
    'client_secret' = $clientSecret
  }
  $response = $null
  $tokenUrl  = 'https://api.cloud.com/cctrustoauth2/root/tokens/clients'
  Try {
    $response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -UseBasicParsing
    if( $null -ne $response ) {
      $bearerToken = "CwsAuth Bearer=$($response | Select-Object -expandproperty access_token)"
      $success = $true
    }
  }
  Catch {
    $bearerToken = "$_.Exception.Message"
  }
  return [PSCustomObject]@{
    bearerToken = $bearerToken
    success = $success
  }
}

Function Get-CCSiteId {
  param (
         [Parameter(Mandatory=$true)]
         [string] $bearerToken,
         [Parameter(Mandatory=$true)]
         [string] $customerId
  )
  $requestUri = "https://api.cloud.com/cvad/manage/me"
  $headers = @{
    "Accept" = "application/json";
    "Authorization" = "$bearerToken";
    "Citrix-CustomerId" = $customerid;
  }
  $response = Invoke-RestMethod -Uri $requestUri -Method GET -Headers $headers 
  return $response.Customers.Sites.Id
}

Function Get-CCSiteDetails {
  param (
        [Parameter(Mandatory=$true)]
        [string] $customerid,
        [Parameter(Mandatory=$true)]
        [string] $sitenameorid,
        [Parameter(Mandatory=$true)]
        [string] $bearerToken
  )
  $requestUri = [string]::Format("https://api.cloud.com/cvad/manage/Sites/{0}", $sitenameorid)
  $headers = @{
        "Accept" = "application/json";
        "Authorization" = "$bearerToken";
        "Citrix-CustomerId" = $customerid;
  }
  $response = Invoke-RestMethod -Uri $requestUri -Method GET -Headers $headers 
  return $response
}

Function Get-CCOrchestrationStatus {
  param (
         [Parameter(Mandatory=$true)]
         [string] $bearerToken,
         [Parameter(Mandatory=$true)]
         [string] $customerId
  )
  $requestUri = [string]::Format("https://{0}.xendesktop.net/citrix/orchestration/api/ping/status", $customerid)
  $headers = @{
    "Accept" = "application/json";
    "Authorization" = "$bearerToken";
    "Citrix-CustomerId" = $customerid;
  }
  $response = Invoke-RestMethod -Uri $requestUri -Method GET -Headers $headers 
  return $response
}

# ==============================================================================================

# The Invoke-RestMethod and Invoke-WebRequest cmdlets will use the default system proxy if configured.
# Setting this variable to True will bypass the proxy, which is certainly important for on-prem connectivity to avoid
# getting "The remote server returned an error: (503) Server Unavailable." error.
If ($NoProxy) {
  # Ensure there is no proxy
  [System.Net.HttpWebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($null)
}
If ($SkipCertificateCheck) {
  add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
  public bool CheckValidationResult(
      ServicePoint srvPoint, X509Certificate certificate,
      WebRequest request, int certificateProblem) { return true; }
}
"@
  [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

# if enabled for Citrix Cloud set the credential profile: 
# Help from https://www.citrix.com/blogs/2016/07/01/introducing-remote-powershell-sdk-v2-for-citrix-cloud/ and 
# from https://hallspalmer.wordpress.com/2019/02/19/manage-citrix-cloud-using-powershell/
$ProfileName = $OutputLog.Substring(0, $OutputLog.Length - 4)
if ( $CitrixCloudCheck -eq "1" ) {
  $CloudAuthSuccess = $True
  $ProfileType = "CloudApi"
  "Connecting to Citrix Cloud" | LogMe -display -progress
  "- CustomerId: $CustomerID" | LogMe -display -progress
  "- ProfileName: $ProfileName" | LogMe -display -progress
  "- ProfileType: $ProfileType" | LogMe -display -progress
  If (Test-Path -path "$SecureClientFile") {
    "- Using the SecureClientFile: $SecureClientFile" | LogMe -display -progress
    "- Executing Set-XDCredentials..." | LogMe -display -progress
    Set-XDCredentials -CustomerId "$CustomerID" -SecureClientFile "$SecureClientFile" -ProfileType "$ProfileType" -StoreAs "$ProfileName"
  } Else {
    "- Using the APIKey and SecretKey" | LogMe -display -progress
    "- Executing Set-XDCredentials..." | LogMe -display -progress
    Set-XDCredentials -CustomerId "$CustomerID" -APIKey "$APIKey" -SecretKey "$SecretKey" -ProfileType "$ProfileType" -StoreAs "$ProfileName"
  }
  Try {
    "- Executing Get-XDAuthentication..." | LogMe -display -progress
    # Using Get-XDAuthenticationEx instead of Get-XDAuthentication will set the Global:XDSerializedNonPersistentMetadata variable. This is the only
    # variable it sets. However, for the Get-XDAuthenticationEx cmdlet the documentation states that "This cmdlet is not intended to be used directly"
    # as it's an "Internally executed cmdlet". So using the Get-XDAuthentication cmdlet ensures we keep within the realms of support. Therefore, using
    # the NonPersistentMetadata property is the best solution.
    #Get-XDAuthenticationEx -ProfileName $ProfileName
    Get-XDAuthentication -ProfileName $ProfileName
    "  - Successfully authenticated using the profile named $ProfileName" | LogMe -display -progress
    "- Executing Get-XDCredentials..." | LogMe -display -progress
    $CCCreds = Get-XDCredentials -ProfileName $ProfileName
    # Get-XDCredentials returns an object of Citrix.Sdk.Proxy.Cmdlets.Base.XDCredentials type from the Citrix.SdkProxy.PowerShellSnapIn.dll module.
    If($Null -ne $CCCreds) {
      If ($CCCreds.CustomerId -eq "$CustomerID") {
        "  - Successfully retrieved the CustomerId from the stored credentials" | LogMe -display -progress
      } Else {
        $CloudAuthSuccess = $False
        "  - Failed to retrieved the CustomerId from the stored credentials" | LogMe -display -error
      }
      "- Getting Global variables set that start with XD..." | LogMe -display -progress
      $XDGlobalVariables = Get-Variable -Name "XD*" -Scope Global
      If (($XDGlobalVariables | Measure-Object).Count -gt 0) {
        $XDGlobalVariables | ForEach-Object {
          "  - $($_.Name) = $($_.Value)" | LogMe -display -progress
        }
      } Else {
        "  - There are no Global variables set that start with XD" | LogMe -display -progress
      }
      $BearerToken = ""
      Try {
        # The Bearer Token is both a direct property of the Get-XDCredentials output or within the NonPersistentMetadata property.
        If ($CCCreds.NonPersistentMetadata.BearerToken -ne $null) {
          "- The NonPersistentMetadata property exists with the BearerToken value" | LogMe -display -progress
          $BearerToken = $CCCreds.NonPersistentMetadata.BearerToken
        }
      }
      Catch {
        "- The property 'NonPersistentMetadata' cannot be found on this object. Will try the Global XDAuthToken variable next." | LogMe -display -warning
      }
      $AdminAddress = ""
      Try {
        If ($CCCreds.NonPersistentMetadata.AdminAddress -ne $null) {
          "- The NonPersistentMetadata property exists with the AdminAddress value" | LogMe -display -progress
          $AdminAddress = $CCCreds.NonPersistentMetadata.AdminAddress
        }
      }
      Catch {
        "- The property 'NonPersistentMetadata' cannot be found on this object. Will try the Global XDSDKProxy variable next." | LogMe -display -warning
      }
      If ([string]::IsNullOrEmpty($BearerToken)) {
        Try {
          Get-Variable -Name "XDAuthToken" -Scope Global -ErrorAction Stop | out-null
          "- The Global XDAuthToken variable is set" | LogMe -display -progress
          $BearerToken = $Global:XDAuthToken
        }
        Catch {
         "- The Global XDSDKProxy variable cannot be retrieved" | LogMe -display -warning
        }
      }
      If ([string]::IsNullOrEmpty($AdminAddress)) {
        Try {
          Get-Variable -Name "XDSDKProxy" -Scope Global -ErrorAction Stop | out-null
          "- The Global XDSDKProxy variable is set" | LogMe -display -progress
          $AdminAddress = $Global:XDSDKProxy
        }
        Catch {
          "- The Global XDSDKProxy variable cannot be retrieved" | LogMe -display -warning
        }
      }
      If (![string]::IsNullOrEmpty($BearerToken) -AND ![string]::IsNullOrEmpty($AdminAddress)) {
        "- BearerToken (AKA XDAuthToken) is set under the BearerToken variable: (For security and confidentiality reasons we do not output the BearerToken here)" | LogMe -display -progress
        "- AdminAddress (AKA XDSDKProxy) is set under the AdminAddress variable: $AdminAddress" | LogMe -display -progress
      } Else {
        $CloudAuthSuccess = $False
        If ([string]::IsNullOrEmpty($BearerToken)) {
            "- The BearerToken (AKA XDAuthToken) cannot be determined" | LogMe -display -error
        }
        If ([string]::IsNullOrEmpty($AdminAddress)) {
          "- The AdminAddress (AKA XDSDKProxy) cannot be determined" | LogMe -display -error
        }
      }
    } Else {
      $CloudAuthSuccess = $False
      "- Get-XDCredentials returned no data" | LogMe -display -error
    }
  }
  Catch {
    "- $($_.Exception.Message)" | LogMe -display -error
    $CloudAuthSuccess = $False
  }
  If ($CloudAuthSuccess) {
    "- Successfully authenticated to Citrix Cloud (DaaS)" | LogMe -display -progress
  } Else {
    "- Failed to authenticate using the profile named $ProfileName" | LogMe -display -error
    Exit
  }
} Else {
  $ProfileType = "OnPrem"
  "Using OnPrem" | LogMe -display -progress
  "- ProfileName: $ProfileName" | LogMe -display -progress
  "- ProfileType: $ProfileType" | LogMe -display -progress
  Try {
    Set-XDCredentials -ProfileType "$ProfileType" -StoreAs "$ProfileName"
    "Connecting to $ProfileType" | LogMe -display -progress
  }
  Catch {
    # The Set-XDCredentials cmdlet is not included with the Studio PowerShell cmdlets
    "Using the PowerShell cmdlets from Citrix Studio and not the PowerShell SDK" | LogMe -display -progress
  }
}
" " | LogMe -display -progress

#==============================================================================================

If ($CitrixCloudCheck -ne 1) { 

  "#### Validating the Delivery Controllers before continuing #########################################" | LogMe -display -progress

  " " | LogMe -display -progress

  $FoundHealthyDeliveryController = $False
  ForEach ($DeliveryController in $DeliveryControllers){
    "Validating Delivery Controller: $DeliveryController" | LogMe -display -progress
    If ($DeliveryController -ieq "LocalHost"){
        $DeliveryController = [System.Net.DNS]::GetHostByName('').HostName
    }
    $results = Test-OpenPort -Targets:"$DeliveryController" -Ports:"80","443" -DisableProgressBar
    ForEach ($result in $results) {
      If ($result.Status) {
        $FoundHealthyDeliveryController = $True
      } Else {
        $FoundHealthyDeliveryController = $False
      }
    }
    If ($FoundHealthyDeliveryController) {
        $AdminAddress = $DeliveryController
        $FoundHealthyDeliveryController = $True
        "- Using: $DeliveryController" | LogMe -display -progress
        break
    }
  }

  If ($FoundHealthyDeliveryController -eq $False) {
    "Unable to validate a healthy Delivery Controller" | LogMe -display -progress
    "- Locally installed Studio PowerShell module or the PowerShell SDK requires TCP 80, 443 open to the remote Delivery Controllers at minimum" | LogMe -display -progress
    "- Exiting the script" | LogMe -display -progress
    Exit
  }

  " " | LogMe -display -progress
} Else {
  "Delivery Controllers/Cloud Connectors are not required for PowerShell SDK connectivity to Citrix Cloud" | LogMe -display -progress
  $AdminAddress = ""
}

" " | LogMe -display -progress

#==============================================================================================

If ($CitrixCloudCheck -ne 1) {
  #Header for Table "XD/XA Controllers" Get-BrokerController
  $XDControllerFirstheaderName = "ControllerServer"
  $XDControllerHeaderNames = "IPv4Address", "Ping", "OSCaption", "OSBuild", "Uptime", "XDPing", "State", "DesktopsRegistered", "ActiveSiteServices"
  $XDControllerTableWidth= 1800
  foreach ($disk in $diskLettersControllers)
  {
    $XDControllerHeaderNames += "$($disk)Freespace"
  }
  $XDControllerHeaderNames += "LogicalProcessors", "Sockets", "CoresPerSocket", "AvgCPU", "TotalPhysicalMemoryinGB", "MemUsg"
}
If ($ShowCrowdStrikeTests -eq 1) {
  $XDControllerHeaderNames += "CSEnabled", "CSGroupTags"
}

if ($CitrixCloudCheck -eq 1 -AND $ShowCloudConnectorTable -eq 1 ) {
  #Header for Table "Cloud Connector Servers"
  $CCFirstheaderName = "CloudConnectorServer"
  $CCHeaderNames = "IPv4Address", "Ping", "OSCaption", "OSBuild", "Uptime", "XDPing", "CitrixServices"
  $CCTableWidth= 1800
  foreach ($disk in $diskLettersControllers)
  {
    $CCHeaderNames += "$($disk)Freespace"
  }
  $CCHeaderNames += "LogicalProcessors", "Sockets", "CoresPerSocket", "AvgCPU", "TotalPhysicalMemoryinGB", "MemUsg"
}
If ($ShowCrowdStrikeTests -eq 1) {
  $CCHeaderNames += "CSEnabled", "CSGroupTags"
}

If ($ShowStorefrontTable -eq 1) {
  #Header for Table "Storefront Servers"
  $SFFirstheaderName = "StorefrontServer"
  $SFHeaderNames = "IPv4Address", "Ping", "OSCaption", "OSBuild", "Uptime", "CitrixServices"
  $SFTableWidth= 1800
  foreach ($disk in $diskLettersControllers)
  {
    $SFHeaderNames += "$($disk)Freespace"
  }
  $SFHeaderNames += "LogicalProcessors", "Sockets", "CoresPerSocket", "AvgCPU", "TotalPhysicalMemoryinGB", "MemUsg"
}
If ($ShowCrowdStrikeTests -eq 1) {
  $SFHeaderNames += "CSEnabled", "CSGroupTags"
}

#Header for Table "Fail Rates" FUTURE
#$CTXFailureFirstheaderName = "Checks"
#$CTXFailureHeaderNames = "#", "in Percentage", "CauseServiceInterruption", "CausePartialServiceInterruption"
#$CTXFailureTableWidth= 1200

#Header for Table "CTX Licenses" Get-BrokerController
$CTXLicFirstheaderName = "LicenseName"
$CTXLicHeaderNames = "LicenseServer", "Count", "InUse", "Available"
$CTXLicTableWidth= 1200
  
#Header for Table "MachineCatalogs" Get-BrokerCatalog
$CatalogHeaderName = "CatalogName"
$CatalogHeaderNames = "AssignedToUser", "AssignedToDG", "NotToUserAssigned", "Unassigned", "ProvisioningType", "AllocationType", "MinimumFunctionalLevel", "RecommendedMinimumFunctionalLevel", "UsedMCSSnapshot", "MasterImageVMDate", "UseFullDiskClone", "UseWriteBackCache", "WriteBackCacheMemSize"
$CatalogTablewidth = 1200

#Header for Table "DeliveryGroups" Get-BrokerDesktopGroup
$AssigmentFirstheaderName = "DeliveryGroup"
$vAssigmentHeaderNames = "PublishedName", "DesktopKind", "SessionSupport", "Enabled", "MaintenanceMode", "ShutdownAfterUse", "TotalMachines", "DesktopsAvailable", "MachinesInMaintMode", "PercentageOfMachinesInMaintMode", "DesktopsUnregistered", "DesktopsPowerStateUnknown", "DesktopsNotUsedLast90Days", "DesktopsInUse", "DesktopsFree", "MinimumFunctionalLevel", "RecommendedMinimumFunctionalLevel"
$Assigmenttablewidth = 1200

#Header for Table "ConnectionFailureOnMachine" Get-BrokerConnectionLog
$BrkrConFailureFirstheaderName = "ConnectionFailureOnMachine"
$BrkrConFailureHeaderNames = "BrokeringTime", "ConnectionFailureReason", "BrokeringUserName", "BrokeringUserUPN"
$BrkrConFailureTableWidth= 1200

#Header for Table "HypervisorConnection" Get-BrokerHypervisorConnection
$HypervisorConnectionFirstheaderName = "HypervisorConnection"
$HypervisorConnectionHeaderNames = "State", "IsReady", "MachineCount", "FaultState", "FaultReason", "TimeFaultStateEntered", "FaultStateDuration"
$HypervisorConnectiontablewidth = 1200

#Header for Table "VDI Singlesession Checks" Get-BrokerMachine
$VDIfirstheaderName = "virtualDesktops"
$VDIHeaderNames = "IPv4Address", "CatalogName", "DeliveryGroup", "Ping", "WinRM", "WMI", "UNC", "MaintMode", "Uptime", "RegState", "PowerState", "LastConnectionTime", "VDAVersion", "OSCaption", "OSBuild"
$VDIHeaderNames += "Spooler", "CitrixPrint", "UPMEnabled", "FSLogixEnabled", "WEMEnabled"
If ($ShowCrowdStrikeTests -eq 1) {
  $VDIHeaderNames += "CSEnabled", "CSGroupTags"
}
$VDIHeaderNames += "AssociatedUserNames"
$VDIHeaderNames += "displaymode", "EDT_MTU"
$VDIHeaderNames += "IsPVS", "IsMCS", "DiskMode", "MCSImageOutOfDate", "PVSvDiskName", "WriteCacheType", "vhdxSize_inGB", "WCdrivefreespace"
$VDIHeaderNames += "NvidiaLicense","NvidiaDriverVer"
$VDIHeaderNames += "LogicalProcessors", "Sockets", "CoresPerSocket", "TotalPhysicalMemoryinGB"
$VDIHeaderNames += "Tags", "HostedOn"
$VDItablewidth = 2400

#Header for Table "XenApp/RDS/Multisession Checks" Get-BrokerMachine
$XenAppfirstheaderName = "virtualApp-Servers"
$XenAppHeaderNames = "IPv4Address", "CatalogName", "DeliveryGroup", "Ping", "WinRM", "WMI", "UNC", "Serverload", "MaintMode", "Uptime", "RegState", "PowerState", "LastConnectionTime", "VDAVersion", "OSCaption", "OSBuild"
$XenAppHeaderNames += "Spooler", "CitrixPrint", "UPMEnabled", "FSLogixEnabled", "WEMEnabled"
If ($ShowCrowdStrikeTests -eq 1) {
  $XenAppHeaderNames += "CSEnabled", "CSGroupTags"
}
$XenAppHeaderNames += "RDSGracePeriod", "RDSGracePeriodExpired", "TerminalServerMode", "LicensingName", "LicensingType", "LicenseServerList"
$XenAppHeaderNames += "IsPVS", "IsMCS", "DiskMode", "MCSImageOutOfDate", "PVSvDiskName", "WriteCacheType", "vhdxSize_inGB", "WCdrivefreespace"
foreach ($disk in $diskLettersWorkers)
{
  $XenAppHeaderNames += "$($disk)Freespace"
}
if ($ShowConnectedXenAppUsers -eq "1") { 
  $XenAppHeaderNames += "ActiveSessions", "ConnectedUsers"
}
else {
  $XenAppHeaderNames += "ActiveSessions"
}
$XenAppHeaderNames += "NvidiaLicense","NvidiaDriverVer"
$XenAppHeaderNames += "LogicalProcessors", "Sockets", "CoresPerSocket", "AvgCPU", "TotalPhysicalMemoryinGB", "MemUsg"
$XenAppHeaderNames += "Tags", "HostedOn"
$XenApptablewidth = 2400

#Header for Table "StuckSessions"  Get-BrokerSession
$StuckSessionsfirstheaderName = "Stuck-Session"
$StuckSessionsHeaderNames  = "CatalogName", "DesktopGroupName", "UserName", "SessionState", "AppState", "SessionStateChangeTime", "LogonInProgress", "LogoffInProgress", "ClientAddress", "ConnectionMode", "Protocol"
$StuckSessionstablewidth = 1800

#==============================================================================================

# There are two great references as follows:
# 1) Steve Noel - Fun with Citrix Functional Levels
#    - https://verticalagetechnologies.com/index.php/2022/11/10/fun-with-citrix-functional-levels
# 2) Carl Webster - Citrix XenApp/XenDesktop/Virtual Apps and Desktop Product Numbers and Versions
#    - https://www.carlwebster.com/citrix-xenapp-xendesktop-virtual-apps-and-desktop-product-numbers-and-versions/
# However, they demonstrate the issue that this information cannot be derived by a simple lookup.
# As Steve documents, you can list all the available Functional Levels, but there is no way to map them to a product version.
# The following table provides a mapping between the Marketing Product Version, Internal Product Version, and Functional Levels
# for the Delivery Group and Machine Catalog MinimumFunctionalLevel.
# For now it will need to be continually maintained with new version numbers and functional levels.

$ProductVersionValues = @"
2603,7.47,L7_34
2511,7.46,L7_34
2507,7.45,L7_34
2503,7.44,L7_34
2411,7.43,L7_34
2407,7.42,L7_34
2402,7.41,L7_34
2311,7.40,L7_34
2308,7.39,L7_34
2305,7.38,L7_34
2303,7.37,L7_34
2212,7.36,L7_34
2209,7.35,L7_34
2206,7.34,L7_34
2203,7.33,L7_30
2112,7.32,L7_30
2109,7.31,L7_30
2106,7.30,L7_30
2103,7.29,L7_25
2012,7.28,L7_25
2009,7.27,L7_25
2006,7.26,L7_25
2003,7.25,L7_25
1912,7.24,L7_20
1909,7.23,L7_20
1906,7.22,L7_20
1903,7.21,L7_20
1811,7.20,L7_20
1808,7.19,L7_9
7.18,7.18,L7_9
7.17,7.17,L7_9
7.16,7.16,L7_9
7.15,7.15,L7_9
7.14,7.14,L7_9
7.13,7.13,L7_9
7.12,7.12,L7_9
7.11,7.11,L7_9
7.9,7.9,L7_9
7.8,7.8,L7_8
7.7,7.7,L7_7
7.6,7.6,L7_6
7.5,7.5,L7
7.1,7.1,L7
7.0,7.0,L7
5.6,5.6,L5
"@

Function Convert-HereStringToArray {
  # This function was written based on information from Doctor Scripto from the Scripting Guys and the comments from Nikolay Kozhemyak
  # https://devblogs.microsoft.com/scripting/powertip-converting-a-here-string-to-an-array-in-one-line-with-powershell/
    param (
        [Parameter(Mandatory,
        ParameterSetName = 'Input')]
        [string[]]
        $HereString
    )
    process {
      return ($HereString.Split(@("`r", "`n"), [StringSplitOptions]::RemoveEmptyEntries))
    }
}

$ProductVersionHashTable = @{}
Convert-HereStringToArray -HereString $ProductVersionValues | ForEach-Object {
  $row = $_ -Split(',')
  if (!$ProductVersionHashTable.ContainsKey($row[0])) {
    $ProductVersionHashTable[$row[0]] = $row
  }
}

Function Convert-FunctionalLevelToVersion {
  # This is a helper function to convert a FunctionalLevel System.String type to a System.Version type.
  param(
        [Parameter(Mandatory)]
        [string]$FunctionalLevel
  )
  if ($FunctionalLevel -match '^L(\d+)_(\d+)$') {
    return [version]"$($matches[1]).$($matches[2])"
  }
  return $null
}

Function Convert-VersionToFunctionalLevel {
  # This is a helper function to convert a System.Version type to a System.String type in the format
  # of the MinimumFunctionalLevel.
  param(
        [Parameter(Mandatory)]
        $InputVersion
  )
  # Convert to [version] if needed
  if ($InputVersion -isnot [version]) {
    $InputVersion = [version]$InputVersion
  }
  # If minor version is 0, return L<Major>
  if ($InputVersion.Minor -eq 0) {
    return "L$($InputVersion.Major)"
  }
  # Otherwise return L<Major>_<Minor>
  return "L{0}_{1}" -f $InputVersion.Major, $InputVersion.Minor
}

Function Convert-ToComparableVersion {
  # This function allows us to convert to type system.version so that we can easily compare values
  # between the new and legacy product version numbering.
  param(
        [Parameter(Mandatory)]
        [string]$InputString
  )
  if ([string]::IsNullOrWhiteSpace($InputString)) { return $null }
  $s = $InputString.Trim().TrimEnd('.')

  # Must contain only digits and dots
  if ($s -notmatch '^[0-9\.]+$') { return $null }

  # Force array output (prevents the "2402" -> "24020" concatenation bug)
  $parts = @($s.Split('.') | Where-Object { $_ -ne '' })
  if ($parts.Count -eq 0) { return $null }

  # Build a valid System.Version string (2–4 segments)
  $vParts = $parts

  if ($vParts.Count -eq 1) { $vParts += '0' }       # "7" -> "7.0", "2402" -> "2402.0"
  if ($vParts.Count -gt 4) { $vParts = $vParts[0..3] } # trim extras

  $normalizedVersionString = ($vParts -join '.')

  $ver = $null
  if (-not [System.Version]::TryParse($normalizedVersionString, [ref]$ver)) {
    return $null
  }

  # Build a "marketing-format" key:
  # - 4-digit major => "2507"
  # - 1-digit major (e.g. 7) => "7.18" (major.minor), default minor 0
  $major = $parts[0]

  $marketingKey =
        if ($major -match '^\d{4}$') {
            $major
        }
        elseif ($major -match '^\d$') {
            $minor = if ($parts.Count -ge 2) { $parts[1] } else { '0' }
            "$major.$minor"
        }
        else {
            $null
        }

  [pscustomobject]@{
        Input        = $InputString
        Version      = $ver                       # [System.Version]
        VersionText  = $ver.ToString()            # normalized version string
        MarketingKey = $marketingKey              # "2507" or "7.18"
  }
}

Function Find-CitrixVersion {
  param (
         [ValidateSet("MarketingProductVersion","InternalProductVersion","MinimumFunctionalLevel")]
         [string]$MatchByColumn = "MarketingProductVersion",
         [string]$VersionToFind
        )
  $result = [PSCustomObject]@{
    Found                                      = $False
    VersionToFind                              = $VersionToFind
    MarketingProductVersion                    = "N/A"
    InternalProductVersion                     = "N/A"
    MinimumFunctionalLevel                     = "N/A"
    LowestSupportedVDAVersion                  = "N/A"
    HighestVDAVersionBeforeNextFunctionalLevel = "N/A"
  }
  If ($MatchByColumn -eq "MarketingProductVersion") {
    $VersionToFind = (Convert-ToComparableVersion $VersionToFind).MarketingKey
  }
  If ($MatchByColumn -eq "InternalProductVersion") {
    $VersionToFind = (Convert-ToComparableVersion $VersionToFind).Version
  }
  $ProductVersionHashTable.GetEnumerator() | ForEach-Object {
    If ($MatchByColumn -eq "MarketingProductVersion") {
      If ($_.Key -eq $VersionToFind) {
        $result.Found = $True
        $result.MarketingProductVersion = $_.Key
        $result.InternalProductVersion  = $_.Value[1]
        $result.MinimumFunctionalLevel  = $_.Value[2]
      }
    }
    If ($MatchByColumn -eq "InternalProductVersion") {
      If ($_.Value[1] -eq $VersionToFind) {
        $result.Found = $True
        $result.MarketingProductVersion = $_.Key
        $result.InternalProductVersion  = $_.Value[1]
        $result.MinimumFunctionalLevel  = $_.Value[2]
      }
    }
    If ($MatchByColumn -eq "MinimumFunctionalLevel") {
      If ($_.Value[2] -eq $VersionToFind) {
        $result.Found = $True
        $result.MinimumFunctionalLevel = $_.Value[2]
        $tempVDAVersion = (Convert-ToComparableVersion $_.Key).Version
        If ($result.LowestSupportedVDAVersion -eq "N/A") {
          $result.LowestSupportedVDAVersion = $_.Key
        } Else {
          $TempLowestSupportedVDAVersion = (Convert-ToComparableVersion $result.LowestSupportedVDAVersion).Version
          If ($tempVDAVersion -lt $TempLowestSupportedVDAVersion) {
            $result.LowestSupportedVDAVersion = $_.Key
          }
        }
        If ($result.HighestVDAVersionBeforeNextFunctionalLevel -eq "N/A") {
          $result.HighestVDAVersionBeforeNextFunctionalLevel = $_.Key
        } Else {
          $TempHighestVDAVersionBeforeNextFunctionalLevel = (Convert-ToComparableVersion $result.HighestVDAVersionBeforeNextFunctionalLevel).Version
          If ($tempVDAVersion -gt $TempHighestVDAVersionBeforeNextFunctionalLevel) {
            $result.HighestVDAVersionBeforeNextFunctionalLevel = $_.Key
          }
        }
      }
    }
  }
  Return $result
}

# Syntax examples:
# 1) Get the Internal Product Version and MinimumFunctionalLevel based on the Marketing Product Version
#    Find-CitrixVersion -MatchByColumn:"MarketingProductVersion" -VersionToFind:"2311"
# 2) Get the Marketing Product Version and MinimumFunctionalLevel based on the Internal Product Version
#    Find-CitrixVersion -MatchByColumn:"InternalProductVersion" -VersionToFind:"7.21"
# 3) Get the Lowest Supported VDA Version based on the MinimumFunctionalLevel and the Highest VDA Version Before the Next FunctionalLevel
#    Find-CitrixVersion -MatchByColumn:"MinimumFunctionalLevel" -VersionToFind:"L7_9"

#==============================================================================================

Function ConvertTo-StructuredData {
<#
  This function converts a PSCustomObject or hashtable into the required format for the SYSLOG
  Structured Data section. It handles required RFC 5424 escaping of " and ] inside structured-data values.

  The Id and param names should be 1–32 chars as per the syslog standard (RFC 5424).
  - In RFC 5424 §6.3.2 (SD-ID), the Structured Data ID (SD-ID) must be 1 to 32 visible US-ASCII characters.
  - In RFC 5424 §6.3.3 (PARAM-NAME), each parameter name inside the structured data also has the same restriction: 1–32 characters.
  This keeps syslog messages interoperable between vendors, platforms, and collectors.
  Why the restriction matters
  - Parsing simplicity: syslog parsers don't need to deal with arbitrarily long tokens.
  - Interoperability: messages won’t get truncated or rejected by downstream collectors.
  - Spec compliance: ensures you can claim "RFC 5424-compliant syslog."
  However, by setting the AllowMoreParamChars parameter, you can allow param names to be longer than 32 characters.

  Example syntax:
  Build SD element directly
    ConvertTo-StructuredData -Id 'diskinfo@32473' -Data ([pscustomobject]@{ size='120GB'; used='95GB' })
    -> [diskinfo@32473 size="120GB" used="95GB"]
  Handles arrays
    ConvertTo-StructuredData -Id 'meta@32473' -Data ([pscustomobject]@{ tags=@('blue','green'); note='C:\Path\] "y"' })
    -> [meta@32473 note="C:\\Path\] \"y\"" tags="blue,green"]
  Handles booleans & datetimes
    ConvertTo-StructuredData -Id 'audit@32473' -Data ([pscustomobject]@{ ok=$true; at=[DateTimeOffset]::Now })
    -> [audit@32473 at="2025-09-01T12:30:15.456+08:00" ok="true"]
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Id,   # e.g. "diskinfo@32473"
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$Data,  # PSCustomObject or hashtable
        [switch]$SortProps,  # Sort the properties in alphabetical (ascending) order
        [switch]$AllowMoreParamChars  #Allow a token greater than 32 characters
    )

    function Validate-Token([string]$Token, [string]$What, [switch]$AllowMoreParamChars) {
        if ([string]::IsNullOrWhiteSpace($Token)) { throw "$What cannot be empty." }
        if ($Token.Length -lt 1 -or $Token.Length -gt 32) {
          if ($Token.Length -gt 32 -AND $AllowMoreParamChars) {
            # We are allowing then length of the token to be greater than 32 characters
          } else {
            throw "$What must be 1-32 characters."
          }
        }
        if ($Token -match '[\s=\]"]') { throw "$What cannot contain space, equals sign, right bracket, or double-quote." }
    }

    function Coerce([object]$v) {
        switch ($v) {
            $null { '' }
            { $_ -is [DateTimeOffset] } { $_.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz"); break }
            { $_ -is [DateTime] }       { ([DateTimeOffset]$_).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz"); break }
            { $_ -is [bool] }           { if ($_) { 'true' } else { 'false' }; break }
            default { [string]$_ }
        }
    }

    # Validate SD-ID
    Validate-Token -Token $Id -What 'SD-ID' -AllowMoreParamChars $false

    # Normalize properties
    if ($Data -is [hashtable]) {
        $props = $Data.GetEnumerator() | Sort-Object Name
    } else {
        if ($SortProps) {
          $props = $Data.PSObject.Properties | Sort-Object Name
        } else {
          $props = $Data.PSObject.Properties
        }
    }

    $pairs = foreach ($p in $props) {
        $name = [string]$p.Name
        Validate-Token -Token $name -What 'Param name' -AllowMoreParamChars $AllowMoreParamChars

        $val = $p.Value
        if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
            # Join arrays/lists into one value (params must be unique)
            $val = (@($val | ForEach-Object { Coerce -v $_ }) -join ',')
        } else {
            $val = Coerce $val
        }

        # RFC 5424 escaping: backslash, quote, right-bracket
        $escaped = $val -replace '\\','\\\\' -replace '"','\"' -replace '\]','\]'
        "$name=""$escaped"""
    }

    if (-not $pairs -or $pairs.Count -eq 0) {
        return ("[{0}]" -f $Id)  # valid empty SD element
    }

    return ("[{0} {1}]" -f $Id, ($pairs -join ' '))
}

Function Write-IetfSyslogEntry {
<#
  This function will create an IETF structured Syslog Format (RFC 5424 compliant) log file and/or send to a syslog server.

  This function only supports UDP (classic syslog) and HTTP/HTTPS (API ingest) only. It does not support classic syslog
  using TCP on port 514 and TLS on port 6514. This is because TCP communications being flagged as a reverse shell.

  Note that in Syslog you cannot create a new/custom facility. The facility is a fixed numeric field defined by the standard.
  The valid facility codes are the well-known ones (kernel, user, mail, daemon, …) plus LOCAL0–LOCAL7 (codes 16–23) reserved
  for site-specific use.
  However, what you can do is...
  - Use LOCAL0–LOCAL7 as your "custom" buckets and document what each means in your environment (e.g., LOCAL4 = MyApp,
    LOCAL5 = Payments, etc.).
  - Add rich tags in Structured Data and/or APP-NAME/MSGID to classify events beyond facility (e.g., [app@32473
    svc="billing" tier="prod" region="au-west"]).
  - On the collector, create routes/streams/indices based on facility, app-name, msgid, or your structured data keys.
  So whilst you cannot add new facilities, you can repurpose LOCAL0–LOCAL7 plus rich structured metadata to get virtually
  unlimited categorization.

  Example usage:
  
  1) UDP (classic syslog daemon, port 514 by default):
     Write-IetfSyslogEntry -Message "Job started" `
       -SyslogServer "rsyslog01.acme.local" `
       -CollectorType Syslog

  2) HTTPS ingest API (defaults to 443) with headers & file copy:
     $sd = [pscustomobject]@{
             Id='audit@32473'
             Data=[pscustomobject]@{ actor='svc.backup'; action='start'; jobId='BKP-001'; outcome='success' }
           }
     Write-IetfSyslogEntry -Message "Backup started" `
                           -StructuredDataObject $sd `
                           -SyslogServer "logs.example.com" `
                           -CollectorType HttpApi -Transport HTTPS -HttpPath "/syslog/ingest" `
                           -HttpHeaders @{ "X-Api-Key" = "abc123" } `
                           -LogFilePath "C:\Logs\backup.log" -UseLocalTime

  3) File-only:
     Write-IetfSyslogEntry -Message "Local audit only" `
                           -CollectorType Syslog `
                           -LogFilePath "C:\Logs\audit.log" -FileOnly
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Message,
 
        [string]$SyslogServer,             # host/IP for UDP; host for HTTP/S
 
        # Choose collector kind → sets sane defaults & validates transports
        [ValidateSet("Syslog","HttpApi")]
        [string]$CollectorType = "Syslog",
 
        # Transport choices per collector:
        # - Syslog: UDP only
        # - HttpApi: HTTP or HTTPS
        [string]$Transport,
 
        # Port (defaults based on CollectorType + Transport)
        [int]$Port,
 
        # HTTP options (only for CollectorType=HttpApi)
        [string]$HttpPath = "/",
        [hashtable]$HttpHeaders,
        [string]$HttpContentType = "application/syslog; charset=utf-8",
        [int]$HttpTimeoutSeconds = 15,
        [switch]$SkipHttpCertValidation,   # lab use only
 
        # Syslog metadata
        [ValidateSet("Emergency","Alert","Critical","Error","Warning","Notice","Informational","Debug")]
        [string]$Severity = "Informational",
 
        [ValidateSet("Kernel","User","Mail","Daemon","Auth","Syslog","LPR","News","UUCP","Cron","AuthPriv","FTP","NTP","LogAudit","LogAlert","ClockDaemon","Local0","Local1","Local2","Local3","Local4","Local5","Local6","Local7")]
        [string]$Facility = "User",
 
        [string]$AppName = "PowerShellApp",
        [string]$ProcId = $PID,
        [string]$MsgId = "-",
 
        # Structured Data: pass either a ready string or a wrapper object
        [string]$StructuredData = "-",
        [object]$StructuredDataObject,
 
        # File logging
        [string]$LogFilePath,
        [switch]$FileOnly,
 
        # Time handling
        [switch]$UseLocalTime
    )
 
    if ([string]::IsNullOrWhiteSpace($SyslogServer)) {
      $FileOnly = $True
    }
    # Build StructuredData from wrapper if provided
    if ($StructuredDataObject) {
        $StructuredData = ConvertTo-StructuredData -StructuredDataObject $StructuredDataObject -SortProps -AllowMoreParamChars
    }
    if (-not $StructuredData) { $StructuredData = "-" }
 
    # PRI
    $facilityMap = @{
        "Kernel"=0;"User"=1;"Mail"=2;"Daemon"=3;"Auth"=4;"Syslog"=5;"LPR"=6;"News"=7;"UUCP"=8;"Cron"=9;
        "AuthPriv"=10;"FTP"=11;"NTP"=12;"LogAudit"=13;"LogAlert"=14;"ClockDaemon"=15;
        "Local0"=16;"Local1"=17;"Local2"=18;"Local3"=19;"Local4"=20;"Local5"=21;"Local6"=22;"Local7"=23
    }
    $severityMap = @{
        "Emergency"=0;"Alert"=1;"Critical"=2;"Error"=3;"Warning"=4;"Notice"=5;"Informational"=6;"Debug"=7
    }
    $pri = ($facilityMap[$Facility] * 8) + $severityMap[$Severity]
 
    $version   = 1
    $hostname  = $env:COMPUTERNAME
    $timestamp = if ($UseLocalTime) {
        [DateTimeOffset]::Now.ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
    } else {
        [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    }
 
    $syslogMsg = "<$pri>$version $timestamp $hostname $AppName $ProcId $MsgId $StructuredData $Message"
    $msgBytes  = [System.Text.Encoding]::UTF8.GetBytes($syslogMsg)
 
    # Optional file copy
    if ($LogFilePath) {
        try { Add-Content -Path $LogFilePath -Value $syslogMsg } catch { Write-Warning "File write failed: $_" }
    }
    if ($FileOnly) { return }
 
    # Defaults + validation
    if (-not $Transport) {
        $Transport = if ($CollectorType -eq "Syslog") { "UDP" } else { "HTTPS" }
    }
 
    if ($CollectorType -eq "Syslog") {
        if ($Transport -ne "UDP") {
            throw "CollectorType=Syslog only supports Transport=UDP in this build."
        }
        if (-not $Port) { $Port = 514 }
    } else {
        if ($Transport -notin @("HTTP","HTTPS")) {
            throw "CollectorType=HttpApi requires Transport=HTTP or HTTPS."
        }
        if (-not $Port) { $Port = if ($Transport -eq "HTTPS") { 443 } else { 80 } }
    }
 
    switch ($CollectorType) {
        "Syslog" {
            try {
                # Classic syslog over UDP/514
                $udp = New-Object System.Net.Sockets.UdpClient
                $udp.Connect($SyslogServer, $Port)
                $udp.Send($msgBytes, $msgBytes.Length) | Out-Null
                $udp.Close()
            } catch {
                Write-Warning "UDP send failed: $_"
            }
        }
        "HttpApi" {
            try {
                $handler = New-Object System.Net.Http.HttpClientHandler
                if ($Transport -eq "HTTPS" -and $SkipHttpCertValidation) {
                    $handler.ServerCertificateCustomValidationCallback = { param($m,$c,$ch,$e) $true }
                }
 
                $client  = [System.Net.Http.HttpClient]::new($handler)
                $client.Timeout = [TimeSpan]::FromSeconds($HttpTimeoutSeconds)
 
                $scheme = $Transport.ToLower()  # http | https
                $uri    = "${scheme}://${SyslogServer}`:${Port$HttpPath}"
 
                $content = New-Object System.Net.Http.StringContent($syslogMsg, [System.Text.Encoding]::UTF8, $HttpContentType)
                if ($HttpHeaders) {
                    foreach ($k in $HttpHeaders.Keys) {
                        $content.Headers.Add($k, [string]$HttpHeaders[$k]) 2>$null
                    }
                }
 
                $resp = $client.PostAsync($uri, $content).GetAwaiter().GetResult()
                if (-not $resp.IsSuccessStatusCode) {
                    Write-Warning ("HTTP{0} {1} {2}" -f ($(if($Transport -eq 'HTTPS'){'S'}else{''}), [int]$resp.StatusCode, $resp.ReasonPhrase))
                }
                $client.Dispose()
            } catch {
                Write-Warning "HTTP API send failed: $_"
            }
        }
    }
}

#==============================================================================================

# These functions will complete the basic remote tests to ensure connectivity and firewall ports are not
# blocking WinRM, WMI and UNC connectivity to speed up the processing and reliability of each test against
# each machine it processes. Columns have been added to the output so you can see when it fails and which
# firewall rules may need to be implemented.
# Run this script from a server or Delivery Controller that has all required firewall rules open to all the
# VDAs. The tests will be cascaded so we try WinRM before WMI(DCOM).
# One of the benefits of WinRM is the ability to process tasks in parallel, rather than sequentially, which
# will allow us to run these tests in parallel in a future release of this script.
# But we need to allow for scenarios where WimRM either fails, or firewall rules are blocking access. When
# WimRM isn't available, we try WMI(DCOM). If neither are available, it skips these health checks altogether.
# If WinRM is unhealthy or there are network connectivity issues, you can get the issue where it tells you
# that it's "Attempting to reconnect for up to 4 minutes...". We do not want these types of delays introduced
# into the tests. This issue us specific to PowerShell remoting over WinRM and typically happens when the
# transport connection drops but the WinRM service doesn't immediately close the session. PowerShell then
# enters a reconnect loop, trying to recover the session for up to 4 minutes. This behaviour is hardcoded in
# PowerShell remoting, which uses a 4-minute reconnect timeout (240000 ms) when a session disconnects
# unexpectedly. This reconnect loop is intended to support transient network issues. I have attempted to make
# sure that the IsWinRMAccessible function validates the health of the service and also used the
# OperationTimeoutSec parameter of the Get-CimInstance cmdlet. I have read that there is an undocumented
# registry key that let's you lower the retry timeout on the remote computers, but I have been unable to
# confirm this:
# - Key: HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Client
# - Type: DWORD
# - Value: max_retry_timeout_ms
# - Data: <test various values here>
# You need to restart the WinRM service after setting the key value. If it's MCS or PVS you would
# set this in the image.

Function Test-Ping {
  param (
         [string]$Target = "8.8.8.8",
         [int]$Count = 3,
         [int]$Timeout = 2000
  )
  for ($i = 1; $i -le $Count; $i++) {
    $Result = "FAILED"
    try {
      # The Send method returns instances of the PingReply class directly. A limitation of this
      # method will only return the Success or TimedOut result for everything else. So it's not
      # possible using this method to get other error states such as "Destination Host Unreachable",
      # etc. Even using C# code, there are limitations of what ICMP sockets can detect reliably
      # when count is more than 1. Have also tried P/Invoke using IcmpSendEcho (Windows API) with
      # unique payloads per attempt. This is most likely due to the way further error messages are
      # suppressed as per the RFC. Only ping.exe and Test-NetConnection cmdlet will ultimately
      # return what is needed to be able to report on a reliable ICMP non-success response for
      # more than 1 count due to the way they have been coded to manage this behaviour.
      $pingSender = New-Object System.Net.NetworkInformation.Ping
      #$pingSender = [System.Net.NetworkInformation.Ping]::new()
      $reply = $pingSender.Send($Target, $Timeout)
      if ($reply.Status -eq 'Success') {
        $Result = "SUCCESS"
        break
      } else {
        $Result = "TIMEDOUT"
      }
      $pingSender = $null
    } catch {
      $Result = "FAILED"
    }
    If ($i -lt $Count) {
      Start-Sleep -Milliseconds 1000
    }
  }
  return $Result
}

Function IsWinRMAccessible {
  # The Get-CimInstance and Invoke-Command cmdlets uses WinRM, so we test to see if WinRM is enabled and responding correctly.
  # The Test-WSMan cmdlet submits an identification request that determines whether the WinRM service is running on a local or remote computer.
  # If the tested computer is running the service, the cmdlet displays the WS-Management identity schema, the protocol version, the product
  # vendor, and the product version of the tested service. By default the request is sent to the remote computer anonymously, without using
  # authentication. So we include the Authentication parameter and with the Default value, which will use the authentication method
  # implemented by the WS-Management protocol. When a client computer and a server are both part of the same Active Directory domain (or
  # trusted domains), WinRM defaults to using Kerberos for authentication. This will fail on Kerberos issues, for example. This can happen
  # when there is a missing SPN (Service Principal Name), the computer account has lost its domain trust, etc. Therefore, this test is only
  # successful if the Test-WSMan cmdlet hasn't returned a null (0.0.0) value for the operating system version.
  param ([string]$hostname)
  $success = $False
  try {
    $output = Test-WSMan -Computername $hostname -Authentication Default -ErrorAction Stop
    # Verify that Test-WSMan hasn't returned a null (0.0.0) value for the operating system version.
    If (($output.ProductVersion -split (" "))[1] -ne "0.0.0") {
      $success = $True
    }
  }
  Catch {
    #$_.Exception.Message
  }
  return $success
}

Function IsWMIAccessible {
  # The [WMI] connection to test WMI DCOM connectivity doesn't have a built-in
  # way to support a timeout. It will wait for between 20 to 60 seconds depending
  # on several conditions, such as network connectivity. However, we can work
  # around this limitation by using a background job to control the timeout.
  param (
          [string]$hostname,
          [int]$timeoutSeconds = 20
        )
  $success = $False
  $job = Start-Job -ScriptBlock {
    param($hostname)
    try {
        $null = [WMI]"\\$hostname\root\cimv2"
        return "Success"
    } catch {
        return "Failed: $_"
    }
  } -ArgumentList $hostname

  # Wait for the job with a timeout
  if (Wait-Job -Job $job -Timeout $timeoutSeconds) {
    $result = Receive-Job -Job $job
    #Write-Output "Result from ${hostname}: $result"
    If ($result -eq "Success") {
      $success = $True
    }
  } else {
    #Write-Warning "WMI check timed out after $timeoutSeconds seconds for $hostname"
    Stop-Job $job | Out-Null
  }
  Remove-Job $job | Out-Null
  return $success
}

Function IsUNCPathAccessible {
  # This function uses Threads to do the checks, as we can set a timeout.
  # To make this reliable it needed to be written in C#. There were too
  # many issues using pure PowerShell code.
  param (
         [string]$hostname,
         [string]$share="C$",
         [int]$TimeoutMilliseconds = 1000
        )
  $Path = "\\$hostname\$share"

  # Only define the type once per session
  if (-not ("UncShareChecker" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Threading.Tasks;

public class UncShareCheckResult
{
    public bool Success { get; set; }
    public string Status { get; set; }
    public string Message { get; set; }
}

public static class UncShareChecker
{
    public static UncShareCheckResult CheckShareExists(string path, int timeoutMs)
    {
        var result = new UncShareCheckResult();

        if (!path.EndsWith("\\")) path += "\\";

        Task<UncShareCheckResult> task = Task.Run(() =>
        {
            var innerResult = new UncShareCheckResult();
            try
            {
                if (Directory.Exists(path))
                {
                    innerResult.Success = true;
                    innerResult.Status = "Exists";
                }
                else
                {
                    innerResult.Success = false;
                    innerResult.Status = "NotFound";
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                innerResult.Success = false;
                innerResult.Status = "AccessDenied";
                innerResult.Message = ex.Message;
            }
            catch (Exception ex)
            {
                innerResult.Success = false;
                innerResult.Status = "Error";
                innerResult.Message = ex.Message;
            }
            return innerResult;
        });

        bool completed = false;
        try
        {
            completed = task.Wait(timeoutMs);
        }
        catch (Exception ex)
        {
            return new UncShareCheckResult {
                Success = false,
                Status = "WaitError",
                Message = ex.Message
            };
        }

        if (!completed)
        {
            return new UncShareCheckResult {
                Success = false,
                Status = "Timeout"
            };
        }

        return task.Result;
    }
}
"@ -Language CSharp
  }

  $result = [UncShareChecker]::CheckShareExists($Path, $TimeoutMilliseconds)

  return [PSCustomObject]@{
      Success = $result.Success
      Status  = $result.Status
      Message = $result.Message
    }
}

#==============================================================================================

Function Get-CpuConfigAndUsage {
  # The function checks the processor counter and check for the CPU usage. Takes an average CPU
  # usage for 5 seconds. It check the current CPU usage for 5 secs.
  # It also gets the total logical processor count, as well as number of sockets and cores per
  # socket, which helps identify misconfigurations.
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname
        )
  $ResultProps = @{
    LogicalProcessors = 0
    Sockets = 0
    CoresPerSocket = 0
    CpuUsage = 0
  }
  $LogicalProcessors = 0
  $Sockets = 0
  $CoresPerSocket = 0
  $ProcessorCountArray = @()
  Try {
    If ($UseWinRM) {
      $CpuConfigAndUsage = Get-CimInstance -ClassName win32_processor -ComputerName $hostname -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Select-Object DeviceID, SocketDesignation, NumberOfLogicalProcessors, LoadPercentage
    } Else {
      $CpuConfigAndUsage = Get-WmiObject -computer $hostname -class win32_processor -ErrorAction Stop | Select-Object DeviceID, SocketDesignation, NumberOfLogicalProcessors, LoadPercentage
    }
    $CpuUsage = ($CpuConfigAndUsage | Measure-Object -property LoadPercentage -Average | Select-Object -ExpandProperty Average)
    $ResultProps.CpuUsage = [math]::round($CpuUsage, 1)
    ForEach ($CpuConfig in $CpuConfigAndUsage) {
      $Sockets = $Sockets + 1
      $CoresPerSocket = $CPUConfig.NumberOfLogicalProcessors
      $LogicalProcessors = $LogicalProcessors + $CPUConfig.NumberOfLogicalProcessors
    }
    $ResultProps.LogicalProcessors = $LogicalProcessors
    $ResultProps.Sockets = $Sockets
    $ResultProps.CoresPerSocket = $CoresPerSocket
  }
  Catch [System.Exception]{
    #$($Error[0].Exception.Message)
    "Error returned while checking the CPU config and usage. If you are unable to collect usage information, there may be an issues with the Perfmon Counters." | LogMe -error; return $null
  }
  return $ResultProps
}

#============================================================================================== 

Function CheckMemoryUsage { 
  # The function checks the memory usage and reports the usage value in percentage
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname
        )
  Try {
    If ($UseWinRM) {
      $SystemInfo = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $hostname -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Select-Object TotalVisibleMemorySize, FreePhysicalMemory)
    } Else {
      $SystemInfo = (Get-WmiObject -computername $hostname -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object TotalVisibleMemorySize, FreePhysicalMemory)
    }
    $TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB 
    $FreeRAM = $SystemInfo.FreePhysicalMemory/1MB 
    $UsedRAM = $TotalRAM - $FreeRAM 
    $RAMPercentUsed = ($UsedRAM / $TotalRAM) * 100 
    $RAMPercentUsed = [math]::round($RAMPercentUsed, 2);
    return $RAMPercentUsed
  }
  Catch {
    "Error returned while checking the Memory usage. Perfmon Counters may be fault" | LogMe -error; return $null
  }
}

#==============================================================================================

Function Get-TotalPhysicalMemory {
  # The function gets the total physical memory, which helps identify misconfigurations.
  # Note that the TotalPhysicalMemory property of the Win32_ComputerSystem class states that under some
  # circumstances this property may not return an accurate value for the physical memory. For example, it
  # is not accurate if the BIOS is using some of the physical memory. For an accurate value, use the
  # Capacity property in Win32_PhysicalMemory instead.
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname
        )
  $TotalPhysicalMemoryinGB = 0
  Try {
    If ($UseWinRM) {
      $PhysicalMemory = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $hostname -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Select-Object Capacity
    } Else {
      $PhysicalMemory = Get-WmiObject -computername $hostname -Class Win32_PhysicalMemory -ErrorAction Stop | Select-Object Capacity
    }
    $TotalPhysicalMemory = 0
    ForEach ($Bank in $PhysicalMemory) {
      $TotalPhysicalMemory = $TotalPhysicalMemory + $Bank.Capacity
    }
    $TotalPhysicalMemoryinGB = [math]::round(($TotalPhysicalMemory / 1GB),0)
  }
  Catch [System.Exception]{
    #$($Error[0].Exception.Message)
  }
  return $TotalPhysicalMemoryinGB
}

#==============================================================================================

Function CheckHardDiskUsage {
  # The function checks the HardDrive usage and reports the usage value in percentage and free space.
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname,
         [string]$deviceID
        )
  Try {
    $HardDisk = $null
    If ($UseWinRM) {
      $HardDisk = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $hostname -Filter "DeviceID='$deviceID' and DriveType='3'" -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Select-Object Size,FreeSpace
    } Else {
      $HardDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $hostname -Filter "DeviceID='$deviceID' and DriveType='3'" -ErrorAction Stop | Select-Object Size,FreeSpace
    }
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

Function Get-UpTime {
  # This function will get the uptime of the remote machine, returning both the LastBootUpTime in Date/Time
  # format, and also as a TimeSpan.
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname
        )
  Try {
    If ($UseWinRM) {
      # LastBootUpTime from Get-CimInstance is already been converted to Date/Time
      $LBTime = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $hostname -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop).LastBootUpTime
    } Else {
      $LBTime = [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject -computername $hostname -Class Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime)
    }
    [TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
    $ResultProps = @{ 
      LBTime = $LBTime
      TimeSpan = $uptime
    }
    return $ResultProps
  }
  Catch {
    #"Error returned while checking the uptime via the LastBootUpTime property" | LogMe -error; return 101
  }
  return $null
}

#==============================================================================================

Function Get-OSVersion { 
  # This function will get the OS version of the remote machine
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$hostname
        )
  $ResultProps = @{ 
    Caption = "Notfound"
    Version = "Notfound"
    BuildNumber = "Notfound"
    Error = "Unknown"
  } 
  Try {
    If ($UseWinRM) {
      $OSInfo = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $hostname -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Select-Object Caption, Version, BuildNumber)
    } Else {
      $OSInfo = (Get-WmiObject -computername $hostname -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object Caption, Version, BuildNumber)
    }
    $ResultProps.Caption = $OSInfo.Caption
    $ResultProps.Version = $OSInfo.Version
    $ResultProps.BuildNumber = $OSInfo.BuildNumber
    $ResultProps.Error = "Success"
  }
  Catch {
    $ResultProps.Error = "Error returned while getting the OS version"
  }
  return $ResultProps
}

#==============================================================================================

<#
  There are 3 ways we can check the Nvidia license status on the session hosts.
  1) Using the Nvidia WMI Namespace and Classes as per the Get-NvidiaDetails function below
  2) Check the last update in the "C:\Users\Public\Documents\NvidiaLogging\Log.NVDisplay.Container.exe.log" log file as per the Check-NvidiaLicenseStatus function below
  3) Use nvidia-smi.exe
     nvidia-smi.exe -q > %Computername%.nvidia-smi.txt
     OR
     nvidia-smi.exe -q -f %Computername%.nvidia-smi.txt
#>

Function Get-NvidiaDetails {
  # This function uses the Nvidia WMI Namespace and Classes to get the remote information about the Nvidia profile type and licensing.
  # NVWMI provider Version 2.25 implements support of licensable feature management in System and Gpu classes.
  # If the licensingServer property returns "N/A", we assume that it's using a DLS (Delegated License Service) Client Configuration
  # Token introduced from vGPU release 13.0 or later, which is Windows driver version 471.68 or later, released August 2021.
  # Written by Jeremy Saunders
  param (
         [int]$WinRMTimeoutSec=30,
         [string]$computername = "$env:computername"
        )
  $results = @()
  $ResultProps = @{ 
    GPU_Product_Name = $null
    GPU_Product_Type = $null
    Licensable_Feature_Type = $null
    Licensable_Product = $null
    License_Status = $null
    License_Server = $null
    Display_Driver_Ver = $null
  }
  $namespace = "root\CIMV2\NV"
  # Use the CIM to take advantage of the session option to make multiple queries in a single session.
  # However, you must use the DCOM protocol instead of the default WSMan protocol to support the System class or you will get the following error:
  # "The WS-Management service cannot process the request. The DMTF class in the repository uses a different major version number from the requested class.
  # This class can be accessed using a non-DMTF resource URI."
  Try {
    $DcomSessionOption = New-CimSessionOption -Protocol Dcom -ErrorAction Stop
    $CIMSession = New-CimSession -ComputerName $computername -SessionOption $DcomSessionOption -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop
  }
  Catch [system.exception] {
    #$_.Exception.Message
    $CIMSession = $null
  }
  If ($CIMSession -ne $null) {
    $classname = "Gpu"
    Try {
      $gpus = Get-CimInstance -class $classname -namespace $namespace -CimSession $CIMSession -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop
    }
    Catch [system.exception] {
      #$_.Exception.Message
      $gpus = $null
    }
    $classname = "System"
    Try {
      $system = Get-CimInstance -class $classname -namespace $namespace -CimSession $CIMSession -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop
    }
    Catch [system.exception] {
      #$_.Exception.Message
      $system = $null
    }
    If ($gpus -ne $null) {
      ForEach ( $gpu in $gpus ) {
        $ResultProps.GPU_Product_Name = $gpu.productName
        $productType = switch ( $gpu.productType )
        {
          0 { "unknown" }
          1 { "GeForce" }
          2 { "Quadro" }
          3 { "NVS" }
          4 { "Tesla" }
          default { 'Unknown' }
        }
        $ResultProps.GPU_Product_Type = $productType
        $licensableFeatures = $gpu.licensableFeatures
        ForEach ($licensableFeature in $licensableFeatures) {
          $ResultProps.Licensable_Feature_Type = $licensableFeature.Split('_')[1]
        }
        $licensedProductName = $gpu.licensedProductName
        ForEach ($ProductName in $licensedProductName) {
          $ResultProps.Licensable_Product = $ProductName
        }
        $licensableStatuses = $gpu.licensableStatus
        ForEach ($Status in $licensableStatuses) {
          If ($Status -eq "1") {
            $ResultProps.License_Status = "Enabled"
          } Else {
            $ResultProps.License_Status = "Disabled"
          }
        }
      }
      If ($system -ne $null) {
        $DisplayDriverVer = ($system.verDisplayDriver -split '=')[1]
        $DisplayDriverVer = $DisplayDriverVer.Substring(0,($DisplayDriverVer.Length-1)).Trim()
        If ($DisplayDriverVer.Length -eq 7) {
          $DisplayDriverVer = $DisplayDriverVer.Substring(0,3) + "." + $DisplayDriverVer.Substring(3,2)
        } ElseIf  ($DisplayDriverVer.Length -eq 8) {
          $DisplayDriverVer = $DisplayDriverVer.Substring(0,4) + "." + $DisplayDriverVer.Substring(4,2)
        }
        $ResultProps.Display_Driver_Ver = $DisplayDriverVer
        If ($system.licensingServer -eq "N/A") {
          $ResultProps.License_Server = "Using DLS Client Configuration Token"
        } Else {
          $ResultProps.License_Server = $system.licensingServer + ":" + $system.licensingPort
        }
      } Else {
        # Failed to get GPU System info
        $ResultProps.License_Server = "Unable to obtain"
      }
    } Else {
      # Failed to get GPUs
      $ResultProps.GPU_Product_Name = "N/A"
      $ResultProps.GPU_Product_Type = "N/A"
      $ResultProps.Licensable_Feature_Type = "N/A"
      $ResultProps.Licensable_Product = "N/A"
      $ResultProps.License_Status = "N/A"
      $ResultProps.License_Server = "N/A"
      $ResultProps.Display_Driver_Ver = "N/A"
    }
  } Else {
    $ResultProps.GPU_Product_Name = "N/A"
    $ResultProps.GPU_Product_Type = "N/A"
    $ResultProps.Licensable_Feature_Type = "N/A"
    $ResultProps.Licensable_Product = "N/A"
    $ResultProps.License_Status = "N/A"
    $ResultProps.License_Server = "N/A"
    $ResultProps.Display_Driver_Ver = "N/A"
  }
  $results += New-Object PsObject -Property $ResultProps
  return $results
}

Function Check-NvidiaLicenseStatus {
  # This function gets the last update from the "C:\Users\Public\Documents\NvidiaLogging\Log.NVDisplay.Container.exe.log" log file to verify the current license status.
  # Examples of significant licensing events that are logged are as follows:
  # - Acquisition of a license
  # - Platform detection successful (Licensed in Azure)
  # - Return of a license
  # - Expiration of a license (have only seen this happen after 7 failed attempts to renew)
  # - Failure to acquire a license
  # - Failure to renew a license
  # - License state changes between the unlicensed restricted state (20 mins), unlicensed state (24 hours), and licensed state
  # Reference: https://docs.nvidia.com/vgpu/latest/grid-licensing-user-guide/index.html
  #
  # When the log shows that "License has expired", you either need to restart the machine or restart the "NVIDIA Display Container LS" service.
  #
  # When using specific Azure N-series VMs you do not need a separate NVIDIA vGPU license server. The necessary licensing for the NVIDIA GRID software is included with
  # the Azure service itself. Microsoft redistributes the Azure-optimized NVIDIA GRID drivers which are pre-licensed for the GRID Virtual GPU Software in the Azure
  # environment. The "Platform detection successful" message indicates that the NVIDIA driver has correctly recognized it is running on a the supported Microsoft Azure
  # virtual machine instance where a license is automatically provide through the platform. 
  #
  # Written by Jeremy Saunders
  param (
         [string]$computername = "$env:computername"
        )
  $results = @()
  $ResultProps = @{
    Licensed = "N/A"
    Output_For_HTML = "NEUTRAL"
    Output_To_Log = "nvidiaLicense: N/A"
  }
  $ErrorActionPreference = "stop"
  Try {
    If (Test-Path "filesystem::\\$computername\c$\Users\Public\Documents\NvidiaLogging") {
      If (Test-Path "filesystem::\\$computername\c$\Users\Public\Documents\NvidiaLogging\Log.NVDisplay.Container.exe.log") {
        $Array = Get-Content -Path "\\$computername\c$\Users\Public\Documents\NvidiaLogging\Log.NVDisplay.Container.exe.log"
        $Length = $Array.count
        If ($Length -gt 0) {
          switch ($Array[$Length -1])
          {
            {$Array[$Length -1] -Like "*License renewed successfully*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: License renewed successfully"
                    $ResultProps.Licensed = "Renewed"
                    $ResultProps.Output_For_HTML = "SUCCESS"
                    break
                   }
            {$Array[$Length -1] -Like "*License acquired successfully*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: License acquired successfully"
                    $ResultProps.Licensed = "Acquired"
                    $ResultProps.Output_For_HTML = "SUCCESS"
                    break
                   }
            {$Array[$Length -1] -Like "*Platform detection successful*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: Platform detection successful"
                    $ResultProps.Licensed = "Licensed"
                    $ResultProps.Output_For_HTML = "SUCCESS"
                    break
                   }
            {$Array[$Length -1] -Like "*Failed to acquire license*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: Failed to acquire license"
                    $ResultProps.Licensed = "Failed"
                    $ResultProps.Output_For_HTML = "ERROR"
                    break
                   }
            {$Array[$Length -1] -Like "*Failed to renew license*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: Failed to renew license"
                    $ResultProps.Licensed = "Failed"
                    $ResultProps.Output_For_HTML = "ERROR"
                    break
                   }
            {$Array[$Length -1] -Like "*Failed server communication*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: Failed server communication"
                    $ResultProps.Licensed = "Failed"
                    $ResultProps.Output_For_HTML = "ERROR"
                    break
                   }
            {$Array[$Length -1] -Like "*License has expired*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: License has expired"
                    $ResultProps.Licensed = "Expired"
                    $ResultProps.Output_For_HTML = "ERROR"
                    break
                   }
            {$Array[$Length -1] -Like "*No active license found for the client on license server*"} {
                    $ResultProps.Output_To_Log = "nvidiaLicense: No active license found for the client on license server"
                    $ResultProps.Licensed = "Failed"
                    $ResultProps.Output_For_HTML = "ERROR"
                    break
                   }
           Default {
                    $ResultProps.Output_To_Log = "nvidiaLicense: License not found"
                    $ResultProps.Licensed = "Not Found"
                    $ResultProps.Output_For_HTML = "ERROR"
                   }
          }
        } Else {
          $ResultProps.Output_To_Log = "nvidiaLicense: The Log.NVDisplay.Container.exe.log is invalid"
          $ResultProps.Licensed = "Not Found"
          $ResultProps.Output_For_HTML = "ERROR"
        }
      } Else {
        $ResultProps.Output_To_Log = "Nvidia: The Log.NVDisplay.Container.exe.log does not exist"
        $ResultProps.Licensed = "Not Found"
        $ResultProps.Output_For_HTML = "ERROR"
      }
    } Else {
      If (Test-Path "filesystem::\\$computername\c$\Users\Public\Documents") {
        $ResultProps.Output_To_Log = "nvidiaLicense: NvidiaLogging is not present"
      }
    }
  }
  Catch [system.exception] {
    #$_.Exception.Message
    $ResultProps.Output_To_Log = "nvidiaLicense: Unable to connect via UNC path"
  }
  $ErrorActionPreference = "Continue"
  $results += New-Object PsObject -Property $ResultProps
  return $results
}

#==============================================================================================

Function Get-RDLicenseGracePeriodEventErrorsSinceBoot {
  # This function will look for any Event ID 1069 errors found in the Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin Event Log.
  # This error relates to "The RD Licensing grace period has expired and Licensing mode for the Remote Desktop Session Host server has not
  # been configured. Licensing mode must be configured for continuous operation."
  # When connecting via Citrix users will receive the error "The remote session was disconnected because there are no Terminal Server License
  # Servers available to provide a license".
  [CmdletBinding()]
  param(
         [switch]$UseWinRM,
         [string]$computername = "$env:computername",
         [datetime]$LastBootTime,
         [int]$timeoutSeconds = 30
  )

  $LogName = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin'
  $EventId = 1069

  $result = [PSCustomObject]@{
    ComputerName  = $computername
    LastBootTime  = $LastBootTime
    Found         = $false
    Count         = 0
    Latest        = "N/A"
    Sample        = "N/A"
  }
  $paramBundle = [PSCustomObject]@{
    result       = $result
    LogName      = $LogName
    EventId      = $EventId
    LastBootTime = $LastBootTime
  }

  if ($UseWinRM) {
    try {
      $job = Start-Job -ScriptBlock {
        param (
               [string]$computername,
               $result,
               $paramBundle
              )
        Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
          param (
                 $paramBundle
                )
          If ($null -ne $paramBundle) {
            $result = $paramBundle.result
            $LogName = $paramBundle.LogName
            $EventId = $paramBundle.EventId
            $LastBootTime = $paramBundle.LastBootTime
          }

          $events = @()
          $evErr = $null
          $null = Get-WinEvent -FilterHashtable @{
                    LogName   = $LogName
                    Id        = $EventId
                    StartTime = $LastBootTime
                  } -ErrorAction SilentlyContinue -ErrorVariable evErr |
                     Sort-Object TimeCreated -Descending |
                     Tee-Object -Variable events > $null

          # Handle "no events found" gracefully
          if ($evErr) {
            $onlyNoEvents = $true
            foreach ($e in $evErr) {
              if ($e.Exception.Message -notlike 'No events were found*') {
                $onlyNoEvents = $false
              }
            }
            if (-not $onlyNoEvents) { throw "Get-WinEvent failed: $($evErr[0].Exception.Message)" }
          }
          if ($events.Count -gt 0) {
            $result.Found = $true
            $result.Count = $events.Count
            $result.Latest = ($events | Select-Object -First 1 -ExpandProperty TimeCreated)
            $result.Sample = $events | Select-Object -First 5 TimeCreated, Id, LevelDisplayName, Message
          } Else {
            $result.Latest = "no events found"
            $result.Sample = "no events found"
          }
          return $result
        } -ArgumentList $paramBundle
      } -ArgumentList (,$computername,$result,$paramBundle)

      # Wait for the job with a timeout
      if (Wait-Job -Job $job -Timeout $timeoutSeconds) {
        $result = Receive-Job -Job $job
      } else {
        Stop-Job $job | Out-Null
        $result.Latest = "timed out checking for the events"
        $result.Sample = "timed out checking for the events"
      }
      Remove-Job $job | Out-Null
      return $result
    }
    catch {
      #$_.Exception.Message
      return $result
    }
  } else {
    try {
      $events = @()
      $evErr = $null

      $null = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                LogName   = $LogName
                Id        = $EventId
                StartTime = $LastBootTime
              } -ErrorAction SilentlyContinue -ErrorVariable evErr |
                Sort-Object TimeCreated -Descending |
                Tee-Object -Variable events > $null

      if ($evErr) {
        $onlyNoEvents = $true
        foreach ($e in $evErr) {
          if ($e.Exception.Message -notlike 'No events were found*') {
            $onlyNoEvents = $false
          }
        }
        if (-not $onlyNoEvents) { throw "Get-WinEvent failed: $($evErr[0].Exception.Message)" }
      }

      if ($events.Count -gt 0) {
        $result.Found = $true
      }
      $result.Count = $events.Count
      $result.Latest = ($events | Select-Object -First 1 -ExpandProperty TimeCreated)
      $result.Sample = $events | Select-Object -First 5 TimeCreated, Id, LevelDisplayName, Message
      return $result
    }
    catch {
      #$_.Exception.Message
      return $result
    }
  }
}

Function Get-RDSLicensingDetails {
  # This functions checks the RDS Licensing Details, including grace period.
  # IMPORTANT to understand the following:
  # - If the "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\RCM\GracePeriod" key is missing (typically because
  #   it has been deleted), the GetGracePeriodDays method cannot be found on the Win32_TerminalServiceSetting class. The check will
  #   be  marked as error.
  # - If TerminalServerMode is set to 1 (AppServer) but licensing has never been configured/applied, the RDS licensing grace period
  #   will be set to 0 (good), as it has never had a reference to start it. This can happen with new builds and the output of this
  #   function can be confusing if this is not understood.
  # - That function will still run against Windows 10/11 multi-session hosts, but the output is not relevant.
  # - If TerminalServerMode is set to 0, it's in RemoteAdmin mode and does not have a GetGracePeriodDays method or path to query.
  #   So we first test for TerminalServerMode before continuing.
  # - It gets the days remaining for the RDS licensing grace period so you can address any concerns well within time.
  # - A grace period of 0 days is good. We report on anything less than 10 as a warning and 5 as an error, as it may take time to
  #   drain sessions for a reboot.
  # - We get the LicensingName, LicensingType and LicenseServerList values, which allows us to check for consistency across all the
  #   machines that are part of the overall health check.
  # - LicenseName Meaning:
  #   - Not Configured = No licensing mode is set (common on fresh systems or when licensing is unconfigured)
  #   - Per Device = Traditional RDS mode - one RDS CAL per device
  #   - Per User = Traditional RDS mode - one RDS CAL per user
  #   - Remote Desktop for Admin = Admin mode (no RDS CALs required, limited to 2 concurrent sessions)
  #   - AAD Per User = Azure Virtual Desktop - licensed via Azure AD and Microsoft 365 per-user license
  #   - AVD License = Alternative string to AAD Per User sometimes seen on AVD hosts. The functionally is equivalent.
  #   - Unknown or blank = Licensing misconfigured, corrupted, or not applicable
  # - LicenseType Meaning:
  #   - 1 = Not configured
  #   - 2 = Per Device
  #   - 4 = Per User
  #   - 5 = Remote Desktop for Administration (Admin mode)
  #   - 6 = Invalid or unknown licensing mode. You may also see this on Windows 10/11 Enterprise multi-session hosts on
  #         Azure, which do not use traditional RDS CALs. This is expected and harmless.
  # Although this function does so much more, the original idea for the grace period was taken from Manuel Winkel's (AKA Deyda) Citrix
  # Morning Report script: https://github.com/Deyda/Citrix/tree/master/Morning%20Report
  # Written by Jeremy Saunders.
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$computername = "$env:computername",
         [string]$errordaystocheck = "5",
         [string]$warningdaystocheck = "10"
        )
  $results = @()
  $ResultProps = @{
    TerminalServerMode = "N/A"
    LicensingName = "N/A"
    LicensingType = "N/A"
    LicenseServerList = "N/A"
    GracePeriod = "N/A"
    Status = "N/A"
    Output_For_HTML = "NEUTRAL"
    Output_To_Log = $null
  }
  Try {
    $GracePeriod = $null
    $LicenseServerList = $null
    If ($UseWinRM) {
      $tsSettings = Get-CimInstance -Namespace "root\cimv2\terminalservices" -ClassName "Win32_TerminalServiceSetting" -ComputerName $computername -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop
      If ($tsSettings.TerminalServerMode -eq 1) {
        $GracePeriod = (Invoke-CimMethod -InputObject $tsSettings -MethodName "GetGracePeriodDays" -ErrorAction Stop).DaysLeft
        $LicenseServerList = (Invoke-CimMethod -InputObject $tsSettings -MethodName "GetSpecifiedLicenseServerList" -ErrorAction Stop).SpecifiedLSList
      } Else {
        $ResultProps.TerminalServerMode = "RemoteAdmin"
      }
    } Else {
      $tsSettings = Get-WmiObject -Namespace "root\cimv2\terminalservices" -Class "Win32_TerminalServiceSetting" -ComputerName $computername -ErrorAction Stop
      If ($tsSettings.TerminalServerMode -eq 1) {
        $GracePeriod = (Invoke-WmiMethod -Path $tsSettings.__PATH -Name GetGracePeriodDays -ErrorAction Stop).DaysLeft
        $LicenseServerList = (Invoke-WmiMethod -Path $tsSettings.__PATH -Name GetSpecifiedLicenseServerList -ErrorAction Stop).SpecifiedLSList
      } Else {
        $ResultProps.TerminalServerMode = "RemoteAdmin"
      }
    }
    $ResultProps.LicensingName = $tsSettings.LicensingName
    $ResultProps.LicensingType = $tsSettings.LicensingType
    If ($GracePeriod -ne $null) {
      $ResultProps.TerminalServerMode = "AppServer"
      $ResultProps.GracePeriod = "$GracePeriod"
      If ($GracePeriod -ige "$warningdaystocheck" -OR $GracePeriod -ieq "0") {
        $ResultProps.Status = "Good"
        $ResultProps.Output_For_HTML = "SUCCESS"
        $ResultProps.Output_To_Log = "RDSGracePeriod: Good [ $GracePeriod days ]"
      } ElseIf ($GracePeriod -ilt "$warningdaystocheck" -AND $GracePeriod -ige "$errordaystocheck") {
        $ResultProps.Status = "Warning"
        $ResultProps.Output_For_HTML = "WARNING"
        $ResultProps.Output_To_Log = "RDSGracePeriod: Warning [ $GracePeriod days ]"
      } Else{
        $ResultProps.Status = "Bad"
        $ResultProps.Output_For_HTML = "ERROR"
        $ResultProps.Output_To_Log = "RDSGracePeriod: Critical [ $GracePeriod days ]"
      }
    } Else {
      $ResultProps.GracePeriod = "Unknown"
      $ResultProps.Status = "Unknown"
      $ResultProps.Output_For_HTML = "NEUTRAL"
      $ResultProps.Output_To_Log = "RDSGracePeriod: Unknown"
    }
    If ($LicenseServerList -ne $null) {
      $LicenseServerList = $LicenseServerList -join ", "
      $ResultProps.LicenseServerList = "$LicenseServerList"
    } Else {
      $ResultProps.LicenseServerList = "Unknown"
    }
  }
  Catch [system.exception] {
    #"$($_.Exception.Message)"
    $ResultProps.TerminalServerMode = "Unknown"
    $ResultProps.LicensingName = "Unknown"
    $ResultProps.LicensingType = "Unknown"
    $ResultProps.LicenseServerList = "Unknown"
    $ResultProps.GracePeriod = "Unknown"
    $ResultProps.Status = "Unknown"
    $ResultProps.Output_For_HTML = "NEUTRAL"
    $ResultProps.Output_To_Log = "RDSGracePeriod: Unknown."
  }
  $results += New-Object PsObject -Property $ResultProps
  return $results
}

#==============================================================================================

Function Get-PersonalityInfo {
  # This function will test the Personality.ini or MCSPersonality.ini to determine what sort of
  # session host it is based on the following logic.
  # - If the Personality.ini contains the WriteCacheType entry, it will return the DiskName, DiskMode
  #   and WriteCacheType, and flag it as a PVS image.
  # - If the Personality.ini contains the ListOfDDCs entry, it will return the DiskMode and flag it
  #   as an MCS image.
  # - If the Personality.ini does not contain a WriteCacheType or ListOfDDCs entry, it will return
  #   the DiskMode and flag it as a Standalone image.
  # - If the MCSPersonality.ini contains the ListOfDDCs entry, it will return the DiskMode and flag
  #   it as an MCS image.
  # - If the MCSPersonality.ini does not contains the ListOfDDCs entry, it will return the DiskMode
  #   and flag it as a Standalone image.
  # - It is possible that from VDA 2311 the Personality.ini may be missing altogether for standalone
  #   VDA deployments. This function allows for that.
  #
  # DiskMode for PVS, MCS and Standalone can be as follows:
  # - S (Standard)
  # - P (Private)
  # - ReadWrite
  # - ReadOnly
  # - Unmanaged
  # - Standard
  # - Private
  # - Shared
  # - ReadWriteAppLayering
  # - ReadOnlyAppLayering
  # - PrivateAppLayering
  # - SharedAppLayering
  # - UNC-Path
  #
  # Written by Jeremy Saunders
  param (
         [string]$computername = "$env:computername"
        )
  $results = @()
  $ResultProps = @{
    DiskMode = "N/A"
    IsPVS = $False
    PVSCacheOnDeviceHardDisk = $False
    PVSCacheType = "N/A"
    Output_For_HTML1 = "NEUTRAL"
    Output_To_Log1 = "PVSCacheType: N/A"
    PVSDiskName = ""
    IsMCS = $False
    IsStandAlone = $True
  }
  # Due to VDA upgrades and the change from Personality.ini to MCSPersonality.ini from VDA 2303, both ini files can potentially exist in the root of the C drive.
  # The assumption is that the one that was last written to is the live ini file.
  $Path = Get-ChildItem -Path "filesystem::\\$computername\c$\" -Filter '*Personality.ini' | Where-Object {$_.Extension -eq ".ini"} | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
  If ($null -ne $path) {
    $Personalityini = Get-Content $Path.FullName
    $DiskMode = ($Personalityini | Select-String "DiskMode" | ForEach-Object {$_.Line}).split('=')[1]
    $ResultProps.DiskMode = $DiskMode
    If ($Path.Name -eq "MCSPersonality.ini") {
      $ListOfDDCs = $Personalityini -match 'ListOfDDCs='
      If ($ListOfDDCs -ne $null) {
        $ResultProps.IsMCS = $True
        $ResultProps.IsStandAlone = $False
      }
    }
    If ($Path.Name -eq "Personality.ini") {
      If ($Personalityini | Where-Object { $_.Contains('WriteCacheType=')}) {
        $ResultProps.IsPVS = $True
        $ResultProps.IsStandAlone = $False
        $WriteCacheType = ($Personalityini | Select-String "WriteCacheType" | ForEach-Object {$_.Line}).split('=')[1]
        $ResultProps.PVSCacheType = $WriteCacheType
        switch ($WriteCacheType)
        {
          "9" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $True
               $ResultProps.Output_To_Log1 = "PVSCacheType: 9 - WC is set to Cache to Device Ram with overflow to HD"
               $ResultProps.PVSCacheType = "9 - WC to Ram with overflow to HD"
               $ResultProps.Output_For_HTML1 = "SUCCESS"
               break
              }
          "0" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $False
               $ResultProps.Output_To_Log1 = "PVSCacheType: 0 - WC is not set because vDisk is in PrivateMode (R/W)"
               $ResultProps.PVSCacheType = "0 - vDisk is in PrivateMode (R/W)"
               $ResultProps.Output_For_HTML1 = "Error"
               break
              }
          "1" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $False
               $ResultProps.Output_To_Log1 = "PVSCacheType: 1 - WC is set to Cache to PVS Server HD"
               $ResultProps.PVSCacheType = "1 - WC is set to Cache to PVS Server HD"
               $ResultProps.Output_For_HTML1 = "Error"
               break
              }
          "3" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $False
               $ResultProps.Output_To_Log1 = "3 - PVSCacheType: WC is set to Cache to Device Ram"
               $ResultProps.PVSCacheType = "3 - WC is set to Cache to Device Ram"
               $ResultProps.Output_For_HTML1 = "WARNING"
               break
              }
          "4" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $True
               $ResultProps.Output_To_Log1 = "PVSCacheType: 4 - WC is set to Cache to Device Hard Disk"
               $ResultProps.PVSCacheType = "4 - WC is set to Cache to Device Hard Disk"
               $ResultProps.Output_For_HTML1 = "WARNING"
               break
              }
          "7" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $False
               $ResultProps.Output_To_Log1 = "PVSCacheType: 7 - WC is set to Cache to PVS Server HD Persistent"
               $ResultProps.PVSCacheType = "7 - WC is set to Cache to PVS Server HD Persistent"
               $ResultProps.Output_For_HTML1 = "Error"
               break
              }
          "8" {
               $ResultProps.PVSCacheOnDeviceHardDisk = $True
               $ResultProps.Output_To_Log1 = "PVSCacheType: 8 - WC is set to Cache to Device Hard Disk Persistent"
               $ResultProps.PVSCacheType = "8 - WC is set to Cache to Device Hard Disk Persistent"
               $ResultProps.Output_For_HTML1 = "Error"
               break
              }
          Default {
               $ResultProps.PVSCacheOnDeviceHardDisk = $False
               $ResultProps.Output_To_Log1 = "PVSCacheType: Unknowm"
               $ResultProps.PVSCacheType = "$WriteCacheType - Unknown WC type"
               $ResultProps.Output_For_HTML1 = "Error"
               $cacheOnDeviceHardDisk = $False
             }
        }
        $DiskName = ($Personalityini | Select-String "DiskName" | ForEach-Object {$_.Line}).split('=')[1]
        $ResultProps.PVSDiskName = $DiskName
      } Else {
        $ListOfDDCs = $Personalityini -match 'ListOfDDCs='
        If ($ListOfDDCs -ne $null) {
          $ResultProps.IsMCS = $True
          $ResultProps.IsStandAlone = $False
        }
     }
    }
  } Else {
    $ResultProps.IsStandAlone = $True
    $ResultProps.DiskMode = "N/A"
  }
  If ($ResultProps.DiskMode -eq "S") {$ResultProps.DiskMode = "Standard"}
  If ($ResultProps.DiskMode -eq "P") {$ResultProps.DiskMode = "Private"}
  return $ResultProps
}

Function Get-WriteCacheDriveInfo {
  # This function will test...
  # For PVS:
  # - The size of the vdiskdif.vhdx write-cache file and available free space on the write-cache drive.
  # - The write cache drive is typically labeled "WCDisk", "Cache", "WriteCache", "Write Cache",
  #   "CacheDisk" (BISF), or "WRcache"
  # For MCSIO:
  # - The size of the mcsdif.vhdx write-cache file and available free space on the write-cache drive.
  # - The write cache drive is labeled "MCSWCDisk" or "CacheDisk" (BISF).
  # Written by Jeremy Saunders
  param (
         [switch]$UseWinRM,
         [int]$WinRMTimeoutSec=30,
         [string]$computername = "$env:computername",
         [switch]$IsPVS,
         [switch]$IsMCS,
         [string]$wcdrive = "D"
        )
  $results = @()
  $wcvolumename = @("MCSWCDisk","WCDisk","Cache","WriteCache","Write Cache","CacheDisk","WRcache")
  If (($IsPVS -AND $IsMCS) -OR ($IsPVS -eq $False -AND $IsMCS -eq $False)) {
    return
  }
  If ($IsPVS -OR $IsMCS) {
    If ($IsPVS -AND $IsMCS -eq $False) {
      $wcfile = "vdiskdif.vhdx"
      $ResultProps = @{
        WCdrivefreespace = "N/A"
        Output_For_HTML2 = "NEUTRAL"
        Output_To_Log2 = "WCdrivefreespace: Drive ${wcdrive} does not exist"
        vhdxSize_inMB = "N/A"
        Output_For_HTML3 = "NEUTRAL"
        Output_To_Log3 = "vdiskdifSize: N/A"
      }
    }
    If ($IsMCS -AND $IsPVS -eq $False) {
      $wcfile = "mcsdif.vhdx"
      $ResultProps = @{
        WCdrivefreespace = "N/A"
        Output_For_HTML2 = "NEUTRAL"
        Output_To_Log2 = "WCdrivefreespace: Drive ${wcdrive} does not exist"
        vhdxSize_inMB = "N/A"
        Output_For_HTML3 = "NEUTRAL"
        Output_To_Log3 = "mcsdifSize: N/A"
      }
    }
    $CanConnectToWriteCacheDrive = $False
    $HardDisk = $null
    Try {
      If ($UseWinRM) {
        $HardDisk = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $computername -Filter "DeviceID='${wcdrive}:' and DriveType='3'" -OperationTimeoutSec $WinRMTimeoutSec -ErrorAction Stop | Where-Object {$wcvolumename -contains $_.VolumeName} | Select-Object Size,FreeSpace
      } Else {
        $HardDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $computername -Filter "DeviceID='${wcdrive}:' and DriveType='3'" -ErrorAction Stop | Where-Object {$wcvolumename -contains $_.VolumeName} | Select-Object Size,FreeSpace
      }
      If ($null -ne $HardDisk) {
        $CanConnectToWriteCacheDrive = $True
        $DiskTotalSize = $HardDisk.Size 
        $DiskFreeSpace = $HardDisk.FreeSpace 
        $CacheDiskGB=[Math]::Round(($DiskFreeSpace/1073741824),2)
        $PercentageDS = (($DiskFreeSpace / $DiskTotalSize ) * 100)
        $PercentageDS = "{0:N2}" -f $PercentageDS 
        If ([int]$PercentageDS -ge 15) {
          $ResultProps.Output_To_Log2 = "WCdrivefreespace: Disk Free is normal [ $PercentageDS % ]"
          $ResultProps.WCdrivefreespace = "$PercentageDS %"
          $ResultProps.Output_For_HTML2 = "SUCCESS"
        } ElseIf (([int]$PercentageDS -lt 15) -and ([int]$PercentageDS -ge 10)) {
          $ResultProps.Output_To_Log2 = "WCdrivefreespace: Disk Free is Low [ $PercentageDS % ]"
          $ResultProps.WCdrivefreespace = "$PercentageDS %"
          $ResultProps.Output_For_HTML2 = "WARNING"
        } ElseIf ([int]$PercentageDS -lt 10) {
          $ResultProps.Output_To_Log2 = "WCdrivefreespace: Disk Free is Critical [ $PercentageDS % ]"
          $ResultProps.WCdrivefreespace = "$PercentageDS %"
          $ResultProps.Output_For_HTML2 = "ERROR"
        } ElseIf ([int]$PercentageDS -eq 0) {
          $ResultProps.Output_To_Log2 = "WCdrivefreespace: Disk Free test failed"
          $ResultProps.WCdrivefreespace = "$PercentageDS %"
          $ResultProps.Output_For_HTML2 = "ERROR"
        } Else {
          $ResultProps.Output_To_Log2 = "WCdrivefreespace: Disk Free is Critical [ $PercentageDS % ]"
          $ResultProps.WCdrivefreespace = "$PercentageDS %"
          $ResultProps.Output_For_HTML2 = "ERROR"
        }
      } Else {
        $ResultProps.Output_To_Log2 = "WCdrivefreespace: Failed to connect"
        $ResultProps.WCdrivefreespace = "Failed to connect"
        $ResultProps.Output_For_HTML2 = "ERROR"
      }
    }
    Catch [system.exception] {
      #$_.Exception.Message
    }
    If ($CanConnectToWriteCacheDrive) {
      $ErrorActionPreference = "stop"
      Try {
        If (Test-Path "filesystem::\\$computername\$wcdrive`$\$wcfile") {
          $ResultProps.vhdxSize_inMB = (Get-ChildItem -Path "filesystem::\\$computername\$wcdrive`$\$wcfile").Length/1024/1024
          $ResultProps.Output_To_Log3 = "vhdxSize: The `"$wcfile`" size is $($ResultProps.vhdxSize_inMB) MB"
        } Else {
          $ResultProps.Output_To_Log3 = "vhdxSize: The `"$wcfile`" file cannot be reached"
        }
      }
      Catch [system.exception] {
        #$_.Exception.Message
      }
      $ErrorActionPreference = "Continue"
    }
    $results += New-Object PsObject -Property $ResultProps
    return $results
  }   
}

#==============================================================================================

Function Get-ProfileAndUserEnvironmentManagementServiceStatus {
  # This function gets the Profile Management and User Environment Management services and their configuration status.
  # - Microsoft FSLogix
  # - Citrix Profile Management
  # - Citrix WEM
  # Written by Jeremy Saunders
  [CmdletBinding()]
  param (
         [string]$ComputerName = "$env:computername",
         [int]$WEMAgentRefresh = 30
        )
  $result = [PSCustomObject]@{
    ComputerName                   = $ComputerName
    FSLogixInstalled               = $false
    FSLogixServiceRunning          = $false
    FSLogixProfileEnabled          = $null
    FSLogixProfileType             = $null
    FSLogixProfileTypeDescription  = $null
    FSLogixOfficeEnabled           = $null
    FSLogixCCDLocations            = $null
    FSLogixVHDLocations            = $null
    FSLogixLogFilePath             = $null
    FSLogixRedirectionType         = $null
    UPMInstalled                   = $false
    UPMServiceRunning              = $false
    UPMServiceActive               = $null
    UPMPathToLogFile               = $null
    UPMPathToUserStore             = $null
    WEMInstalled                   = $false
    WEMServiceRunning              = $false
    WEMAgentRegistered             = $false
    WEMAgentConfigurationSets      = $null
    WEMAgentCacheSyncMode          = $null
    WEMAgentCachePath              = $null
  }
  $paramBundle = [PSCustomObject]@{
    result           =  $result
    WEMAgentRefresh  =  $WEMAgentRefresh
  }
  Try {
    Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
      param (
             $paramBundle
            )
      If ($null -ne $paramBundle) {
        $result = $paramBundle.result
        $WEMAgentRefresh = $paramBundle.WEMAgentRefresh
      }
      $frxsvcInstalled = $False
      $frxsvcRunning = $False
      $frxccdsInstalled = $False
      $frxccdsRunning = $False
      $WemAgentSvcInstalled = $False
      $WemAgentSvcRunning = $False
      $WemLogonSvcInstalled = $False
      $WemLogonSvcRunning = $False

      # Get-Service the services
      # - FSLogix Apps Services (frxsvc)
      # - FSLogix Cloud Caching Service (frxccds)
      # - Citrix Profile Management (ctxProfile)
      # - Citrix WEM Agent Host Service (WemAgentSvc)
      # - Citrix WEM User Logon Service (WemLogonSvc)
      try {
        $services = Get-Service -ErrorAction Stop | where-object {$_.Name -eq 'frxsvc' -OR $_.Name -eq 'frxccds' -OR $_.Name -eq 'ctxProfile' -OR $_.Name -eq 'WemAgentSvc' -OR $_.Name -eq 'WemLogonSvc'}
        $services | ForEach-Object {
          if ($_.Name -eq "frxsvc") {
            $frxsvcInstalled = $True
            if ($_.Status -Match "Running") {
              $frxsvcRunning = $True
            }
          }
          if ($_.Name -eq "frxccds") {
            $frxccdsInstalled = $True
            if ($_.Status -Match "Running") {
              $frxccdsRunning = $True
            }
          }
          if ($_.Name -eq "ctxProfile") {
            $result.UPMInstalled = $true
            if ($_.Status -Match "Running") {
              $result.UPMServiceRunning = ($_.Status -eq "Running")
            }
          }
          if ($_.Name -eq "WemAgentSvc") {
            $WemAgentSvcInstalled = $True
            if ($_.Status -Match "Running") {
              $WemAgentSvcRunning = $True
            }
          }
          if ($_.Name -eq "WemLogonSvc") {
            $WemLogonSvcInstalled = $True
            if ($_.Status -Match "Running") {
              $WemLogonSvcRunning = $True
            }
          }
        }

        If ($frxsvcInstalled -AND $frxccdsInstalled) {
          $result.FSLogixInstalled = $true
        }
        If ($frxsvcRunning -AND $frxccdsRunning) {
          $result.FSLogixServiceRunning = $true
        }
        If ($WemAgentSvcInstalled -AND $WemLogonSvcInstalled) {
          $result.WEMInstalled = $true
        }
        If ($WemAgentSvcRunning -AND $WemLogonSvcRunning) {
          $result.WEMServiceRunning = $true
        }
      }
      catch {
        #$_.Exception.Message
        $result.FSLogixInstalled = $false
        $result.FSLogixServiceRunning = $false
        $result.UPMInstalled = $false
        $result.UPMServiceRunning = $false
        $result.WEMInstalled = $false
        $result.WEMServiceRunning = $false
      }

      # Registry locations
      $FSLogixprofileReg = "HKLM:\SOFTWARE\FSLogix\Profiles"
      $FSLogixofficeReg  = "HKLM:\SOFTWARE\FSLogix\ODFC"
      $FSLogixloggingReg = "HKLM:\SOFTWARE\FSLogix\Logging"
      $UPMprofileReg = "HKLM:\SOFTWARE\Policies\Citrix\UserProfileManager"

      $ErrorActionPreference = "stop"
      If ($result.FSLogixInstalled -AND $result.FSLogixServiceRunning) {
        # FSLogix Profile Container settings
        try {
          $FSLogixprofileProps = Get-ItemProperty -Path $FSLogixprofileReg
          $result.FSLogixProfileEnabled = $FSLogixprofileProps.Enabled
          $result.FSLogixProfileType = $FSLogixprofileProps.ProfileType
          switch ($FSLogixprofileProps.ProfileType)
          {
               "0" {
                    $result.FSLogixProfileTypeDescription = "Standard connections - Normal profile behavior"
                    break
                   }
               "1" {
                    $result.FSLogixProfileTypeDescription = "Machine should only be the RW profile instance"
                    break
                   }
               "2" {
                    $result.FSLogixProfileTypeDescription = "Machine should only be the RO profile instance"
                    break
                   }
               "3" {
                    $result.FSLogixProfileTypeDescription = "Multiple concurrent connections - Machine should try to take the RW role and if it can't, it should fall back to a RO role"
                    break
                   }
           Default {
                    $result.FSLogixProfileTypeDescription = "Unknown profile type"
                   }
          }
          $result.FSLogixCCDLocations = ($FSLogixprofileProps.CCDLocations -join "; ")
          $result.FSLogixVHDLocations = ($FSLogixprofileProps.VHDLocations -join "; ")
          $result.FSLogixRedirectionType = $FSLogixprofileProps.VolumeType
        }
        catch {
         #$_.Exception.Message
        }
        # FSLogix Office Container settings
        try {
          $FSLogixofficeProps = Get-ItemProperty -Path $FSLogixofficeReg
          $result.FSLogixOfficeEnabled = $FSLogixofficeProps.Enabled
        }
        catch {
          #$_.Exception.Message
        }
        # FSLogix Log file path
        try {
          $FSLogixlogProps = Get-ItemProperty -Path $FSLogixloggingReg
          $result.FSLogixLogFilePath = $FSLogixlogProps.LogFile
        }
        catch {
          #$_.Exception.Message
        }
      }
      # UPM Config
      If ($result.UPMInstalled -AND $result.UPMServiceRunning) {
        try {
          $UPMProps = Get-ItemProperty -Path $UPMprofileReg
          If (![string]::IsNullOrEmpty($UPMProps.ServiceActive)) {
            $result.UPMServiceActive = $UPMProps.ServiceActive
          } Else {
            $result.UPMServiceActive = 1
          }
          $result.UPMPathToLogFile = $UPMProps.PathToLogFile
          $result.UPMPathToUserStore = $UPMProps.PathToUserStore
        }
        catch {
          #$_.Exception.Message
          $result.UPMServiceActive = 1
        }
      }
      $ErrorActionPreference = "Continue"

      If ($result.WEMInstalled -AND $result.WEMServiceRunning) {
        # How to test to see if WEM is configured. Thanks to Nick Panaccio from the World of EUC Slack group for his assistance.
        # The WEM Agent initiates a configuration settings refresh every 15 minutes by default, so we can check if it has successfully
        # registered with its configuration sets within the last 30 minutes via the WEM Agent Service event log. There are two failure
        # scenarios to be aware of for the WEMAgentRegistered property to remain as false:
        # 1) If the uptime of the machine is less than 30 minutes where the Agent hasn't yet registered, and therefore the event cannot
        #    be found. Hence why we look back 30 minutes instead of 15 to avoid missing a registration event due to a busy boot process
        #    and timing.
        # 2) The agent must be successfully registered with a "configuration set". If it's not registered with any configuration set,
        #    it will not be set to true.
        Try {
          $Registered = Get-WinEvent -LogName "WEM Agent Service" -ErrorAction Stop | Where { $_.Message -like "Agent successfully registered with configuration set*" -and $_.TimeCreated -gt (Get-Date).AddMinutes(-$WEMAgentRefresh) } | Select TimeCreated, Message | Select-Object -First 1
          # As the agent will refresh the cache (15 min default), we can also get the "Agent cache sync mode" and "Agent cache path".
          $CacheInfo = Get-WinEvent -LogName "WEM Agent Service" -ErrorAction Stop | Where { $_.Message -like "Agent cache info*" -and $_.TimeCreated -gt (Get-Date).AddMinutes(-$WEMAgentRefresh) } | Select TimeCreated, Message | Select-Object -First 1

          If ($null -ne $Registered) {
            $result.WEMAgentRegistered = $True
            $lines = ($Registered.Message) -split '\r?\n'
            ForEach ($line in $lines) {
              If ($line -like "Agent successfully registered*") {
                $result.WEMAgentConfigurationSets = $line.split(':',2)[1].Trim()
              } 
            }
          }
          If ($null -ne $CacheInfo) {
            $lines = ($CacheInfo.Message) -split '\r?\n'
            ForEach ($line in $lines) {
              If ($line -like "Agent cache sync mode*") {
                $result.WEMAgentCacheSyncMode = $line.split(':',2)[1].Trim()
              }
              If ($line -like "Agent cache path*") {
                $result.WEMAgentCachePath = $line.split(':',2)[1].Trim()
              }
            }
          }
        }
        Catch {
          #$_.Exception.Message
        }
      }
      return $result
    } -ArgumentList $paramBundle
  }
  Catch {
    #$_.Exception.Message
    # I have wrapped this in a try/catch because you may, on very rare occasions, get errors like "The WSMan service could not launch a host process
    # to process the given request. Make sure the WSMan provider host server and proxy are properly registered". I suspect it happens only when a
    # machine is interrogated whilst it's still starting up after a reboot.
    return $result
  }
}

#==============================================================================================

Function Get-CrowdStrikeServiceStatus {
  # This function gets the CrowdStrike service, System Driver, configuration data and installed version.
  # - We get the installation settings from the following registry key, which is created as as part of the install:
  #   - HKLM\SYSTEM\CrowdStrike\{9b03c1d9-3138-44ed-9fae-d9f4c034b88d}\{16e0423f-7058-48c9-a204-725362b67639}\Default
  #   - Company ID (CID) from the CI registry value.
  #   - VDI installation switch, which is used to prevent duplicate host names when using imaging technologies.
  #   - SensorGroupingTags - Assigned on each device either during installation or using CsSensorSettings.exe
  #                          Tags are case sensitive
  #                          To use multiple tags, separate with commas
  #   - Once the agent has registered with the server, we can also get the Device ID or Agent ID (AID) from the AG 
  #     registry value.
  # - Once the system driver starts the following key may be populated with more data:
  #   - HKLM\SYSTEM\CurrentControlSet\Services\CSAgent\Sim
  # We cannot get FalconGroupingTags from the agents. These are assigned in the UI or via API. They can be assigned
  # to devices in groups. To apply a FalconGroupingTag, use the Host Management screen and multi-select the target
  # devices, then use Actions > Add Falcon Grouping Tags
  # References:
  # - https://www.powershellgallery.com/packages/PSFalcon/2.1.9/Content/Public%5Cpsf-sensors.ps1
  # - https://community.automox.com/find-share-worklets-12/retaining-registry-key-with-hostname-2066
  # - https://www.reddit.com/r/crowdstrike/comments/186rw57/how_to_retrieve_crowdstrike_agent_idhost_id_from/?tl=fi
  # - https://blog.1password.com/how-to-tell-if-crowdstrike-falcon-sensor-is-running/
  # Written by Jeremy Saunders
  [CmdletBinding()]
  param (
         [string]$ComputerName
        )
  $result = [PSCustomObject]@{
    ComputerName           = $env:COMPUTERNAME
    CSFalconInstalled      = $false
    CSFalconServiceRunning = $false
    CSAgentInstalled       = $false
    CSAgentServiceRunning  = $false
    CID                    = $null
    AID                    = $null
    SensorGroupingTags     = $null
    VDI                    = $false
    InstalledVersion       = $null
  }
  $paramBundle = [PSCustomObject]@{
    result           =  $result
  }
  Try {
    Invoke-Command -ComputerName $ComputerName -ErrorAction Stop -ScriptBlock {
      param(
            $paramBundle
           )
      If ($null -ne $paramBundle) {
        $result = $paramBundle.result
      }
      $CSFalconInstalled = $False
      $CSFalconRunning = $False
      $CSAgentInstalled = $False
      $CSAgentRunning = $False

      # Get-Services and the Win32_SystemDriver:
      # - CrowdStrike Falcon Sensor Service (CSFalconService)
      # - CrowdStrike Falcon (CSAgent) is a driver with startup type of 1 (System).
      #   - We cannot use the Get-Service or Get-WmiObject/Get-CIMInstance Win32_Service class to see if it's running.
      #   - We need to use the Get-WmiObject/Get-CIMInstance Win32_SystemDriver class instead.
      #   Whilst we may not need to do this, or it could be seen as excessive, it may be a good way to pick up broken
      #   installs. So I do it anyway.
      try {
        $services = Get-Service -ErrorAction Stop | where-object {$_.Name -eq 'CSFalconService'}
        $services | ForEach-Object {
          if ($_.Name -eq "CSFalconService") {
            $CSFalconInstalled = $True
            if ($_.Status -Match "Running") {
              $CSFalconRunning = $True
            }
          }
        }
        If ($CSFalconInstalled) {
          $result.CSFalconInstalled = $true
        }
        If ($CSFalconRunning) {
          $result.CSFalconServiceRunning = $true
        }
      }
      catch {
        #$($Error[0].Exception.Message)
        $result.CSFalconInstalled = $false
        $result.CSFalconServiceRunning = $false
      }
      Try {
        $VerbosePreference = 'SilentlyContinue'
        $SystemDrivers = Get-CimInstance -ClassName Win32_SystemDriver -ErrorAction Stop | where-object {$_.Name -eq 'CSAgent'}
        $VerbosePreference = 'Continue'
        $SystemDrivers | ForEach-Object {
          if ($_.Name -eq "CSAgent") {
            $CSAgentInstalled = $True
            if ($_.State -Match "Running") {
              $CSAgentRunning = $True
            }
          }
        }
        If ($CSAgentInstalled) {
          $result.CSAgentInstalled = $true
        }
        If ($CSAgentRunning) {
          $result.CSAgentServiceRunning = $true
        }
      }
      Catch [System.Exception]{
        #$($Error[0].Exception.Message)
        $result.CSAgentInstalled = $false
        $result.CSAgentServiceRunning = $false
      }

      # Registry locations
      $CSFalconReg = "HKLM:\SYSTEM\CrowdStrike\{9b03c1d9-3138-44ed-9fae-d9f4c034b88d}\{16e0423f-7058-48c9-a204-725362b67639}\Default"

      $ErrorActionPreference = "stop"
      If ($result.CSFalconInstalled -AND $result.CSFalconServiceRunning -AND $result.CSAgentInstalled -AND $result.CSAgentServiceRunning) {
        try {
          $CSFalconProps = Get-ItemProperty -Path $CSFalconReg
          If ($CSFalconProps.AG -ne $null) {
            $result.AID = ([System.BitConverter]::ToString($CSFalconProps.AG) -replace '-','')
          }
          $result.CID = ([System.BitConverter]::ToString($CSFalconProps.CU) -replace '-','')
          $SensorGroupingTagsData = $CSFalconProps.GroupingTags
          if ($SensorGroupingTagsData) {
            if ($SensorGroupingTagsData.GetType().FullName -eq "System.String") {
              $result.SensorGroupingTags = $SensorGroupingTagsData
            } Else {
              # $CSFalconProps.GroupingTags.GetType().FullName -eq "System.Byte[]"
              $filteredBytes = $SensorGroupingTagsData | Where-Object { $_ -ne 0 }
              $i = 0
              $asciiLine = ""
              foreach ($byte in $filteredBytes) {
                $hex = '{0:X2}' -f $byte
                $char = if ($byte -ge 32 -and $byte -le 126) { [char]$byte } else { '.' }
                if ($i % 16 -eq 0) {
                  $asciiLine += ""
                }
                $asciiLine += $char
                $i++
              }
              if ($asciiLine) {
                $result.SensorGroupingTags = $asciiLine
              }
            }
          }
          If ($CSFalconProps.VDI -ne $null) {
            If ($CSFalconProps.VDI.GetType().FullName -eq "System.Int32") {
              If ($CSFalconProps.VDI[0] -eq 1) {
                $result.VDI = $True
              }
            } Else {
              $CSFalconProps.VDI.GetType().FullName -eq "System.Byte[]"
              # Check if the first byte (as an integer) is 1
              If ($CSFalconProps.VDI[0] -eq 1) {
                $result.VDI = $True
              }
            }
          }
        }
        catch {
          #$($Error[0].Exception.Message)
        }
      }
      $ErrorActionPreference = "Continue"

      $softwareDisplayName = "CrowdStrike Windows Sensor"
      $InstalledSoftware = @()
      $UninstallKeys=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                       "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                      )
      ForEach($UninstallKey in $UninstallKeys) {
        $ErrorActionPreference = "stop"
        Try {
          $InstalledSoftware += Get-ItemProperty $uninstallKey | Where-Object {$_.DisplayName -match $softwareDisplayName}
        }
        Catch {
          #$($Error[0].Exception.Message)
        }
      }
      If (($InstalledSoftware | Measure-Object).Count -eq 1) {
        If ($InstalledSoftware.DisplayVersion -ne $null) {
          $result.InstalledVersion = $InstalledSoftware.DisplayVersion
        }
      }
      $ErrorActionPreference = "Continue"
      return $result
    } -ArgumentList $paramBundle
  }
  Catch {
    #$_.Exception.Message
    # I have wrapped this in a try/catch because you may, on very rare occasions, get errors like "The WSMan service could not launch a host process
    # to process the given request. Make sure the WSMan provider host server and proxy are properly registered". I suspect it happens only when a
    # machine is interrogated whilst it's still starting up after a reboot.
    return $result
  }
}

#==============================================================================================

Function XDPing {
  # This function performs an XDPing to make sure the Delivery Controller or Cloud Connector is in a healthy state.
  # It tests whether the Broker service is reachable, listening and processing requests on its configured port.
  # We do this by issuing a blank HTTP POST requests to the Broker's Registrar service. Including "Expect: 100-continue"
  # in the body will ensure we receive a respose of "HTTP/1.1 100 Continue", which is what we use to verify that it's in
  # a healthy state.
  # You will notice that you can also pass proxy parameters to the function. This is for test and development ONLY. I
  # added this as I was using Fiddler to test the functionality and make sure the raw data sent was correctly formatted.
  # I decided to leave these parameters in the function so that others can learn and understand how this works.
  # To work out the best way to write this function I decompiled the VDAAssistant.Backend.dll from the Citrix Health
  # Assistant tool using JetBrains decompiler.
  # Written by Jeremy Saunders
  param(
    [Parameter(Mandatory=$True)][String]$ComputerName, 
    [Parameter(Mandatory=$True)][Int32]$Port,
    [String]$ProxyServer="", 
    [Int32]$ProxyPort,
    [Switch]$ConsoleOutput
  )
  $service = "http://${ComputerName}:${Port}/Citrix/CdsController/IRegistrar"
  $s = "POST $service HTTP/1.1`r`nContent-Type: application/soap+xml; charset=utf-8`r`nHost: ${ComputerName}:${Port}`r`nContent-Length: 1`r`nExpect: 100-continue`r`nConnection: Close`r`n`r`n"
  $log = New-Object System.Text.StringBuilder
  $log.AppendLine("Attempting an XDPing against $ComputerName on TCP port number $port") | Out-Null
  $listening = $false
  If ([string]::IsNullOrEmpty($ProxyServer)) {
    $ConnectToHost = $ComputerName
    [int]$ConnectOnPort = $Port
  } Else {
    $ConnectToHost = $ProxyServer
    [int]$ConnectOnPort = $ProxyPort
    $log.AppendLine("- Connecting via a proxy: ${ProxyServer}:${ProxyPort}") | Out-Null
  }
  try {
    $socket = New-Object System.Net.Sockets.Socket ([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp)
    try {
      $socket.Connect($ConnectToHost,$ConnectOnPort)
      if ($socket.Connected) {
        $log.AppendLine("- Socket connected") | Out-Null
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($s)
        $socket.Send($bytes) | Out-Null
        $log.AppendLine("- Sent the data") | Out-Null
        $numArray = New-Object byte[] 21
        $socket.ReceiveTimeout = 5000
        $socket.Receive($numArray) | Out-Null
        $log.AppendLine("- Received the following 21 byte array: " + [BitConverter]::ToString($numArray)) | Out-Null
        $strASCII = [System.Text.Encoding]::ASCII.GetString($numArray)
        $strUTF8 = [System.Text.Encoding]::UTF8.GetString($numArray)
        $log.AppendLine("- Converting to ASCII: `"$strASCII`"") | Out-Null
        $log.AppendLine("- Converting to UTF8: `"$strUTF8`"") | Out-Null
        $socket.Send([byte[]](32)) | Out-Null
        $log.AppendLine("- Sent a single byte with the value 32 (which represents the ASCII space character) to the connected socket.") | Out-Null
        $log.AppendLine("- This is done to gracefully signal the end of the communication.") | Out-Null
        $log.AppendLine("- This ensures it does not block/consume unnecessary requests needed by VDAs.") | Out-Null
        if ($strASCII.Trim().StartsWith("HTTP/1.1 100 Continue", [System.StringComparison]::CurrentCultureIgnoreCase)) {
          $listening = $true
          $log.AppendLine("- The service is listening and healthy") | Out-Null
        } else {
          $log.AppendLine("- The service is not listening") | Out-Null
        }
        try {
          $socket.Close()
          $log.AppendLine("- Socket closed") | Out-Null
        } catch {
          $log.AppendLine("- Failed to close socket") | Out-Null
          $log.AppendLine("- ERROR: $_") | Out-Null
        }
        $socket.Dispose()
      } else {
        $log.AppendLine("- Socket failed to connect") | Out-Null
      }
    } catch {
      $log.AppendLine("- Failed to connect to service") | Out-Null
      $log.AppendLine("- ERROR: $_") | Out-Null
    }
  } catch {
    $log.AppendLine("- Failed to create socket") | Out-Null
    $log.AppendLine("- ERROR: $_") | Out-Null
  }
  If ($ConsoleOutput) {
    Write-Host $log.ToString().TrimEnd()
  }
  return $listening
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
width: fit-content;
}
body {
margin-left: 5px;
margin-top: 5px;
margin-right: 0px;
margin-bottom: 10px;
table {
width: 100%
table-layout:fixed;
border: thin solid #000000;
}
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
param($fileName, $firstheaderName, $headerNames, $tablewidth)
$tableHeader = @"
  
<table width='$tablewidth'><tbody>
<tr bgcolor=#CCCCCC>
<td align='center'><strong>$firstheaderName</strong></td>
"@

$i = 0
while ($i -lt $headerNames.count) {
$headerName = $headerNames[$i]
$tableHeader += "<td align='center'><strong>$headerName</strong></td>"
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
# Replace a regular hyphen with a non-breaking hyphen to reduce text breaking to the next line.
$nonBreakingHyphen = [char]0x2011
If (![string]::IsNullOrEmpty($testResult)) {
$testResult = ($testResult -replace '-', $nonBreakingHyphen)
}
$tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")
}
$tableEntry += "</tr>"
}
$tableEntry | Out-File $fileName -append
}

Function writeDataSortedByHeaderName
{
param($data, $fileName, $headerNames, $headerToSortBy)

$tableEntry  =""
#  Hashtables are inherently unordered, meaning you cannot directly sort a hashtable and maintain the order. However, you
#  can achieve a sorted representation by extracting the data, sorting it, and then displaying or using it in the desired
#  order. We do this using the GetEnumerator() method, which converts the hashtable into a collection of key-value pairs
#  (DictionaryEntry objects), which can then be piped to Sort-Object to sort on Key (Name), Value, or a Value within the
#  Value. The Sort-Object cmdlet allows us to sort for multiple columns. Use commas to create a list of properties to sort
#  by, in order to precedence. So in this case we want to primarily sort on the $headerToSortBy column, and then by Key
#  (Name). If you want to preserve the order after sorting, you can then pipe it to a ForEach-Object to be stored into a
#  new ordered dictionary.
#$sortedData = $data.GetEnumerator() | Sort-Object {$_.Value.$headerToSortBy},Key
$sortedData = [ordered]@{}
$data.GetEnumerator() | Sort-Object {$_.Value.$headerToSortBy},Key | ForEach-Object {
  $sortedData[$_.Key] = $_.Value
}
$sortedData.Keys | ForEach-Object {
$tableEntry += "<tr>"
$computerName = $_
$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")
#$sortedData.$_.Keys | foreach {
$headerNames | ForEach-Object {
#"$computerName : $_" | LogMe -display
try {
if ($sortedData.$computerName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
elseif ($sortedData.$computerName.$_[0] -eq "WARNING") { $bgcolor = "#FF7700"; $fontColor = "#FFFFFF" }
elseif ($sortedData.$computerName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF" }
else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
$testResult = $sortedData.$computerName.$_[1]
}
catch {
$bgcolor = "#CCCCCC"; $fontColor = "#003399"
$testResult = ""
}
# Replace a regular hyphen with a non-breaking hyphen to reduce text breaking to the next line.
$nonBreakingHyphen = [char]0x2011
If (![string]::IsNullOrEmpty($testResult)) {
$testResult = ($testResult -replace '-', $nonBreakingHyphen)
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
param([string]$fileName,[switch]$cloud)
If ($cloud -eq $False) {
$thefooter = @"
</table>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#000000' size='2'>

<strong>Uptime Threshold: </strong> $maxUpTimeDays days <br>
<strong>Maximum Disconnect Time Threshold: </strong> $MaxDisconnectTimeInHours hours <br>
<strong>Database: </strong> $dbinfo <br>
<strong>LicenseServerName: </strong> $lsname <strong>LicenseServerPort: </strong> $lsport <br>
<strong>ConnectionLeasingEnabled: </strong> $CLeasing <br>
<strong>LocalHostCacheEnabled: </strong> $LHC <br>

</font>
</td>
</table>
</body>
</html>
"@
} Else {
$thefooter = @"
</table>
<table width='1200'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#000000' size='2'>

<strong>Uptime Threshold: </strong> $maxUpTimeDays days <br>
<strong>Maximum Disconnect Time Threshold: </strong> $MaxDisconnectTimeInHours hours <br>
<strong>ConnectionLeasingEnabled: </strong> $CLeasing <br>
<strong>LocalHostCacheEnabled: </strong> $LHC <br>

</font>
</td>
</table>
</body>
</html>
"@
}
$thefooter | Out-File $FileName -append
}

# ==============================================================================================

Function ToHumanReadable()
{
  param($timespan)
  
  If ($timespan.TotalHours -lt 1) {
    If ($timespan.Minutes -gt 0) {
      return $timespan.Minutes.ToString() + " minutes"
    } Else {
      return $timespan.Seconds.ToString() + " seconds"
    }
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

#==============================================================================================
# == MAIN SCRIPT ==
#==============================================================================================

"#### Begin with Citrix XenDestop / XenApp HealthCheck #########################################" | LogMe -display -progress

" " | LogMe -display -progress

# Log the loaded Citrix PS Snapins
"Listing all the active and registered Citrix PowerShell snap-ins that are loaded" | LogMe -display -progress
ForEach ($SnapIn in $SnapIns) {
  "- $($SnapIn.Name)" | LogMe -display -progress
}

" " | LogMe -display -progress

# get some Site info, which will be presented in footer
"Getting some Site information" | LogMe -display -progress
If ($CitrixCloudCheck -ne 1) {
  $dbinfo = Get-BrokerDBConnection -AdminAddress $AdminAddress
  "- Database: $($dbinfo)" | LogMe -display -progress
  $brokersiteinfos = Get-BrokerSite -AdminAddress $AdminAddress
  $sitename = $brokersiteinfos.Name
} Else {
  $dbinfo = $null
  $brokersiteinfos = Get-BrokerSite
  If ($AppendCCSiteIdToName) {
    $siteid = Get-CCSiteId -BearerToken:$BearerToken -CustomerID:$CustomerID
    $sitename = $brokersiteinfos.Name + "-" + $siteid
  } Else {
    $sitename = $brokersiteinfos.Name
  }
  $orchestrationstatus = Get-CCOrchestrationStatus -BearerToken:$BearerToken -CustomerID:$CustomerID
  $productinternalversion = $orchestrationstatus.ProductInternalVersion
  "- ProductInternalVersion: $productinternalversion" | LogMe -display -progress
  $productexternalversion = $orchestrationstatus.ProductExternalVersion
  "- ProductExternalVersion: $productexternalversion" | LogMe -display -progress
  $controllerversion = $productinternalversion
}
"- SiteName: $($sitename)" | LogMe -display -progress
$lsname = $brokersiteinfos.LicenseServerName
"- LicenseServerName: $($lsname)" | LogMe -display -progress
$lsport = $brokersiteinfos.LicenseServerPort
"- LicenseServerPort: $($lsport)" | LogMe -display -progress
$CLeasing = $brokersiteinfos.ConnectionLeasingEnabled
"- ConnectionLeasingEnabled: $($CLeasing)" | LogMe -display -progress
$LHC = $brokersiteinfos.LocalHostCacheEnabled
"- LocalHostCacheEnabled: $($LHC)" | LogMe -display -progress

" " | LogMe -display -progress

"Getting the Site maintenance information for last $($MaintenanceModeActionsFromPastinDays) days" | LogMe -display -progress
# Get the maintenance information from the Site using the Get-LogLowLevelOperation cmdlet
# Original code written by Stefan Beckmann: http://www.beckmann.ch/blog/2016/11/01/get-user-who-set-maintenance-mode-for-a-server-or-client/
# Enhanced by Jeremy Saunders (jeremy@jhouseconsulting.com) so we can also get information for who placed Delivery Groups and Hypervisor Connections
# into maintenance mode. Refer to the Where-Object filter for the Get-LogLowLevelOperation cmdlet. Added PropertyName and TargetType to the output.
# IMPORTANT: The further back you go with the $MaintenanceModeActionsFromPastinDays variable, the more records the Get-LogLowLevelOperation needs to
# enumerate. The record enumeration starts from the $StartDate. So if you reach the MaxRecordCount before the EndTime, you may be missing the newer
# records. To workaround this we are retieving records in batches of $maxmachines rows using the StartTime from the last record collected as a
# reference for the next batch.
# Citrix provide an example of how to do this here: https://www.citrix.com/blogs/2021/10/18/announcing-remote-powershell-sdk-record-limits/
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-$MaintenanceModeActionsFromPastinDays)
# Note that the Maintenance Mode property name is inconsistent across different object types (Machines, Delivery Groups and Hypervisor Connections),
# so we use an array to filter objects where the 'PropertyName' property is in a list of names.
$MaintModePropNames = @("MAINTENANCEMODE", "INMAINTENANCEMODE", "MaintenanceMode")
$LogEntrys = @()
$lastStartTime = $StartDate
while ($true) {
  $TotalEntries = $LogEntrys.Length
  $previousStartTime = $lastStartTime
  If ($CitrixCloudCheck -ne 1) {
    $LogEntrys += @(Get-LogLowLevelOperation -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Filter { StartTime -gt $lastStartTime -and EndTime -le $EndDate } -Sortby 'StartTime' -ErrorAction Stop | Where-Object { $_.Details.PropertyName -in $MaintModePropNames })
  } Else {
    $LogEntrys += @(Get-LogLowLevelOperation -MaxRecordCount $maxmachines -Filter { StartTime -gt $lastStartTime -and EndTime -le $EndDate } -Sortby 'StartTime' -ErrorAction Stop | Where-Object { $_.Details.PropertyName -in $MaintModePropNames })
  }
  if (($LogEntrys.Length - $TotalEntries) -eq 0) {
    break;
  }
  $lastStartTime = $LogEntrys[-1].StartTime
  If ($lastStartTime -eq $previousStartTime) {
    break;
  }
}
$LogEntrys = $LogEntrys | Sort-Object EndTime -Descending

# Build an object with the data for the output
[array]$Maintenance = @()
ForEach ($LogEntry in $LogEntrys) {
  $TempObj = New-Object -TypeName psobject -Property @{
    User = $LogEntry.User
    PropertyName = $LogEntry.Details.PropertyName
    TargetType = $LogEntry.Details.TargetType
    TargetName = $LogEntry.Details.TargetName
    NewValue = $LogEntry.Details.NewValue
    PreviousValue = $LogEntry.Details.PreviousValue
    StartTime = $LogEntry.Details.StartTime
    EndTime = $LogEntry.Details.EndTime
  } #TempObj
  $Maintenance += $TempObj
} #ForEach

" " | LogMe -display -progress

#== Controller Check ============================================================================================
"Check Controllers #############################################################################" | LogMe -display -progress

" " | LogMe -display -progress
  
$ControllerResults = @{}
If ($CitrixCloudCheck -ne 1) { 
  $Controllers = Get-BrokerController -AdminAddress $AdminAddress

  # Get first DDC version (should be all the same unless an upgrade is in progress)
  $ControllerVersion = $Controllers[0].ControllerVersion
  "Version: $controllerversion" | LogMe -display -progress
  
  "XenDesktop/XenApp Version above 7.x ($controllerversion) - XenApp and DesktopCheck will be performed" | LogMe -display -progress

  foreach ($Controller in $Controllers) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}
  
    #Name of $Controller
    $ControllerDNS = $Controller | ForEach-Object{ $_.DNSName }
    "Controller: $ControllerDNS" | LogMe -display -progress

    # Column IPv4Address
    if (!([string]::IsNullOrWhiteSpace($ControllerDNS))) {
      Try {
        $IPv4Address = ([System.Net.Dns]::GetHostAddresses($ControllerDNS) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) -join ", "
        "IPv4Address: $IPv4Address" | LogMe -display -progress
        $tests.IPv4Address = "NEUTRAL", $IPv4Address
      }
      Catch [System.Net.Sockets.SocketException] {
        "Failed to lookup host in DNS: $($_.Exception.Message)" | LogMe -display -warning
         $IsSeverityWarningLevel = $True
      }
      Catch {
        "An unexpected error occurred: $($_.Exception.Message)" | LogMe -display -warning
        $IsSeverityWarningLevel = $True
      }
    }

    #Test-Ping $Controller
    $result = Test-Ping -Target:$ControllerDNS -Timeout:200 -Count:3
    # Column Ping
    If ($result -eq "SUCCESS") {
      $tests.Ping = "SUCCESS", $result
    } Else {
      $tests.Ping = "NORMAL", $result
    }
    "Is Pingable: $result" | LogMe -display -progress

    $XDPing = XDPing -ComputerName:$ControllerDNS -Port:80
    If ($XDPing) {
      $tests.XDPing = "SUCCESS", "SUCCESS"
      "XDPing (health status of the CdsController Iregistrar service): SUCCESS" | LogMe -display -progress
    } Else {
      $tests.XDPing = "ERROR", "FAILED"
      "XDPing (health status of the CdsController Iregistrar service): FAILED" | LogMe -display -error
      "- If this fails and machines are still registering okay to $ControllerDNS, check that the firewall is not dropping the traffic" | LogMe -display -error
      "  - This test sends a blank HTTP POST requests to the Broker's Registrar service, including 'Expect: 100-continue' in the body" | LogMe -display -error
      "  - A response of 'HTTP/1.1 100 Continue' should be returned, which is what we use to verify that it's in a healthy state" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    $IsWinRMAccessible = IsWinRMAccessible -hostname:$ControllerDNS

    #State of this controller
    $ControllerState = $Controller | ForEach-Object{ $_.State }
    "State: $ControllerState" | LogMe -display -progress
    if ($ControllerState -ne "Active") {
      $tests.State = "ERROR", $ControllerState
      $IsSeverityErrorLevel = $True
    } else {
      $tests.State = "SUCCESS", $ControllerState
    }

    #DesktopsRegistered on this controller
    $ControllerDesktopsRegistered = $Controller | ForEach-Object{ $_.DesktopsRegistered }
    "Registered: $ControllerDesktopsRegistered" | LogMe -display -progress
    $tests.DesktopsRegistered = "NEUTRAL", $ControllerDesktopsRegistered
  
    #ActiveSiteServices on this controller
    $ActiveSiteServices = $Controller | ForEach-Object{ $_.ActiveSiteServices }
    "ActiveSiteServices $ActiveSiteServices" | LogMe -display -progress
    $tests.ActiveSiteServices = "NEUTRAL", $ActiveSiteServices

    # Check CrowdStrike State
    If ($ShowCrowdStrikeTests -eq 1) {
      If ($IsWinRMAccessible) {
        $return = Get-CrowdStrikeServiceStatus -ComputerName:$ControllerDNS
        If ($null -ne $return) {
          If ($return.CSFalconInstalled -AND $return.CSAgentInstalled) {
            "CrowdStrike Installed: True" | LogMe -display -progress
            "- CrowdStrike Windows Sensor Version: $($return.InstalledVersion)" | LogMe -display -progress
            $tests.CSVersion = "NORMAL", $return.InstalledVersion
            "- CrowdStrike Company ID (CID): $($return.CID)" | LogMe -display -progress
            $tests.CSCID = "NORMAL", $return.CID
            "- CrowdStrike Sensor Grouping Tags: $($return.SensorGroupingTags)" | LogMe -display -progress
            $tests.CSGroupTags = "NORMAL", $return.SensorGroupingTags
            "- CrowdStrike VDI switch: $($return.VDI)" | LogMe -display -progress
            If ($return.CSFalconServiceRunning -AND $return.CSAgentServiceRunning -AND (![string]::IsNullOrEmpty($return.AID))) {
              $tests.CSEnabled = "SUCCESS", $True
              "- CrowdStrike Agent ID (AID): $($return.AID)" | LogMe -display -progress
              $tests.CSAID = "NORMAL", $return.AID
            } Else {
              $tests.CSEnabled = "WARNING", $False
              If ([string]::IsNullOrEmpty($return.AID)) {
                "- CrowdStrike Agent ID (AID) is missing" | LogMe -display -warning
                $tests.CSAID = "NORMAL", "Missing"
              } Else {
                "- CrowdStrike is installed, but not running" | LogMe -display -warning
              }
              $IsSeverityWarningLevel = $True
            }
          } else {
            "CrowdStrike Installed: False" | LogMe -display -progress
          }
        } else {
          "Unable to get the CrowdStrike Service Status" | LogMe -display -error
          $IsSeverityErrorLevel = $True
        }
      }
    }

    #==============================================================================================
    #               CHECK CPU AND MEMORY USAGE
    #==============================================================================================

    # Get the CPU configuration and check the AvgCPU value for 5 seconds
    $CpuConfigAndUsage = Get-CpuConfigAndUsage -hostname:$ControllerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $CpuConfigAndUsage) {
      $AvgCPUval = $CpuConfigAndUsage.CpuUsage
      if( [int] $AvgCPUval -lt 75) { "CPU usage is normal [ $AvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -lt 85) { "CPU usage is medium [ $AvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$AvgCPUval %" }   	
      elseif([int] $AvgCPUval -lt 95) { "CPU usage is high [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
      else { "CPU usage is Critical [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }   
      $AvgCPUval = 0
      "CPU Configuration:" | LogMe -display -progress
      If ($CpuConfigAndUsage.LogicalProcessors -gt 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -progress
        $tests.LogicalProcessors = "NEUTRAL", $CpuConfigAndUsage.LogicalProcessors
      } ElseIf ($CpuConfigAndUsage.LogicalProcessors -eq 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -warning
        $tests.LogicalProcessors = "WARNING", $CpuConfigAndUsage.LogicalProcessors
        $IsSeverityWarningLevel = $True
      } Else {
        "- LogicalProcessors: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.Sockets -gt 0) {
        "- Sockets: $($CpuConfigAndUsage.Sockets)" | LogMe -display -progress
        $tests.Sockets = "NEUTRAL", $CpuConfigAndUsage.Sockets
      } Else {
        "- Sockets: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.CoresPerSocket -gt 0) {
        "- CoresPerSocket: $($CpuConfigAndUsage.CoresPerSocket)" | LogMe -display -progress
        $tests.CoresPerSocket = "NEUTRAL", $CpuConfigAndUsage.CoresPerSocket
      } Else {
        "- CoresPerSocket: Unable to detect." | LogMe -display -progress
      }
    } else {
      "Unable to get CPU configuration and usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check the Physical Memory usage       
    $UsedMemory = CheckMemoryUsage -hostname:$ControllerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $UsedMemory) {
      if( $UsedMemory -lt 75) { "Memory usage is normal [ $UsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$UsedMemory %" }
      elseif( [int] $UsedMemory -lt 85) { "Memory usage is medium [ $UsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$UsedMemory %" ; $IsSeverityWarningLevel= $True }   	
      elseif( [int] $UsedMemory -lt 95) { "Memory usage is high [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }
      elseif( [int] $UsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
      else { "Memory usage is Critical [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }   
      $UsedMemory = 0  
    } else {
      "Unable to get Memory usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Get the total Physical Memory
    $TotalPhysicalMemoryinGB = Get-TotalPhysicalMemory -hostname:$ControllerDNS -UseWinRM:$IsWinRMAccessible
    If ($TotalPhysicalMemoryinGB -ge 4) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -progress
      $tests.TotalPhysicalMemoryinGB = "NEUTRAL", $TotalPhysicalMemoryinGB
    } ElseIf ($TotalPhysicalMemoryinGB -ge 2) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -warning
      $tests.TotalPhysicalMemoryinGB = "WARNING", $TotalPhysicalMemoryinGB
      $IsSeverityWarningLevel = $True
    } Else {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -error
      $tests.TotalPhysicalMemoryinGB = "ERROR", $TotalPhysicalMemoryinGB
      $IsSeverityErrorLevel = $True
    }

    foreach ($disk in $diskLettersControllers)
    {
      # Check Disk Usage 
      "Checking free space on $($disk):" | LogMe -display
      $HardDisk = CheckHardDiskUsage -hostname:$ControllerDNS -deviceID:"$($disk):" -UseWinRM:$IsWinRMAccessible
      if ($null -ne $HardDisk) {	
        $XAPercentageDS = $HardDisk.PercentageDS
        $frSpace = $HardDisk.frSpace
        If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
        ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
        ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True } 
        ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" ; $IsSeverityWarningLevel = $True }     
        Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True }  
        $XAPercentageDS = 0
        $frSpace = 0
        $HardDisk = $null
      } else {
        "Unable to get Hard Disk usage" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    }

    # Column OSBuild 
    $return = Get-OSVersion -hostname:$ControllerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $return) {
      If ($return.Error -eq "Success") {
        $tests.OSCaption = "NEUTRAL", $return.Caption
        $tests.OSBuild = "NEUTRAL", $return.Version
        "OS Caption: $($return.Caption)" | LogMe -display -progress
        "OS Version: $($return.Version)" | LogMe -display -progress
      } Else {
        $tests.OSCaption = "ERROR", $return.Caption
        $tests.OSBuild = "ERROR", $return.Version
        "OS Test: $($return.Error)" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    } else {
      "Unable to get OS Version and Caption" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check uptime
    $hostUptime = Get-UpTime -hostname:$ControllerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $hostUptime) {
      if ($hostUptime.TimeSpan.days -lt $minUpTimeDaysDDC) {
        "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -warning
        $tests.Uptime = "WARNING", (ToHumanReadable($hostUptime.TimeSpan))
        $IsSeverityWarningLevel = $True
      } else {
        "Uptime: $(ToHumanReadable($hostUptime.TimeSpan))" | LogMe -display -progress
        $tests.Uptime = "SUCCESS", (ToHumanReadable($hostUptime.TimeSpan))
      }
    } else {
      "WinRM or WMI connection failed" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    " --- " | LogMe -display -progress
    #Fill $tests into array
    $ControllerResults.$ControllerDNS = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'DeliveryController' = $ControllerDNS
      }
      $ControllerResults.$ControllerDNS.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$ControllerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$ControllerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
  }#Close off foreach $Controller
}#Close off $CitrixCloudCheck

#== Cloud Connector Check ===========================================================================================
if ($CitrixCloudCheck -eq 1 -AND $ShowCloudConnectorTable -eq 1 ) {
  "Check Cloud Connector Servers #################################################################" | LogMe -display -progress

  " " | LogMe -display -progress

  $CCResults = @{}

  foreach ($CloudConnectorServer in $CloudConnectorServers) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}

    #Name of $CloudConnectorServer
    $CloudConnectorServerDNS = $CloudConnectorServer
    "Cloud Connector Server: $CloudConnectorServerDNS" | LogMe -display -progress

    # Column IPv4Address
    if (!([string]::IsNullOrWhiteSpace($CloudConnectorServerDNS))) {
      Try {
        $IPv4Address = ([System.Net.Dns]::GetHostAddresses($CloudConnectorServerDNS) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) -join ", "
        "IPv4Address: $IPv4Address" | LogMe -display -progress
        $tests.IPv4Address = "NEUTRAL", $IPv4Address
      }
      Catch [System.Net.Sockets.SocketException] {
        "Failed to lookup host in DNS: $($_.Exception.Message)" | LogMe -display -warning
         $IsSeverityWarningLevel = $True
      }
      Catch {
        "An unexpected error occurred: $($_.Exception.Message)" | LogMe -display -warning
        $IsSeverityWarningLevel = $True
      }
    }

    #Ping $CloudConnectorServer
    $result = Test-Ping -Target:$CloudConnectorServerDNS -Timeout:200 -Count:3
    # Column Ping
    If ($result -eq "SUCCESS") {
      $tests.Ping = "SUCCESS", $result
    } Else {
      $tests.Ping = "NORMAL", $result
    }
    "Is Pingable: $result" | LogMe -display -progress

    $XDPing = XDPing -ComputerName:$CloudConnectorServerDNS -Port:80
    If ($XDPing) {
      $tests.XDPing = "SUCCESS", "SUCCESS"
      "XDPing (health status of the CdsController Iregistrar service): SUCCESS" | LogMe -display -progress
    } Else {
      $tests.XDPing = "ERROR", "FAILED"
      "XDPing (health status of the CdsController Iregistrar service): FAILED" | LogMe -display -error
      "- If this fails and machines are still registering okay to $ControllerDNS, check that the firewall is not dropping the traffic" | LogMe -display -error
      "  - This test sends a blank HTTP POST requests to the Broker's Registrar service, including 'Expect: 100-continue' in the body" | LogMe -display -error
      "  - A response of 'HTTP/1.1 100 Continue' should be returned, which is what we use to verify that it's in a healthy state" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    $IsWinRMAccessible = IsWinRMAccessible -hostname:$CloudConnectorServerDNS

    # Check services
    # The Get-Service command with -ComputerName parameter made use of DCOM and such functionality is
    # removed from PowerShell 7. So we use the Invoke-Command, which uses WinRM to run a ScriptBlock
    # instead.
    If ($IsWinRMAccessible) {
      Try {
        $CCActiveSiteServices = Invoke-Command -ComputerName $CloudConnectorServerDNS -ErrorAction Stop -ScriptBlock{Get-Service |?{ ($_.Name -ilike "Citrix*") -and ($_.StartType -eq "Automatic") -and ($_.Status -ne "Running")}}
        # Check if there are any stopped services
        if ($CCActiveSiteServices) {
          # If there are stopped services, print the list of stopped services
          "The following services are not running:$(($CCActiveSiteServices).Name)" | LogMe -display -warning
          $NotRunning_Service = $CCActiveSiteServices | ForEach-Object { $_.Name }
          $tests.CitrixServices = "Warning","$NotRunning_Service"
          $IsSeverityWarningLevel = $True
        } else {
          # If no services are stopped, print success message
          "All services are running successfully." | LogMe -display -progress
          $tests.CitrixServices = "SUCCESS","OK"
        }
      }
      Catch {
        #"Error returned while checking the services" | LogMe -error; return 101
        $IsSeverityErrorLevel = $True
      }
    } Else {
      $tests.CitrixServices = "WARNING","Cannot connect via WinRM"
      "Cannot connect via WinRM" | LogMe -display -warning
      $IsSeverityWarningLevel = $True
    }

    # Check CrowdStrike State
    If ($ShowCrowdStrikeTests -eq 1) {
      If ($IsWinRMAccessible) {
        $return = Get-CrowdStrikeServiceStatus -ComputerName:$CloudConnectorServerDNS
        If ($null -ne $return) {
          If ($return.CSFalconInstalled -AND $return.CSAgentInstalled) {
            "CrowdStrike Installed: True" | LogMe -display -progress
            "- CrowdStrike Windows Sensor Version: $($return.InstalledVersion)" | LogMe -display -progress
            $tests.CSVersion = "NORMAL", $return.InstalledVersion
            "- CrowdStrike Company ID (CID): $($return.CID)" | LogMe -display -progress
            $tests.CSCID = "NORMAL", $return.CID
            "- CrowdStrike Sensor Grouping Tags: $($return.SensorGroupingTags)" | LogMe -display -progress
            $tests.CSGroupTags = "NORMAL", $return.SensorGroupingTags
            "- CrowdStrike VDI switch: $($return.VDI)" | LogMe -display -progress
            If ($return.CSFalconServiceRunning -AND $return.CSAgentServiceRunning -AND (![string]::IsNullOrEmpty($return.AID))) {
              $tests.CSEnabled = "SUCCESS", $True
              "- CrowdStrike Agent ID (AID): $($return.AID)" | LogMe -display -progress
              $tests.CSAID = "NORMAL", $return.AID
            } Else {
              $tests.CSEnabled = "WARNING", $False
              If ([string]::IsNullOrEmpty($return.AID)) {
                "- CrowdStrike Agent ID (AID) is missing" | LogMe -display -warning
                $tests.CSAID = "NORMAL", "Missing"
              } Else {
                "- CrowdStrike is installed, but not running" | LogMe -display -warning
              }
              $IsSeverityWarningLevel = $True
            }
          } else {
            "CrowdStrike Installed: False" | LogMe -display -progress
          }
        } else {
          "Unable to get the CrowdStrike Service Status" | LogMe -display -error
          $IsSeverityErrorLevel = $True
        }
      }
    }

    #==============================================================================================
    #               CHECK CPU AND MEMORY USAGE
    #==============================================================================================

    # Get the CPU configuration and check the AvgCPU value for 5 seconds
    $CpuConfigAndUsage = Get-CpuConfigAndUsage -hostname:$CloudConnectorServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $CpuConfigAndUsage) {
      $AvgCPUval = $CpuConfigAndUsage.CpuUsage
      if( [int] $AvgCPUval -lt 75) { "CPU usage is normal [ $AvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -lt 85) { "CPU usage is medium [ $AvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$AvgCPUval %" }   	
      elseif([int] $AvgCPUval -lt 95) { "CPU usage is high [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
      else { "CPU usage is Critical [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }   
      $AvgCPUval = 0
      "CPU Configuration:" | LogMe -display -progress
      If ($CpuConfigAndUsage.LogicalProcessors -gt 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -progress
        $tests.LogicalProcessors = "NEUTRAL", $CpuConfigAndUsage.LogicalProcessors
      } ElseIf ($CpuConfigAndUsage.LogicalProcessors -eq 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -warning
        $tests.LogicalProcessors = "WARNING", $CpuConfigAndUsage.LogicalProcessors
        $IsSeverityWarningLevel = $True
      } Else {
        "- LogicalProcessors: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.Sockets -gt 0) {
        "- Sockets: $($CpuConfigAndUsage.Sockets)" | LogMe -display -progress
        $tests.Sockets = "NEUTRAL", $CpuConfigAndUsage.Sockets
      } Else {
        "- Sockets: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.CoresPerSocket -gt 0) {
        "- CoresPerSocket: $($CpuConfigAndUsage.CoresPerSocket)" | LogMe -display -progress
        $tests.CoresPerSocket = "NEUTRAL", $CpuConfigAndUsage.CoresPerSocket
      } Else {
        "- CoresPerSocket: Unable to detect." | LogMe -display -progress
      }
    } else {
      "Unable to get CPU configuration and usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check the Physical Memory usage       
    $UsedMemory = CheckMemoryUsage -hostname:$CloudConnectorServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $UsedMemory) {
      if( $UsedMemory -lt 75) { "Memory usage is normal [ $UsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$UsedMemory %" }
      elseif( [int] $UsedMemory -lt 85) { "Memory usage is medium [ $UsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$UsedMemory %" ; $IsSeverityWarningLevel= $True }   	
      elseif( [int] $UsedMemory -lt 95) { "Memory usage is high [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }
      elseif( [int] $UsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
      else { "Memory usage is Critical [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }   
      $UsedMemory = 0  
    } else {
      "Unable to get Memory usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Get the total Physical Memory
    $TotalPhysicalMemoryinGB = Get-TotalPhysicalMemory -hostname:$CloudConnectorServerDNS -UseWinRM:$IsWinRMAccessible
    If ($TotalPhysicalMemoryinGB -ge 4) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -progress
      $tests.TotalPhysicalMemoryinGB = "NEUTRAL", $TotalPhysicalMemoryinGB
    } ElseIf ($TotalPhysicalMemoryinGB -ge 2) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -warning
      $tests.TotalPhysicalMemoryinGB = "WARNING", $TotalPhysicalMemoryinGB
      $IsSeverityWarningLevel = $True
    } Else {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -error
      $tests.TotalPhysicalMemoryinGB = "ERROR", $TotalPhysicalMemoryinGB
      $IsSeverityErrorLevel = $True
    }

    foreach ($disk in $diskLettersControllers)
    {
      # Check Disk Usage 
      "Checking free space on $($disk):" | LogMe -display
      $HardDisk = CheckHardDiskUsage -hostname:$CloudConnectorServerDNS -deviceID:"$($disk):" -UseWinRM:$IsWinRMAccessible
      if ($null -ne $HardDisk) {	
        $XAPercentageDS = $HardDisk.PercentageDS
        $frSpace = $HardDisk.frSpace
        If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
        ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
        ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True } 
        ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" ; $IsSeverityWarningLevel = $True }     
        Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True }  
        $XAPercentageDS = 0
        $frSpace = 0
        $HardDisk = $null
      } else {
        "Unable to get Hard Disk usage" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    }

    # Column OSBuild 
    $return = Get-OSVersion -hostname:$CloudConnectorServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $return) {
      If ($return.Error -eq "Success") {
        $tests.OSCaption = "NEUTRAL", $return.Caption
        $tests.OSBuild = "NEUTRAL", $return.Version
        "OS Caption: $($return.Caption)" | LogMe -display -progress
        "OS Version: $($return.Version)" | LogMe -display -progress
      } Else {
        $tests.OSCaption = "ERROR", $return.Caption
        $tests.OSBuild = "ERROR", $return.Version
        "OS Test: $($return.Error)" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    } else {
      "Unable to get OS Version and Caption" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check uptime
    $hostUptime = Get-UpTime -hostname:$CloudConnectorServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $hostUptime) {
      if ($hostUptime.TimeSpan.days -lt $minUpTimeDaysDDC) {
        "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -warning
        $tests.Uptime = "WARNING", (ToHumanReadable($hostUptime.TimeSpan))
        $IsSeverityWarningLevel = $True
      } else {
        "Uptime: $(ToHumanReadable($hostUptime.TimeSpan))" | LogMe -display -progress
        $tests.Uptime = "SUCCESS", (ToHumanReadable($hostUptime.TimeSpan))
      }
    } else {
      "WinRM or WMI connection failed" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    " --- " | LogMe -display -progress
    #Fill $tests into array
    $CCResults.$CloudConnectorServerDNS = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'CloudConnector' = $CloudConnectorServerDNS
      }
      $CCResults.$CloudConnectorServerDNS.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$CloudConnectorServerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$CloudConnectorServerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
  }#Close off foreach $CloudConnectorServer
}#Close off if $ShowCloudConnectorTable

#== Storefront Check ================================================================================================
If ($ShowStorefrontTable -eq 1) {
  "Check Storefront Servers ######################################################################" | LogMe -display -progress

  " " | LogMe -display -progress

  $SFResults = @{}

  foreach ($StoreFrontServer in $StoreFrontServers) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}

    #Name of $StoreFrontServer
    $StoreFrontServerDNS = $StoreFrontServer
    "Storefront Server: $StoreFrontServerDNS" | LogMe -display -progress

    # Column IPv4Address
    if (!([string]::IsNullOrWhiteSpace($StoreFrontServerDNS))) {
      Try {
        $IPv4Address = ([System.Net.Dns]::GetHostAddresses($StoreFrontServerDNS) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) -join ", "
        "IPv4Address: $IPv4Address" | LogMe -display -progress
        $tests.IPv4Address = "NEUTRAL", $IPv4Address
      }
      Catch [System.Net.Sockets.SocketException] {
        "Failed to lookup host in DNS: $($_.Exception.Message)" | LogMe -display -warning
         $IsSeverityWarningLevel = $True
      }
      Catch {
        "An unexpected error occurred: $($_.Exception.Message)" | LogMe -display -warning
        $IsSeverityWarningLevel = $True
      }
    }

    #Ping $StoreFrontServer
    $result = Test-Ping -Target:$StoreFrontServerDNS -Timeout:200 -Count:3
    # Column Ping
    If ($result -eq "SUCCESS") {
      $tests.Ping = "SUCCESS", $result
    } Else {
      $tests.Ping = "NORMAL", $result
    }
    "Is Pingable: $result" | LogMe -display -progress

    $IsWinRMAccessible = IsWinRMAccessible -hostname:$StoreFrontServerDNS

    # Check services
    # The Get-Service command with -ComputerName parameter made use of DCOM and such functionality is
    # removed from PowerShell 7. So we use the Invoke-Command, which uses WinRM to run a ScriptBlock
    # instead.
    If ($IsWinRMAccessible) {
      Try {
        $SFActiveSiteServices = Invoke-Command -ComputerName $StoreFrontServerDNS -ErrorAction Stop -ScriptBlock{Get-Service |?{ (($_.Name -ilike "Citrix*") -or ($_.Name -like "W3SVC*")) -and ($_.StartType -eq "Automatic") -and ($_.Status -ne "Running")}}
        # Check if there are any stopped services
        if ($SFActiveSiteServices) {
          # If there are stopped services, print the list of stopped services
          "The following services are not running:$(($CCActiveSiteServices).Name)" | LogMe -display -warning
          $NotRunning_Service = $SFActiveSiteServices | ForEach-Object { $_.Name }
          $tests.CitrixServices = "Warning","$NotRunning_Service"
          $IsSeverityWarningLevel = $True
        } else {
          # If no services are stopped, print success message
          "All services are running successfully." | LogMe -display -progress
          $tests.CitrixServices = "SUCCESS","OK"
        }
      }
      Catch {
        #"Error returned while checking the services" | LogMe -error; return 101
        $IsSeverityErrorLevel = $True
      }
    } Else {
      $tests.CitrixServices = "WARNING","Cannot connect via WinRM"
      "Cannot connect via WinRM" | LogMe -display -warning
      $IsSeverityWarningLevel = $True
    }

    # Check CrowdStrike State
    If ($ShowCrowdStrikeTests -eq 1) {
      If ($IsWinRMAccessible) {
        $return = Get-CrowdStrikeServiceStatus -ComputerName:$StoreFrontServerDNS
        If ($null -ne $return) {
          If ($return.CSFalconInstalled -AND $return.CSAgentInstalled) {
            "CrowdStrike Installed: True" | LogMe -display -progress
            "- CrowdStrike Windows Sensor Version: $($return.InstalledVersion)" | LogMe -display -progress
            $tests.CSVersion = "NORMAL", $return.InstalledVersion
            "- CrowdStrike Company ID (CID): $($return.CID)" | LogMe -display -progress
            $tests.CSCID = "NORMAL", $return.CID
            "- CrowdStrike Sensor Grouping Tags: $($return.SensorGroupingTags)" | LogMe -display -progress
            $tests.CSGroupTags = "NORMAL", $return.SensorGroupingTags
            "- CrowdStrike VDI switch: $($return.VDI)" | LogMe -display -progress
            If ($return.CSFalconServiceRunning -AND $return.CSAgentServiceRunning -AND (![string]::IsNullOrEmpty($return.AID))) {
              $tests.CSEnabled = "SUCCESS", $True
              "- CrowdStrike Agent ID (AID): $($return.AID)" | LogMe -display -progress
              $tests.CSAID = "NORMAL", $return.AID
            } Else {
              $tests.CSEnabled = "WARNING", $False
              If ([string]::IsNullOrEmpty($return.AID)) {
                "- CrowdStrike Agent ID (AID) is missing" | LogMe -display -warning
                $tests.CSAID = "NORMAL", "Missing"
              } Else {
                "- CrowdStrike is installed, but not running" | LogMe -display -warning
              }
              $IsSeverityWarningLevel = $True
            }
          } else {
            "CrowdStrike Installed: False" | LogMe -display -progress
          }
        } else {
          "Unable to get the CrowdStrike Service Status" | LogMe -display -error
          $IsSeverityErrorLevel = $True
        }
      }
    }

    #==============================================================================================
    #               CHECK CPU AND MEMORY USAGE
    #==============================================================================================

    # Get the CPU configuration and check the AvgCPU value for 5 seconds
    $CpuConfigAndUsage = Get-CpuConfigAndUsage -hostname:$StoreFrontServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $CpuConfigAndUsage) {
      $AvgCPUval = $CpuConfigAndUsage.CpuUsage
      if( [int] $AvgCPUval -lt 75) { "CPU usage is normal [ $AvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -lt 85) { "CPU usage is medium [ $AvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$AvgCPUval %" }   	
      elseif([int] $AvgCPUval -lt 95) { "CPU usage is high [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }
      elseif([int] $AvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
      else { "CPU usage is Critical [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$AvgCPUval %" }   
      $AvgCPUval = 0
      "CPU Configuration:" | LogMe -display -progress
      If ($CpuConfigAndUsage.LogicalProcessors -gt 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -progress
        $tests.LogicalProcessors = "NEUTRAL", $CpuConfigAndUsage.LogicalProcessors
      } ElseIf ($CpuConfigAndUsage.LogicalProcessors -eq 1) {
        "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -warning
        $tests.LogicalProcessors = "WARNING", $CpuConfigAndUsage.LogicalProcessors
        $IsSeverityWarningLevel = $True
      } Else {
        "- LogicalProcessors: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.Sockets -gt 0) {
        "- Sockets: $($CpuConfigAndUsage.Sockets)" | LogMe -display -progress
        $tests.Sockets = "NEUTRAL", $CpuConfigAndUsage.Sockets
      } Else {
        "- Sockets: Unable to detect." | LogMe -display -progress
      }
      If ($CpuConfigAndUsage.CoresPerSocket -gt 0) {
        "- CoresPerSocket: $($CpuConfigAndUsage.CoresPerSocket)" | LogMe -display -progress
        $tests.CoresPerSocket = "NEUTRAL", $CpuConfigAndUsage.CoresPerSocket
      } Else {
        "- CoresPerSocket: Unable to detect." | LogMe -display -progress
      }
    } else {
      "Unable to get CPU configuration and usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check the Physical Memory usage       
    $UsedMemory = CheckMemoryUsage -hostname:$StoreFrontServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $UsedMemory) {
      if( $UsedMemory -lt 75) { "Memory usage is normal [ $UsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$UsedMemory %" }
      elseif( [int] $UsedMemory -lt 85) { "Memory usage is medium [ $UsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$UsedMemory %" ; $IsSeverityWarningLevel= $True }   	
      elseif( [int] $UsedMemory -lt 95) { "Memory usage is high [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }
      elseif( [int] $UsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
      else { "Memory usage is Critical [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$UsedMemory %" ; $IsSeverityErrorLevel = $True }   
      $UsedMemory = 0  
    } else {
      "Unable to get Memory usage" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Get the total Physical Memory
    $TotalPhysicalMemoryinGB = Get-TotalPhysicalMemory -hostname:$StoreFrontServerDNS -UseWinRM:$IsWinRMAccessible
    If ($TotalPhysicalMemoryinGB -ge 4) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -progress
      $tests.TotalPhysicalMemoryinGB = "NEUTRAL", $TotalPhysicalMemoryinGB
    } ElseIf ($TotalPhysicalMemoryinGB -ge 2) {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -warning
      $tests.TotalPhysicalMemoryinGB = "WARNING", $TotalPhysicalMemoryinGB
      $IsSeverityWarningLevel = $True
    } Else {
      "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -error
      $tests.TotalPhysicalMemoryinGB = "ERROR", $TotalPhysicalMemoryinGB
      $IsSeverityErrorLevel = $True
    }

    foreach ($disk in $diskLettersControllers)
    {
      # Check Disk Usage 
      "Checking free space on $($disk):" | LogMe -display
      $HardDisk = CheckHardDiskUsage -hostname:$StoreFrontServerDNS -deviceID:"$($disk):" -UseWinRM:$IsWinRMAccessible
      if ($null -ne $HardDisk) {	
        $XAPercentageDS = $HardDisk.PercentageDS
        $frSpace = $HardDisk.frSpace
        If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
        ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "Err" ; $IsSeverityErrorLevel = $True }
        ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True } 
        ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" ; $IsSeverityWarningLevel = $True }     
        Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $IsSeverityErrorLevel = $True }  
        $XAPercentageDS = 0
        $frSpace = 0
        $HardDisk = $null
      } else {
        "Unable to get Hard Disk usage" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    }

    # Column OSBuild 
    $return = Get-OSVersion -hostname:$StoreFrontServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $return) {
      If ($return.Error -eq "Success") {
        $tests.OSCaption = "NEUTRAL", $return.Caption
        $tests.OSBuild = "NEUTRAL", $return.Version
        "OS Caption: $($return.Caption)" | LogMe -display -progress
        "OS Version: $($return.Version)" | LogMe -display -progress
      } Else {
        $tests.OSCaption = "ERROR", $return.Caption
        $tests.OSBuild = "ERROR", $return.Version
        "OS Test: $($return.Error)" | LogMe -display -error
        $IsSeverityErrorLevel = $True
      }
    } else {
      "Unable to get OS Version and Caption" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Check uptime
    $hostUptime = Get-UpTime -hostname:$StoreFrontServerDNS -UseWinRM:$IsWinRMAccessible
    If ($null -ne $hostUptime) {
      if ($hostUptime.TimeSpan.days -lt $minUpTimeDaysDDC) {
        "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -warning
        $tests.Uptime = "WARNING", (ToHumanReadable($hostUptime.TimeSpan))
        $IsSeverityWarningLevel = $True
      } else {
        "Uptime: $(ToHumanReadable($hostUptime.TimeSpan))" | LogMe -display -progress
        $tests.Uptime = "SUCCESS", (ToHumanReadable($hostUptime.TimeSpan))
      }
    } else {
      "WinRM or WMI connection failed" | LogMe -display -error
      $IsSeverityErrorLevel = $True
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    " --- " | LogMe -display -progress
    #Fill $tests into array
    $SFResults.$StoreFrontServerDNS = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'StoreFront' = $StoreFrontServerDNS
      }
      $SFResults.$StoreFrontServerDNS.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$StoreFrontServerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$StoreFrontServerDNS" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
  }#Close off foreach $StoreFrontServer
}#Close off if $ShowStorefrontTable

#== Citrix Licensing Check =========================================================================================

If ($CitrixCloudCheck -ne 1) {
  "Check Citrix Licensing ######################################################################" | LogMe -display -progress
  # ======= License Check ========
  if ($ShowCTXLicense -eq 1 ) {

    $UseWinRM = $True
    $myCollection = @()
    try {
      If ($UseWinRM) {
        $LicQuery = Get-CimInstance -namespace "ROOT\CitrixLicensing" -ComputerName $lsname -query "select * from Citrix_GT_License_Pool" -ErrorAction Stop | ? {$_.PLD -in $CTXLicenseMode}
      } Else {
        $LicQuery = Get-WmiObject -namespace "ROOT\CitrixLicensing" -ComputerName $lsname -query "select * from Citrix_GT_License_Pool" -ErrorAction Stop | ? {$_.PLD -in $CTXLicenseMode}
      }

      foreach ($group in $($LicQuery | group pld)) {
        $lics = $group | Select-Object -ExpandProperty group
        $i = 1

        $myArray_Count = 0
        $myArray_InUse = 0
        $myArray_Available = 0

        foreach ($lic in @($lics)) {
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
    catch {
      $myArray = "" | Select-Object LicenseServer,LicenceName,Count,InUse,Available
      $myArray.LicenseServer = $lsname
      $myArray.LicenceName = "n/a"
      $myArray.Count = "n/a"
      $myArray.InUse = "n/a"
      $myArray.Available = "n/a"
      $myCollection += $myArray 
    }

    $CTXLicResults = @{}

    foreach ($line in $myCollection) {
      $tests = @{}

      if ($line.LicenceName -eq "n/a") {
        $tests.LicenseServer ="error", $line.LicenseServer
        $tests.Count ="error", $line.Count
        $tests.InUse ="error", $line.InUse
        $tests.Available ="error", $line.Available
      } else {
        $tests.LicenseServer ="NEUTRAL", $line.LicenseServer
        $tests.Count ="NEUTRAL", $line.Count
        $tests.InUse ="NEUTRAL", $line.InUse
        $tests.Available ="NEUTRAL", $line.Available
      }
      $CTXLicResults.($line.LicenceName) =  $tests
    }
    If ($CTXLicResults.Count -eq 0) {
      "No license data was returned. This may be because License Activation Service (LAS) is enabled." | LogMe -display -progress
    }
  } else {"CTX License Check skipped because ShowCTXLicense = 0 " | LogMe -display -progress } #Close off $ShowCTXLicense
  " --- " | LogMe -display -progress
} #Close off $CitrixCloudCheck

#== Catalog Check ============================================================================================
"Check Catalog #################################################################################" | LogMe -display -progress
" " | LogMe -display -progress

$ActualExcludedCatalogs = @()
$ActualIncludedCatalogs = @()
$CatalogResults = @{}
If ($CitrixCloudCheck -ne 1) { 
  $Catalogs = Get-BrokerCatalog -AdminAddress $AdminAddress
} Else {
  $Catalogs = Get-BrokerCatalog
}

foreach ($Catalog in $Catalogs) {
  $IsSeverityErrorLevel = $False
  $IsSeverityWarningLevel = $False
  $tests = @{}

  #Name of MachineCatalog
  $FullCatalogNameIncAdminFolder = $Catalog | ForEach-Object{ $_.Name }
  "Catalog: $FullCatalogNameIncAdminFolder" | LogMe -display -progress

  # The CatalogName property is not always available across versions and configurations.
  Try {
    $CatalogName = $Catalog | ForEach-Object{ $_.CatalogName}
  } Catch {
    If ($FullCatalogNameIncAdminFolder.indexof('\\') -ne 0) {
      $CatalogName = ($FullCatalogNameIncAdminFolder -Split '\\')[-1]
    } Else {
      $CatalogName = $FullCatalogNameIncAdminFolder
    }
  }

  $Found = $False
  If ($ExcludedCatalogs.Count -gt 0) {
    If (!([String]::IsNullOrEmpty($ExcludedCatalogs[0]))) {
      ForEach ($ExcludedCatalog in $ExcludedCatalogs) {
        If ($FullCatalogNameIncAdminFolder -Like $ExcludedCatalog -OR $CatalogName -Like $ExcludedCatalog) {
          $Found = $True
          break
        }
      }
      If ($ExcludedCatalogs -contains $FullCatalogNameIncAdminFolder -OR $ExcludedCatalogs -contains $CatalogName) {
        $Found = $True
      }
    }
  }

  if ($Found) {
    $ActualExcludedCatalogs += $Catalog.Name
    "Excluded Catalog, skipping" | LogMe -display -progress
  } else {
     $ActualIncludedCatalogs += $Catalog.Name

    #CatalogAssignedCount
    # The number of assigned machines (machines that have been assigned to a user/users or a client name/address).
    $CatalogAssignedCount = $Catalog | ForEach-Object{ $_.AssignedCount }
    "Assigned: $CatalogAssignedCount" | LogMe -display -progress
    $tests.AssignedToUser = "NEUTRAL", $CatalogAssignedCount
  
    #CatalogUnassignedCount
    # The number of unassigned machines (machines not assigned to users).
    $CatalogUnAssignedCount = $Catalog | ForEach-Object{ $_.UnassignedCount }
    "Unassigned: $CatalogUnAssignedCount" | LogMe -display -progress
    $tests.NotToUserAssigned = "NEUTRAL", $CatalogUnAssignedCount
  
    # Assigned to DeliveryGroup
    # The number of machines in the catalog that are in a desktop group.
    $CatalogUsedCountCount = $Catalog | ForEach-Object{ $_.UsedCount }
    "Used: $CatalogUsedCountCount" | LogMe -display -progress
    $tests.AssignedToDG = "NEUTRAL", $CatalogUsedCountCount

    # Unassigned Machines
    # The number of available machines (those not in any desktop group) that are not assigned to users.
    $CatalogAvailableUnassignedCount = $Catalog | ForEach-Object{ $_.AvailableUnassignedCount }
    If ($CatalogAvailableUnassignedCount -eq 0) {
      "AvailableUnassignedCount: $CatalogAvailableUnassignedCount" | LogMe -display -progress
      $tests.Unassigned = "NEUTRAL", $CatalogAvailableUnassignedCount
    } Else {
      "AvailableUnassignedCount: $CatalogAvailableUnassignedCount" | LogMe -display -warning
      $tests.Unassigned = "WARNING", $CatalogAvailableUnassignedCount
      $IsSeverityWarningLevel = $True
    }

    #MinimumFunctionalLevel
    $MinimumFunctionalLevel = $Catalog | ForEach-Object{ $_.MinimumFunctionalLevel }
    "MinimumFunctionalLevel: $MinimumFunctionalLevel" | LogMe -display -progress
    $tests.MinimumFunctionalLevel = "NEUTRAL", $MinimumFunctionalLevel
  
    #AllocationType
    $CatalogAllocationType = $Catalog | ForEach-Object{ $_.AllocationType }
    "AllocationType: $CatalogAllocationType" | LogMe -display -progress
    $tests.AllocationType = "NEUTRAL", $CatalogAllocationType

    #ProvisioningType
    $CatalogProvisioningType = $Catalog | ForEach-Object{ $_.ProvisioningType }
    "ProvisioningType: $CatalogProvisioningType" | LogMe -display -progress
    $tests.ProvisioningType = "NEUTRAL", $CatalogProvisioningType

    # The GUID of the provisioning scheme associated with the catalog only applies if the provisioning type is MCS.
    # This may change in the future as Citrix further integrate PVS into the stack.
    If ($CatalogProvisioningType -eq "MCS") {
      $CatalogProvisioningSchemeId = $Catalog | ForEach-Object{ $_.ProvisioningSchemeId }

      #UsedMcsSnapshot 
      $UsedMcsSnapshot = ""
      $MCSInfo = $null
      $MasterImageVMDate = ""
      $UseFullDiskClone = ""
      $UseWriteBackCache = ""
      $WriteBackCacheMemSize = ""

      "ProvisioningSchemeId: $CatalogProvisioningSchemeId " | LogMe -display -progress
      Try {
        If ($CitrixCloudCheck -ne 1) {
          $MCSInfo = (Get-ProvScheme -AdminAddress $AdminAddress -ProvisioningSchemeUid $CatalogProvisioningSchemeId)
        } Else {
          $MCSInfo = (Get-ProvScheme -ProvisioningSchemeUid $CatalogProvisioningSchemeId)
        }
        "ProvisioningScheme Info:" | LogMe -display -progress
        # MachineProfile can be null, so we just wrap it in a try/catch to manage errors.
        Try {
          "- MachineProfile: $($MCSInfo.MachineProfile)" | LogMe -display -progress
        }
        Catch [system.exception] {
          #$_.Exception.Message
        }
        "- MasterImageVM: $($MCSInfo.MasterImageVM)" | LogMe -display -progress
        "- MasterImageVMDate: $($MCSInfo.MasterImageVMDate)" | LogMe -display -progress
        "- UseFullDiskCloneProvisioning: $($MCSInfo.UseFullDiskCloneProvisioning)" | LogMe -display -progress
        "- UseWriteBackCache: $($MCSInfo.UseWriteBackCache)" | LogMe -display -progress
        "- WriteBackCacheDiskSize: $($MCSInfo.WriteBackCacheDiskSize)" | LogMe -display -progress
        "- WriteBackCacheMemorySize:$( $MCSInfo.WriteBackCacheMemorySize)" | LogMe -display -progress
        "- WriteBackCacheDiskIndex: $($MCSInfo.WriteBackCacheDiskIndex)" | LogMe -display -progress
        # Note that the Get-ProvScheme cmdlet may does not yet have a parameter for "WriteBackCacheDiskLetter", even though this
        # was added for the New-ProvScheme cmdlet mid 2024 and supported from VDA version 2305 (CTX575525). This should not be
        # needed as the drive letter of MCSIO WBC disk is determined by Windows OS and is typically either D or E drive (the next
        # free drive letter after C). The Base Image Script Framework (BIS-F) should be used to help make this consistent as part
        # of your imaging standards. So we just wrap it in a try/catch to manage errors.
        Try {
          "- WriteBackCacheDiskLetter: $($MCSInfo.WriteBackCacheDiskLetter)" | LogMe -display -progress
        }
        Catch [system.exception] {
          #$_.Exception.Message
        }
        # WindowsActivationType is supported from 2303 and successive VDA versions. Any previous VDA version or if the vda is not
        # of Windows Operating System Type, the field would be "UnsupportedVDA". However, as it can be null, we just wrap it in a
        # try/catch to manage errors.
        Try {
          "- WindowsActivationType: $($MCSInfo.WindowsActivationType)" | LogMe -display -progress
        }
        Catch [system.exception] {
          #$_.Exception.Message
        }
        $UsedMcsSnapshot = $MCSInfo.MasterImageVM
        $UsedMcsSnapshot = $UsedMcsSnapshot.trimstart("XDHyp:\HostingUnits\")
        $UsedMcsSnapshot = $UsedMcsSnapshot.trimend(".template")
        $MasterImageVMDate = $MCSInfo.MasterImageVMDate
        $UseFullDiskClone = $MCSInfo.UseFullDiskCloneProvisioning
        $UseWriteBackCache = $MCSInfo.UseWriteBackCache
        $WriteBackCacheMemorySize = $MCSInfo.WriteBackCacheMemorySize
      }
      Catch [system.exception] {
        #$_.Exception.Message
      }
      "UsedMcsSnapshot: = $UsedMcsSnapshot"
      $tests.UsedMcsSnapshot  = "NEUTRAL", $UsedMcsSnapshot
      If (!([String]::IsNullOrEmpty($MasterImageVMDate))) {
        # Date format will always be MM/dd/yyyy so we specify that so that PowerShell doesn't parse it incorrectly, interpreting the
        # day and month around the opposite way.
        $format = "MM/dd/yyyy HH:mm:ss"
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $style = [System.Globalization.DateTimeStyles]::None
        $parsedDate = [datetime]::MinValue
        $success = [datetime]::TryParseExact($MasterImageVMDate, $format, $culture, $style, [ref]$parsedDate)
        $thresholdDate = (Get-Date).AddDays(-90)
        if ($parsedDate -lt $thresholdDate) {
          "The MasterImageVMDate is more than 90 days old." | LogMe -display -warning
          $tests.MasterImageVMDate = "WARNING", $MasterImageVMDate
          $IsSeverityWarningLevel = $True
        } else {
          $tests.MasterImageVMDate = "NEUTRAL", $MasterImageVMDate
        }
      }
      $tests.UseFullDiskClone = "NEUTRAL", $UseFullDiskClone
      $tests.UseWriteBackCache = "NEUTRAL", $UseWriteBackCache
      $tests.WriteBackCacheMemSize = "NEUTRAL", $WriteBackCacheMemSize
    } Else {
      "This is not an MCS provisioned catalog." | LogMe -display -progress
    }

    # Get agent versions of machines in the Machine Catalog so that we can see if the MinimumFunctionalLevel can be increased.
    $MachineAgentVersions = $null
    If ($CitrixCloudCheck -ne 1) {
      $MachineAgentVersions = Group-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Property AgentVersion -CatalogName $FullCatalogNameIncAdminFolder
    } Else {
      $MachineAgentVersions = Group-BrokerMachine -MaxRecordCount $maxmachines -Property AgentVersion -CatalogName $FullCatalogNameIncAdminFolder
    }
    # We want to get the lowest AgentVersion value in the Machine Catalog and then find if that matches the MinimumFunctionalLevel of the Machine Catalog.
    $RecommendedMinimumFunctionalLevel = $MinimumFunctionalLevel
    [version[]]$DerivedMinimumFunctionalLevels = @()
    ForEach($MachineAgentVersion in $MachineAgentVersions) {
      if ($null -ne $MachineAgentVersion) {
        if (![string]::IsNullOrWhiteSpace($MachineAgentVersion.Name)) {
          $TempMinimumFunctionalLevel = (Find-CitrixVersion -MatchByColumn:"MarketingProductVersion" -VersionToFind:"$($MachineAgentVersion.Name)").MinimumFunctionalLevel
          If ($TempMinimumFunctionalLevel -ne "N/A") {
            $DerivedMinimumFunctionalLevels += [version](Convert-FunctionalLevelToVersion $TempMinimumFunctionalLevel)
          }
        }
      }
    }
    $LowestSupportedMinimumFunctionalLevel = $DerivedMinimumFunctionalLevels | Sort-Object | Select-Object -First 1
    If ($LowestSupportedMinimumFunctionalLevel -gt (Convert-FunctionalLevelToVersion $MinimumFunctionalLevel)) {
      $RecommendedMinimumFunctionalLevel = (Convert-VersionToFunctionalLevel $LowestSupportedMinimumFunctionalLevel)
      "RecommendedMinimumFunctionalLevel: The recommended minimum functional level for this Machine Catalog should be changed to $RecommendedMinimumFunctionalLevel" | LogMe -display -warning
      $tests.RecommendedMinimumFunctionalLevel = "WARNING", $RecommendedMinimumFunctionalLevel
      $IsSeverityWarningLevel = $True
    } Else {
      "RecommendedMinimumFunctionalLevel: The minimum functional level for this Machine Catalog must remain at $RecommendedMinimumFunctionalLevel" | LogMe -display -progress
      $tests.RecommendedMinimumFunctionalLevel = "NEUTRAL", $RecommendedMinimumFunctionalLevel
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    "", ""
    $CatalogResults.$FullCatalogNameIncAdminFolder = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'MachineCatalog' = $FullCatalogNameIncAdminFolder
      }
      $CatalogResults.$FullCatalogNameIncAdminFolder.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$FullCatalogNameIncAdminFolder" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$FullCatalogNameIncAdminFolder" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
  }  
  " --- " | LogMe -display -progress
}

#== DeliveryGroups Check ============================================================================================
"Check Assigments #############################################################################" | LogMe -display -progress
  
" " | LogMe -display -progress

$ActualExcludedDeliveryGroups = @()
$ActualIncludedDeliveryGroups = @()
$AssigmentsResults = @{}
If ($CitrixCloudCheck -ne 1) {
  $Assigments = Get-BrokerDesktopGroup -AdminAddress $AdminAddress
} Else {
  $Assigments = Get-BrokerDesktopGroup
}

foreach ($Assigment in $Assigments) {
  $IsSeverityErrorLevel = $False
  $IsSeverityWarningLevel = $False
  $tests = @{}

  #Name of DeliveryGroup
  $FullDeliveryGroupNameIncAdminFolder = $Assigment | ForEach-Object{ $_.Name }
  "DeliveryGroup: $FullDeliveryGroupNameIncAdminFolder" | LogMe -display -progress

  $PublishedName = $Assigment | ForEach-Object{ $_.PublishedName }
  # The DesktopGroupName property is not always available across versions and configurations.
  Try {
    $DesktopGroupName = $Assigment | ForEach-Object{ $_.DesktopGroupName }
  }
  Catch {
    If ($FullDeliveryGroupNameIncAdminFolder.indexof('\\') -ne 0) {
      $DesktopGroupName = ($FullDeliveryGroupNameIncAdminFolder -Split '\\')[-1]
    } Else {
      $DesktopGroupName = $FullDeliveryGroupNameIncAdminFolder
    }
  }

  $Found = $False
  If ($ExcludedDeliveryGroups.Count -gt 0) {
    If (!([String]::IsNullOrEmpty($ExcludedDeliveryGroups[0]))) {
      ForEach ($ExcludedDeliveryGroup in $ExcludedDeliveryGroups) {
        If ($FullDeliveryGroupNameIncAdminFolder -Like $ExcludedDeliveryGroup -OR $PublishedName -Like $ExcludedDeliveryGroup -OR $DesktopGroupName -Like $ExcludedDeliveryGroup) {
          $Found = $True
          break
        }
      }
      if ($ExcludedDeliveryGroups -contains $FullDeliveryGroupNameIncAdminFolder -OR $ExcludedDeliveryGroups -contains $PublishedName -OR $ExcludedDeliveryGroups -contains $DesktopGroupName) {
        $Found = $True
      }
    }
  }
  if ($Found) {
    $ActualExcludedDeliveryGroups += $Assigment.Name
    "Excluded Delivery Group, skipping" | LogMe -display -progress
  } else {
    $ActualIncludedDeliveryGroups += $Assigment.Name

    #PublishedName
    $AssigmentDesktopPublishedName = $PublishedName
    "PublishedName: $AssigmentDesktopPublishedName" | LogMe -display -progress
    $tests.PublishedName = "NEUTRAL", $AssigmentDesktopPublishedName
  
    #DesktopsTotal
    $TotalDesktops = $Assigment | ForEach-Object{ $_.TotalDesktops }
    "TotalDesktops: $TotalDesktops" | LogMe -display -progress
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
    $IsSeverityErrorLevel = $True
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
          $IsSeverityWarningLevel = $True
        }
      } else {
        $tests.DesktopsFree = "NEUTRAL", "N/A"
      }
    }

    $AssigmentDesktopsenabled = $Assigment | ForEach-Object{ $_.enabled }
    "Enabled: $AssigmentDesktopsenabled" | LogMe -display -progress
    if ($AssigmentDesktopsenabled) { $tests.Enabled = "SUCCESS", "TRUE" }
    else { $tests.Enabled = "WARNING", "FALSE" ; $IsSeverityWarningLevel = $True }

    #inMaintenanceMode
    $AssigmentDesktopsinMaintenanceMode = $Assigment | ForEach-Object{ $_.inMaintenanceMode }
    "inMaintenanceMode: $AssigmentDesktopsinMaintenanceMode" | LogMe -display -progress
    if ($AssigmentDesktopsinMaintenanceMode) {
      $objMaintenance = $null
      Try {
        $objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $Assigment.Name.ToUpper() } | Select-Object -First 1
      }
      Catch {
        # Avoid the error "The property 'TargetName' cannot be found on this object."
      }
      If ($null -ne $objMaintenance){$AssigmentDesktopsinMaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$AssigmentDesktopsinMaintenanceModeOn = "ON"}
      # The Get-LogLowLevelOperation cmdlet will tell us who placed a Delivery Group into maintanance mode. However, the Get-BrokerDesktopGroup
      # cmdlet will provide the MaintenanceReason, where the underlying reason, if manually entered, is stored in the MetadataMap property as
      # part of a Dictionary object.
      $MetadataMapDictionary = $Assigment | ForEach-Object{ $_.MetadataMap }
      foreach ($key in $MetadataMapDictionary.Keys) {
        if ($key -eq "MaintenanceModeMessage") {
          $AssigmentDesktopsinMaintenanceModeOn = $AssigmentDesktopsinMaintenanceModeOn + ", " + $MetadataMapDictionary[$key]
        }
      }
      "MaintenanceModeInfo: $AssigmentDesktopsinMaintenanceModeOn" | LogMe -display -progress
      $tests.MaintenanceMode = "WARNING", $AssigmentDesktopsinMaintenanceModeOn
      $IsSeverityWarningLevel = $True
    }
    else { $tests.MaintenanceMode = "SUCCESS", "OFF" }

    #DesktopsUnregistered
    $AssigmentDesktopsUnregistered = $Assigment | ForEach-Object{ $_.DesktopsUnregistered }
    "DesktopsUnregistered: $AssigmentDesktopsUnregistered" | LogMe -display -progress    
    if ($AssigmentDesktopsUnregistered -gt 0 ) {
      "DesktopsUnregistered > 0 ! ($AssigmentDesktopsUnregistered)" | LogMe -display -warning
      $tests.DesktopsUnregistered = "WARNING", $AssigmentDesktopsUnregistered
      $IsSeverityWarningLevel = $True
    } else {
      $tests.DesktopsUnregistered = "SUCCESS", $AssigmentDesktopsUnregistered
      "DesktopsUnregistered <= 0 ! ($AssigmentDesktopsUnregistered)" | LogMe -display -progress
    }

    $DesktopsPowerStateUnknown = 0
    If ($CitrixCloudCheck -ne 1) {
      #$DesktopsPowerStateUnknown = (Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -PowerState Unknown | Measure-Object).Count
      $DesktopsPowerStateUnknown = (Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -Filter "PowerState -eq 'Unknown'" | Measure-Object).Count
    } Else {
      #$DesktopsPowerStateUnknown = (Get-BrokerMachine -MaxRecordCount $maxmachines -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -PowerState Unknown | Measure-Object).Count
      $DesktopsPowerStateUnknown = (Get-BrokerMachine -MaxRecordCount $maxmachines -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -Filter "PowerState -eq 'Unknown'" | Measure-Object).Count
    }
    if ($DesktopsPowerStateUnknown -eq 0 ) {
      "DesktopsPowerStateUnknown: $($DesktopsPowerStateUnknown)" | LogMe -display -progress
      $tests.DesktopsPowerStateUnknown = "SUCCESS", $DesktopsPowerStateUnknown
    } else {
      "DesktopsPowerStateUnknown: $($DesktopsPowerStateUnknown)" | LogMe -display -error
      $tests.DesktopsPowerStateUnknown = "ERROR", $DesktopsPowerStateUnknown
      $IsSeverityErrorLevel = $True
    }

    $DesktopsNotUsedLast90Days = 0
    If ($CitrixCloudCheck -ne 1) {
      $DesktopsNotUsedLast90Days = (Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -Filter "LastConnectionTime -lt '-90'" | Measure-Object).Count
    } Else {
      $DesktopsNotUsedLast90Days = (Get-BrokerMachine -MaxRecordCount $maxmachines -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder -Filter "LastConnectionTime -lt '-90'" | Measure-Object).Count
    }
    if ($DesktopsNotUsedLast90Days -eq 0 ) {
      "DesktopsNotUsedLast90Days: $($DesktopsNotUsedLast90Days)" | LogMe -display -progress
      $tests.DesktopsNotUsedLast90Days = "SUCCESS", $DesktopsNotUsedLast90Days
    } else {
      "DesktopsNotUsedLast90Days: $($DesktopsNotUsedLast90Days)" | LogMe -display -warning
      $tests.DesktopsNotUsedLast90Days = "WARNING", $DesktopsNotUsedLast90Days
      $IsSeverityWarningLevel = $True
    }

    # Get percentage of machines in the Delivery Group that are in Maintenance Mode.
    # Mark as warning if greater than 0% and error if 50% or greater.
    $MachinesInMaintMode = 0
    If ($CitrixCloudCheck -ne 1) {
      $MachineMaintModeStatus = Group-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Property InMaintenanceMode -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder
    } Else {
      $MachineMaintModeStatus = Group-BrokerMachine -MaxRecordCount $maxmachines -Property InMaintenanceMode -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder
    }
    ForEach($MaintModeStatus in $MachineMaintModeStatus) {
      If ($MaintModeStatus.Name -eq "True") {
        $MachinesInMaintMode = $MaintModeStatus.Count
      }
    }
    [float]$PercentageOfMachinesInMaintMode = 0.00
    If ([int]$TotalDesktops -gt 0) {
      $PercentageOfMachinesInMaintMode = (($MachinesInMaintMode / [int]$TotalDesktops) * 100)
      [float]$PercentageOfMachinesInMaintMode = "{0:N2}" -f $PercentageOfMachinesInMaintMode
    }
    If ($PercentageOfMachinesInMaintMode -lt [float]50.00) {
      If ($PercentageOfMachinesInMaintMode -eq [float]00.00) {
        "PercentageOfMachinesInMaintMode: $($PercentageOfMachinesInMaintMode)" | LogMe -display -progress
        $tests.PercentageOfMachinesInMaintMode = "SUCCESS", $PercentageOfMachinesInMaintMode
        "MachinesInMaintMode: $($MachinesInMaintMode)" | LogMe -display -progress
        $tests.MachinesInMaintMode = "SUCCESS", $MachinesInMaintMode
      } Else {
        "PercentageOfMachinesInMaintMode: $($PercentageOfMachinesInMaintMode)" | LogMe -display -warning
        $tests.PercentageOfMachinesInMaintMode = "WARNING", $PercentageOfMachinesInMaintMode
        "MachinesInMaintMode: $($MachinesInMaintMode)" | LogMe -display -warning
        $tests.MachinesInMaintMode = "WARNING", $MachinesInMaintMode
        $IsSeverityWarningLevel = $True
      }
    } Else {
      "PercentageOfMachinesInMaintMode: $($PercentageOfMachinesInMaintMode)" | LogMe -display -error
      $tests.PercentageOfMachinesInMaintMode = "ERROR", $PercentageOfMachinesInMaintMode
      "MachinesInMaintMode: $($MachinesInMaintMode)" | LogMe -display -error
      $tests.MachinesInMaintMode = "ERROR", $MachinesInMaintMode
      $IsSeverityErrorLevel = $True
    }

    # Get agent versions of machines in the Delivery Group so that we can see if the MinimumFunctionalLevel can be increased.
    $MachineAgentVersions = $null
    If ($CitrixCloudCheck -ne 1) {
      $MachineAgentVersions = Group-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Property AgentVersion -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder
    } Else {
      $MachineAgentVersions = Group-BrokerMachine -MaxRecordCount $maxmachines -Property AgentVersion -DesktopGroupName $FullDeliveryGroupNameIncAdminFolder
    }
    # We want to get the lowest AgentVersion value in the Delivery Group and then find if that matches the MinimumFunctionalLevel of the Delivery Group.
    $RecommendedMinimumFunctionalLevel = $MinimumFunctionalLevel
    [version[]]$DerivedMinimumFunctionalLevels = @()
    ForEach($MachineAgentVersion in $MachineAgentVersions) {
      if ($null -ne $MachineAgentVersion) {
        if (![string]::IsNullOrWhiteSpace($MachineAgentVersion.Name)) {
          $TempMinimumFunctionalLevel = (Find-CitrixVersion -MatchByColumn:"MarketingProductVersion" -VersionToFind:"$($MachineAgentVersion.Name)").MinimumFunctionalLevel
          If ($TempMinimumFunctionalLevel -ne "N/A") {
            $DerivedMinimumFunctionalLevels += [version](Convert-FunctionalLevelToVersion $TempMinimumFunctionalLevel)
          }
        }
      }
    }
    $LowestSupportedMinimumFunctionalLevel = $DerivedMinimumFunctionalLevels | Sort-Object | Select-Object -First 1
    If ($LowestSupportedMinimumFunctionalLevel -gt (Convert-FunctionalLevelToVersion $MinimumFunctionalLevel)) {
      $RecommendedMinimumFunctionalLevel = (Convert-VersionToFunctionalLevel $LowestSupportedMinimumFunctionalLevel)
      "RecommendedMinimumFunctionalLevel: The recommended minimum functional level for this Delivery Group should be changed to $RecommendedMinimumFunctionalLevel" | LogMe -display -warning
      $tests.RecommendedMinimumFunctionalLevel = "WARNING", $RecommendedMinimumFunctionalLevel
      $IsSeverityWarningLevel = $True
    } Else {
      "RecommendedMinimumFunctionalLevel: The minimum functional level for this Delivery Group must remain at $RecommendedMinimumFunctionalLevel" | LogMe -display -progress
      $tests.RecommendedMinimumFunctionalLevel = "NEUTRAL", $RecommendedMinimumFunctionalLevel
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    #Fill $tests into array
    $AssigmentsResults.$FullDeliveryGroupNameIncAdminFolder = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'DeliveryGroup' = $FullDeliveryGroupNameIncAdminFolder
      }
      $AssigmentsResults.$FullDeliveryGroupNameIncAdminFolder.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$FullDeliveryGroupNameIncAdminFolder" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$FullDeliveryGroupNameIncAdminFolder" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
  }
  " --- " | LogMe -display -progress
}

#== Get Broker Tags ================================================================================================
"Get Boker Tags ##############################################################################" | LogMe -display -progress

# The Broker Tags are collected to create the $ActualExcludedBrokerTags and $ActualIncludedBrokerTags arrays used for filtering

$ActualExcludedBrokerTags = @()
$ActualIncludedBrokerTags = @()
$BrkrTagsResults = @{}
$BrkrTags = $null
If ($CitrixCloudCheck -ne 1) {
  $BrkrTags = Get-BrokerTag -MaxRecordCount $maxmachines -AdminAddress $AdminAddress
} Else {
  $BrkrTags = Get-BrokerTag -MaxRecordCount $maxmachines
}

If ($null -ne $BrkrTags) {

  foreach ($BrkrTag in $BrkrTags) {
    $tests = @{}

    #Name of Tag
    $Tag = $BrkrTag | ForEach-Object{ $_.Name }
    "Tag: $Tag" | LogMe -display -progress

    $Found = $False
    If ($ExcludedTags.Count -gt 0) {
      If (!([String]::IsNullOrEmpty($ExcludedTags[0]))) {
        ForEach ($ExcludedTag in $ExcludedTags) {
          If ($Tag -Like $ExcludedTag) {
            $Found = $True
            break
          }
        }
        if ($ExcludedTags -contains $Tag) {
          $Found = $True
        }
      }
    }
    if ($Found) {
      $ActualExcludedBrokerTags += $BrkrTag.Name
      "Excluded Tag, skipping" | LogMe -display -progress
    } else {

      $ActualIncludedBrokerTags += $BrkrTag.Name

      $Description = $BrkrTag.Description
      "Description: $($BrkrTag.Description)" | LogMe -display -progress

      # Add the SiteName to the tests for the Syslog output
      $tests.SiteName = "NORMAL", $sitename

      #Fill $tests into array
      $BrkrTagsResults.$Tag = $tests
    }

   " --- " | LogMe -display -progress
  }
} Else {
  " --- " | LogMe -display -progress
}

#== Broker Connection Failure Results ===============================================================================
"Check Broker Connection Failures #############################################################" | LogMe -display -progress

if($ShowBrokerConnectionFailuresTable -eq 1 ) {

  $BrokerConnectionLogResults = @{}

  # Get the total FAILED Connections from the last x hours
  $BrokerConnectionFailuresForLastxHours = [DateTime]::Now - [TimeSpan]::FromHours($BrokerConnectionFailuresinHours)

  $BrkrConFailures = $null
  If ($CitrixCloudCheck -ne 1) {
    $BrkrConFailures = Get-BrokerConnectionLog -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Filter {BrokeringTime -gt $BrokerConnectionFailuresForLastxHours -and ConnectionFailureReason -ne 'None' -and ConnectionFailureReason -ne $null}
  } Else {
    $BrkrConFailures = Get-BrokerConnectionLog -MaxRecordCount $maxmachines -Filter {BrokeringTime -gt $BrokerConnectionFailuresForLastxHours -and ConnectionFailureReason -ne 'None' -and ConnectionFailureReason -ne $null}
  }

  If ($null -ne $BrkrConFailures) {

    foreach ($BrkrConFailure in $BrkrConFailures) {
      $IsSeverityErrorLevel = $False
      $IsSeverityWarningLevel = $False
      $tests = @{}

      # Machine DNS Name
      $MachineDNSName = $BrkrConFailure.MachineDNSName
      "MachineDNSName: $($BrkrConFailure.MachineDNSName)" | LogMe -display -progress

      $validMachine = $null
      If ($CitrixCloudCheck -ne 1) {
        $validMachine = Get-BrokerMachine -DNSName $MachineDNSName -AdminAddress $AdminAddress | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
      } Else {
        $validMachine = Get-BrokerMachine -DNSName $MachineDNSName | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
      }
      If ($null -eq $validMachine) {
        "Excluded, skipping this machine as it's in either the ExludedCatalogs, ExcludedDeliveryGroups and ExcludedTags" | LogMe -display -progress
      } Else {

        $BrokeringTime = $BrkrConFailure.BrokeringTime
        "BrokeringTime: $($BrkrConFailure.BrokeringTime)" | LogMe -display -progress
        $tests.BrokeringTime = "NEUTRAL", $BrokeringTime

        $ConnectionFailureReason = $BrkrConFailure.ConnectionFailureReason
        # As per Article ID: CTX137378, Failure Reasons and Causes:
        # - None - No failure (successful connection - session went active).
        # - SessionPreparation - Failure to prepare session, typically Virtual Desktop Agent (VDA) refused 'prepare' call, or communication error on prepare call to VDA.
        # - RegistrationTimeout - Timeout while waiting for worker to register when worker is being spun up to satisfy the launch.
        # - ConnectionTimeout - Timeout while waiting for client to connect to VDA after successful brokering part.
        # - Licensing - Licensing issue (for example - unable to verify a license).
        # - Ticketing - Failure during ticketing, indicating that the client connection to VDA does not match the brokered request.
        # - Other - Other
        "ConnectionFailureReason: $($BrkrConFailure.ConnectionFailureReason)" | LogMe -display -progress
        $tests.ConnectionFailureReason = "WARNING", $ConnectionFailureReason
        $IsSeverityWarningLevel = $True

        $BrokeringUserName = $BrkrConFailure.BrokeringUserName
        "BrokeringUserName: $($BrkrConFailure.BrokeringUserName)" | LogMe -display -progress
        $tests.BrokeringUserName = "NEUTRAL", $BrokeringUserName

        $BrokeringUserUPN = $BrkrConFailure.BrokeringUserUPN
        "BrokeringUserUPN: $($BrkrConFailure.BrokeringUserUPN)" | LogMe -display -progress
        $tests.BrokeringUserUPN = "NEUTRAL", $BrokeringUserUPN

        # Add the SiteName to the tests for the Syslog output
        $tests.SiteName = "NORMAL", $sitename

        #Fill $tests into array
        $BrokerConnectionLogResults.$MachineDNSName = $tests

        If ($CheckOutputSyslog) {
          # Set up the severity of the log entry based on the output of each test.
          $Severity = "Informational"
          If ($IsSeverityWarningLevel) { $Severity = "Warning" }
          If ($IsSeverityErrorLevel) { $Severity = "Error" }
          # Setup the PSCustomObject that will become the Data within the Structured Data
          $Data = [PSCustomObject]@{
            'BrokerConnectionFailure' = $MachineDNSName
          }
          $BrokerConnectionLogResults.$MachineDNSName.GetEnumerator() | ForEach-Object {
            $MyKey = $_.Key -replace " ", ""
            $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
          }
          $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
          If ($SyslogFileOnly) {
            Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$MachineDNSName" `
                                  -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                  -LogFilePath "$resultsSyslog" -FileOnly
          } Else {
            Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$MachineDNSName" `
                                  -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                  -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
          }
        }
      }
     " --- " | LogMe -display -progress
    }
  } Else {
    " --- " | LogMe -display -progress
  }
} #Close off $ShowBrokerConnectionFailuresTable

#== Hypervisor Connection Check =====================================================================================
"Check Hypervisor Connections #################################################################" | LogMe -display -progress

$HypervisorConnectionResults = @{}

$BrkrHvsCons = $null
If ($CitrixCloudCheck -ne 1) {
  $BrkrHvsCons = Get-Brokerhypervisorconnection -AdminAddress $AdminAddress
} Else {
  $BrkrHvsCons = Get-Brokerhypervisorconnection
}

If ($null -ne $BrkrHvsCons) {
  foreach ($BrkrHvsCon in $BrkrHvsCons) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}

    #Name of HypervisorConnection
    $HypervisorConnectionName = $BrkrHvsCon.Name
    "HypervisorConnection: $HypervisorConnectionName" | LogMe -display -progress
    $tests.HypervisorConnection = "NEUTRAL", $HypervisorConnectionName

    #State
    $BrkrHvsConState = $BrkrHvsCon.State
    "State: $($BrkrHvsConState)" | LogMe -display -progress
    If ($BrkrHvsConState -eq "On") {
      $tests.State = "NEUTRAL", $BrkrHvsCon.State
    } ElseIf ($BrkrHvsConState -eq "InMaintenanceMode") {
      $objMaintenance = $null
      Try {
        $objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $BrkrHvsCon.Name.ToUpper() } | Select-Object -First 1
      }
      Catch {
        # Avoid the error "The property 'TargetName' cannot be found on this object."
      }
      If ($null -ne $objMaintenance){$BrkrHvsConState = ("InMaintenanceMode, " + $objMaintenance.User)} Else {$BrkrHvsConState = "InMaintenanceMode"}
      # The Get-LogLowLevelOperation cmdlet will tell us who placed a Delivery Group into maintanance mode. However, the Get-Brokerhypervisorconnection
      # cmdlet will provide the MaintenanceReason, where the underlying reason, if manually entered, is stored in the MetadataMap property as part of a
      # Dictionary object.
      $MetadataMapDictionary = $BrkrHvsCon | ForEach-Object{ $_.MetadataMap }
      foreach ($key in $MetadataMapDictionary.Keys) {
        if ($key -eq "MaintenanceModeMessage") {
          $BrkrHvsConState = $BrkrHvsConState + ", " + $MetadataMapDictionary[$key]
        }
      }
      "MaintenanceModeInfo: $BrkrHvsConState" | LogMe -display -progress
      $tests.State = "WARNING", $BrkrHvsConState
      $IsSeverityWarningLevel = $True
    } Else {
      $tests.State = "ERROR", $BrkrHvsCon.State
      $IsSeverityErrorLevel = $True
    }

    "IsReady: $($BrkrHvsCon.IsReady)" | LogMe -display -progress
    If ($BrkrHvsCon.IsReady) {
      $tests.IsReady = "NEUTRAL", $BrkrHvsCon.IsReady
    } Else {
      $tests.IsReady = "ERROR", $BrkrHvsCon.IsReady
      $IsSeverityErrorLevel = $True
    }

    "MachineCount: $($BrkrHvsCon.MachineCount)" | LogMe -display -progress
    $tests.MachineCount = "NEUTRAL", $BrkrHvsCon.MachineCount

    Try {
      "FaultState: $($BrkrHvsCon.FaultState)" | LogMe -display -progress
      If ($BrkrHvsCon.FaultState -eq "None" -OR [string]::IsNullOrWhiteSpace($BrkrHvsCon.FaultState)) {
        $tests.FaultState = "NEUTRAL", $BrkrHvsCon.FaultState
      } Else {
        $tests.FaultState = "ERROR", $BrkrHvsCon.FaultState
        $IsSeverityErrorLevel = $True
      }
    }
    Catch {
      # Not all Broker Connections will have these properties
      "FaultState: Property does not exist on this object" | LogMe -display -progress
    }
    Try {
      "FaultReason: $($BrkrHvsCon.FaultReason)" | LogMe -display -progress
      $tests.FaultReason = "NEUTRAL", $BrkrHvsCon.FaultReason
    }
    Catch {
      # Not all Broker Connections will have these properties
      "FaultReason: Property does not exist on this object" | LogMe -display -progress
    }
    Try {
      "TimeFaultStateEntered: $($BrkrHvsCon.TimeFaultStateEntered)" | LogMe -display -progress
      $tests.TimeFaultStateEntered = "NEUTRAL", $BrkrHvsCon.TimeFaultStateEntered
    }
    Catch {
      # Not all Broker Connections will have these properties
      "TimeFaultStateEntered: Property does not exist on this object" | LogMe -display -progress
    }
    Try {
      "FaultStateDuration: $($BrkrHvsCon.FaultStateDuration)" | LogMe -display -progress
      $tests.FaultStateDuration = "NEUTRAL", $BrkrHvsCon.FaultStateDuration
    }
    Catch {
      # Not all Broker Connections will have these properties
      "FaultStateDuration: Property does not exist on this object" | LogMe -display -progress
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    #Fill $tests into array
    $HypervisorConnectionResults.$HypervisorConnectionName = $tests

    If ($CheckOutputSyslog) {
      # Set up the severity of the log entry based on the output of each test.
      $Severity = "Informational"
      If ($IsSeverityWarningLevel) { $Severity = "Warning" }
      If ($IsSeverityErrorLevel) { $Severity = "Error" }
      # Setup the PSCustomObject that will become the Data within the Structured Data
      $Data = [PSCustomObject]@{
        'HostingConnection' = $HypervisorConnectionName
      }
      $HypervisorConnectionResults.$HypervisorConnectionName.GetEnumerator() | ForEach-Object {
        $MyKey = $_.Key -replace " ", ""
        $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
      }
      $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
      If ($SyslogFileOnly) {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$HypervisorConnectionName" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -FileOnly
      } Else {
        Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$HypervisorConnectionName" `
                              -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                              -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
      }
    }
    " --- " | LogMe -display -progress
  }
} Else {
  " --- " | LogMe -display -progress
}

#==============================================================================================
# Start of VDI (single-session) Check
#==============================================================================================

"Check VDI (single-session) Desktops #########################################################" | LogMe -display -progress

" " | LogMe -display -progress

if($ShowDesktopTable -eq 1 ) {

  If (($ActualExcludedCatalogs | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Catalogs from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedCatalog in $ActualExcludedCatalogs) {
      "- $ActualExcludedCatalog" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedDeliveryGroups | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Delivery Groups from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedDeliveryGroup in $ActualExcludedDeliveryGroups) {
      "- $ActualExcludedDeliveryGroup" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedBrokerTags | Measure-Object).Count -gt 0) {
    "Excluding machines with the following Tags from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedBrokerTag in $ActualExcludedBrokerTags) {
      "- $ActualExcludedBrokerTag" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }

  $allResults = @{}

  If ($CitrixCloudCheck -ne 1) {
    $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Filter "SessionSupport -eq 'SingleSession'" | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)} | Sort-Object DNSName
  } Else {
    $machines = Get-BrokerMachine -MaxRecordCount $maxmachines -Filter "SessionSupport -eq 'SingleSession'" | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)} | Sort-Object DNSName
  }

  If (($machines | Measure-Object).Count -eq 0) {
    "There are no machines to process" | LogMe -display -progress
    " " | LogMe -display -progress
  }

  foreach($machine in $machines) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}
    $ErrorVDI = 0
  
    # Column Name of VDI
    $machineDNS = $machine | ForEach-Object{ $_.DNSName }
    If (![string]::IsNullOrWhiteSpace($machineDNS)) {
      "Machine: $machineDNS" | LogMe -display -progress
    } Else {
      "Machine: This is an invalid computer object" | LogMe -display -error
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityErrorLevel = $True
    }

    # Column IPv4Address
    if (!([string]::IsNullOrWhiteSpace($machineDNS))) {
      Try {
        $IPv4Address = ([System.Net.Dns]::GetHostAddresses($machineDNS) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) -join ", "
        "IPv4Address: $IPv4Address" | LogMe -display -progress
        $tests.IPv4Address = "NEUTRAL", $IPv4Address
      }
      Catch [System.Net.Sockets.SocketException] {
        "Failed to lookup host in DNS: $($_.Exception.Message)" | LogMe -display -warning
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityWarningLevel = $True
      }
      Catch {
        "An unexpected error occurred: $($_.Exception.Message)" | LogMe -display -warning
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityWarningLevel = $True
      }
    }

    # Column CatalogName
    $CatalogName = $machine | ForEach-Object{ $_.CatalogName }
    "Catalog: $CatalogName" | LogMe -display -progress
    $tests.CatalogName = "NEUTRAL", $CatalogName

    # Column DeliveryGroup
    $DeliveryGroup = $machine | ForEach-Object{ $_.DesktopGroupName }
    if (!([string]::IsNullOrWhiteSpace($DeliveryGroup))) {
      "DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
      $tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup
    } Else {
      "DeliveryGroup: This machine is not assigned to a Delivery Group" | LogMe -display -warning
      $tests.DeliveryGroup = "WARNING", $DeliveryGroup
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityWarningLevel = $True
    }

    # Column Powerstate
    $Powered = $machine | ForEach-Object{ $_.PowerState }
    if ($Powered -eq "On" -OR $Powered -eq "Unmanaged") {
      "PowerState: $Powered" | LogMe -display -progress
      $tests.PowerState = "SUCCESS", $Powered
    } ElseIf ($Powered -eq "Off") {
      "PowerState: $Powered" | LogMe -display -progress
      $tests.PowerState = "NEUTRAL", $Powered
    } else {
      "PowerState: $Powered" | LogMe -display -error
      $tests.PowerState = "ERROR", $Powered
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityErrorLevel = $True
    }

    # Column displaymode when a User has a Session
    $sessionUser = $machine | ForEach-Object{ $_.SessionUserName }

    if ($Powered -eq "On" -OR $Powered -eq "Unknown" -OR $Powered -eq "Unmanaged") {

      $IsPingable = Test-Ping -Target:$machineDNS -Timeout:200 -Count:3
      # Column Ping
      If ($IsPingable -eq "SUCCESS") {
        $tests.Ping = "SUCCESS", $IsPingable
      } Else {
        $tests.Ping = "NORMAL", $IsPingable
      }
      "Is Pingable: $IsPingable" | LogMe -display -progress

      $IsWinRMAccessible = IsWinRMAccessible -hostname:$machineDNS
      # Column WinRM
      If ($IsWinRMAccessible) {
        $tests.WinRM = "SUCCESS", $IsWinRMAccessible
      } Else {
        $tests.WinRM = "WARNING", $IsWinRMAccessible
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via WinRM: $IsWinRMAccessible" | LogMe -display -progress

      $IsWMIAccessible = IsWMIAccessible -hostname:$machineDNS -timeoutSeconds:20
      # Column WMI
      If ($IsWMIAccessible) {
        $tests.WMI= "SUCCESS", $IsWMIAccessible
      } Else {
        $tests.WMI = "WARNING", $IsWMIAccessible
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via WMI: $IsWMIAccessible" | LogMe -display -progress

      $IsUNCPathAccessible = (IsUNCPathAccessible -hostname:$machineDNS).Success
      # Column UNC
      If ($IsUNCPathAccessible) {
        $tests.UNC= "SUCCESS", $IsUNCPathAccessible
      } Else {
        $tests.UNC = "WARNING", $IsUNCPathAccessible
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via UNC: $IsUNCPathAccessible" | LogMe -display -progress

      if ($IsWinRMAccessible -OR $IsWMIAccessible -AND $machineDNS -NotLike "AAD-*") {

        If ($IsWinRMAccessible) {
          $UseWinRM = $True
        } Else {
          $UseWinRM = $False
        }

        #==============================================================================================
        #  Column Uptime
        #==============================================================================================
        $hostUptime = Get-UpTime -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $hostUptime) {
          if ($hostUptime.TimeSpan.days -gt ($maxUpTimeDays * 3)) {
            "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -error
            $tests.Uptime = "ERROR", $hostUptime.TimeSpan.days
            $ErrorVDI = $ErrorVDI + 1
            $IsSeverityErrorLevel = $True
          } elseif ($hostUptime.TimeSpan.days -gt $maxUpTimeDays) {
            "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -warning
            $tests.Uptime = "WARNING", $hostUptime.TimeSpan.days
            $ErrorVDI = $ErrorVDI + 1
            $IsSeverityWarningLevel = $True
          } else { 
            If ($hostUptime.TimeSpan.days -gt 0) {
              "Uptime: $($hostUptime.TimeSpan.days)"  | LogMe -display -progress
            } Else {
              "Uptime: $(ToHumanReadable($hostUptime.TimeSpan))" | LogMe -display -progress
            }
            $tests.Uptime = "SUCCESS", $hostUptime.TimeSpan.days
          }
        } else {
          "Unable to get host uptime" | LogMe -display -error
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityErrorLevel = $True
        }

        #==============================================================================================
        #  Get the CPU configuration
        #==============================================================================================
        $CpuConfigAndUsage = Get-CpuConfigAndUsage -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $CpuConfigAndUsage) {
          "CPU Configuration:" | LogMe -display -progress
          If ($CpuConfigAndUsage.LogicalProcessors -gt 1) {
            "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -progress
            $tests.LogicalProcessors = "NEUTRAL", $CpuConfigAndUsage.LogicalProcessors
          } ElseIf ($CpuConfigAndUsage.LogicalProcessors -eq 1) {
            "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -warning
            $tests.LogicalProcessors = "WARNING", $CpuConfigAndUsage.LogicalProcessors
            $ErrorVDI = $ErrorVDI + 1
            $IsSeverityWarningLevel = $True
          } Else {
            "- LogicalProcessors: Unable to detect." | LogMe -display -progress
          }
          If ($CpuConfigAndUsage.Sockets -gt 0) {
            "- Sockets: $($CpuConfigAndUsage.Sockets)" | LogMe -display -progress
            $tests.Sockets = "NEUTRAL", $CpuConfigAndUsage.Sockets
          } Else {
            "- Sockets: Unable to detect." | LogMe -display -progress
          }
          If ($CpuConfigAndUsage.CoresPerSocket -gt 0) {
            "- CoresPerSocket: $($CpuConfigAndUsage.CoresPerSocket)" | LogMe -display -progress
            $tests.CoresPerSocket = "NEUTRAL", $CpuConfigAndUsage.CoresPerSocket
          } Else {
            "- CoresPerSocket: Unable to detect." | LogMe -display -progress
          }
        } else {
          "Unable to get CPU configuration and usage" | LogMe -display -error
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityErrorLevel = $True
        }

        #==============================================================================================
        #  Get the Physical Memory size
        #==============================================================================================
        # Get the total Physical Memory
        $TotalPhysicalMemoryinGB = Get-TotalPhysicalMemory -hostname:$machineDNS -UseWinRM:$IsWinRMAccessible
        If ($TotalPhysicalMemoryinGB -ge 4) {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -progress
          $tests.TotalPhysicalMemoryinGB = "NEUTRAL", $TotalPhysicalMemoryinGB
        } ElseIf ($TotalPhysicalMemoryinGB -ge 2) {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -warning
          $tests.TotalPhysicalMemoryinGB = "WARNING", $TotalPhysicalMemoryinGB
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityWarningLevel = $True
        } Else {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -error
          $tests.TotalPhysicalMemoryinGB = "ERROR", $TotalPhysicalMemoryinGB
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityErrorLevel = $True
        }

        # Column OSBuild 
        $WinEntMultisession = $False
        $return = Get-OSVersion -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $return) {
          If ($return.Error -eq "Success") {
            $tests.OSCaption = "NEUTRAL", $return.Caption
            $tests.OSBuild = "NEUTRAL", $return.Version
            "OS Caption: $($return.Caption)" | LogMe -display -progress
            If ($return.Caption -like "*Enterprise*" -AND $return.Caption -like "*multi-session*") {
              $WinEntMultisession = $True
            }
            "OS Version: $($return.Version)" | LogMe -display -progress
          } Else {
            $tests.OSCaption = "ERROR", $return.Caption
            $tests.OSBuild = "ERROR", $return.Version
            "OS Test: $($return.Error)" | LogMe -display -error
            $ErrorVDI = $ErrorVDI + 1
            $IsSeverityErrorLevel = $True
          }
        } else {
          "Unable to get OS Version and Caption" | LogMe -display -error
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityErrorLevel = $True
        }

        ################ Start PVS SECTION ###############

        If ($IsUNCPathAccessible) {
          $PersonalityInfo = Get-PersonalityInfo -computername:$machineDNS
          $wcdrive = "D"
          If ($PersonalityInfo.IsPVS -OR $PersonalityInfo.IsMCS) {
            If ($PersonalityInfo.IsPVS ) { $wcdrive = $PvsWriteCacheDrive }
            If ($PersonalityInfo.IsMCS ) { $wcdrive = $MCSIOWriteCacheDrive }
            $WriteCacheDriveInfo = Get-WriteCacheDriveInfo -computername:$machineDNS -IsPVS:$PersonalityInfo.IsPVS -IsMCS:$PersonalityInfo.IsMCS -wcdrive:$wcdrive -UseWinRM:$UseWinRM

            If ($PersonalityInfo.IsPVS) {
              $tests.IsPVS = "SUCCESS", $PersonalityInfo.IsPVS
              $PersonalityInfo.Output_To_Log1 | LogMe -display -progress
              $tests.WriteCacheType = $PersonalityInfo.Output_For_HTML1, $PersonalityInfo.PVSCacheType
              $tests.PVSvDiskName = "NORMAL", $PersonalityInfo.PVSDiskName
              "Image Type: PVS"  | LogMe -display -progress
              "PVS vDisk Name: $($PersonalityInfo.PVSDiskName)"  | LogMe -display -progress
            }
            If ($PersonalityInfo.IsMCS) {
              $tests.IsMCS = "SUCCESS", $PersonalityInfo.IsMCS
              "Image Type: MCS"  | LogMe -display -progress
            }
            If ($PersonalityInfo.IsPVS -OR $PersonalityInfo.IsMCS) {
              $tests.DiskMode = "NEUTRAL", $PersonalityInfo.DiskMode
              "DiskMode: $($PersonalityInfo.DiskMode)"  | LogMe -display -progress
            } Else {
              "This is a standalone VDA"  | LogMe -display -progress
            }

            If ($PersonalityInfo.IsMCS -AND $WriteCacheDriveInfo.Output_To_Log2 -Like "*Failed to connect") {
              "It is assumed this is not using MCSIO as the script failed to connect to the $wcdrive drive." | LogMe -display -progress
              # If this is an Azure VM, Machine Creation Services (MCS) supports using Azure Ephemeral OS disk for
              # non-persistent VMs. Ephemeral disks should be fast IO because it uses temp storage on the local host.
              # MCSIO is not compatible with Azure Ephemeral Disks, and is therefore not an available configuration.
              $tests.WCdrivefreespace = "NEUTRAL", "N/A"
            } Else {
              $tests.WCdrivefreespace = $WriteCacheDriveInfo.Output_For_HTML2, $WriteCacheDriveInfo.WCdrivefreespace
              if (($WriteCacheDriveInfo.Output_To_Log2 -like "*normal*") -OR ($WriteCacheDriveInfo.Output_To_Log2 -like "*does not exist")) {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -progress
              }
              elseif ($WriteCacheDriveInfo.Output_To_Log2 -like "*low*") {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -warning
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
              else {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -error
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              }
            }

            If ($WriteCacheDriveInfo.vhdxSize_inMB -ne "N/A") {
              $CacheDiskMB = [long]($WriteCacheDriveInfo.vhdxSize_inMB)
              $CacheDiskGB = ($CacheDiskMB / 1024)
              "Write Cache file size: {0:n3} MB" -f($CacheDiskMB) | LogMe -display
              "Write Cache file size: {0:n3} GB" -f($CacheDiskMB / 1024) | LogMe -display
              "Write Cache max size: {0:n2} GB" -f($WriteCacheMaxSizeInGB) | LogMe -display
              if ($CacheDiskGB -lt ($WriteCacheMaxSizeInGB * 0.5)) {
                # If the cache file is less than 50% the max size, flag as all good
                "WriteCache file size is low" | LogMe
                 $tests.vhdxSize_inGB = "SUCCESS", "{0:n3} GB" -f($CacheDiskGB)
              }
              elseif ($CacheDiskGB -lt ($WriteCacheMaxSizeInGB * 0.8)) {
                # If the cache file is less than 80% the max size, flag as a warning
                "WriteCache file size moderate" | LogMe -display -warning
                $tests.vhdxSize_inGB = "WARNING", "{0:n3} GB" -f($CacheDiskGB)
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
              else {
                # Flag as an error when 80% or greater
                "WriteCache file size is high" | LogMe -display -error
                $tests.vhdxSize_inGB = "ERROR", "{0:n3} GB" -f($CacheDiskGB)
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              }
            } Else {
              $tests.vhdxSize_inGB = $WriteCacheDriveInfo.Output_For_HTML3, $WriteCacheDriveInfo.vhdxSize_inMB
            }
            if (($WriteCacheDriveInfo.Output_To_Log3 -like "*size is*") -OR ($WriteCacheDriveInfo.Output_To_Log3 -like "*N/A")) {
              $WriteCacheDriveInfo.Output_To_Log3 | LogMe -display -progress
            }
            else {
              $WriteCacheDriveInfo.Output_To_Log3 | LogMe -display -error
              $ErrorVDI = $ErrorVDI + 1
              $IsSeverityErrorLevel = $True
            }

          }
        } Else {
          # Cannot access UNC path so cannot test for the Personality.ini or MCSPersonality.ini
        }

        ################ End PVS SECTION ###############

        ################ Start Nvidia License Check SECTION ###############

        $NvidiaDriverFound = $False
        $NvidiaLicensedDriver = $False
        If ($IsWMIAccessible) {
          $return = Get-NvidiaDetails -computername:$machineDNS

          If ($return.Display_Driver_Ver -ne "N/A") {
            "Nvidia Driver Version: $($return.Display_Driver_Ver)" | LogMe -display -progress
            $tests.NvidiaDriverVer = "NEUTRAL", $return.Display_Driver_Ver
            $NvidiaDriverFound = $True
            If ($return.Licensable_Product -ne "N/A") {
              $NvidiaLicensedDriver = $True
            }
          }
        } Else {
          # WMI is not accessible
        }
        If ($NvidiaLicensedDriver) {
          If ($IsUNCPathAccessible) {
            $return = Check-NvidiaLicenseStatus -computername:$machineDNS
            $tests.nvidiaLicense = $return.Output_For_HTML, $return.Licensed
            if (($return.Output_To_Log -like "*successfully") -OR ($return.Output_To_Log -like "*does not exist") -OR ($return.Output_To_Log -like "*N/A")) {
              $return.Output_To_Log | LogMe -display -progress
            } else {
              $return.Output_To_Log | LogMe -display -error
              $ErrorVDI = $ErrorVDI + 1
              $IsSeverityErrorLevel = $True
            }
          } Else {
            # UNC Path is not accessible
          }
        } Else {
          If ($NvidiaDriverFound) {
            $tests.nvidiaLicense = "NEUTRAL", "N/A"
            "This host does not contain an Nvidia licensable product" | LogMe -display -progress
          }
        }

        ################ End Nvidia License Check SECTION ###############

        # Check services
        # The Get-Service command with -ComputerName parameter made use of DCOM and such functionality is
        # removed from PowerShell 7. So we use the Invoke-Command, which uses WinRM to run a ScriptBlock
        # instead.

        $ServicesChecked = $False
        If ($IsWinRMAccessible) {
          $ServicesChecked = $True
          Try {
            $services = Invoke-Command -ComputerName $machineDNS -ErrorAction Stop -ScriptBlock {Get-Service | where-object {$_.Name -eq 'Spooler' -OR $_.Name -eq 'cpsvc'}}

            if (($services | Where-Object {$_.Name -eq "Spooler"}).Status -Match "Running") {
              "SPOOLER service running..." | LogMe
              $tests.Spooler = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "SPOOLER service stopped" | LogMe -display -error
                $tests.Spooler = "ERROR","Error"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "SPOOLER service stopped" | LogMe -display -warning
                $tests.Spooler = "WARNING","Warning"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
            }

            if (($services | Where-Object {$_.Name -eq "cpsvc"}).Status -Match "Running") {
              "Citrix Print Manager service running..." | LogMe
              $tests.CitrixPrint = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "Citrix Print Manager service stopped" | LogMe -display -error
                $tests.CitrixPrint = "ERROR","Error"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "Citrix Print Manager service stopped" | LogMe -display -warning
                $tests.CitrixPrint = "WARNING","Warning"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
            }
          }
          Catch {
            #"Error returned while checking the services" | LogMe -error; return 101
            #$ErrorVDI = $ErrorVDI + 1
            #$IsSeverityErrorLevel = $True
          }
        } Else {
          # Cannot connect via WinRM
        }

        If ($IsWMIAccessible -AND $ServicesChecked -eq $False) {
          Try {
            $services = Get-WmiObject -ComputerName $machineDNS -Class Win32_Service -ErrorAction Stop | where-object {$_.Name -eq 'Spooler' -OR $_.Name -eq 'cpsvc'}

            if (($services | Where-Object {$_.Name -eq "Spooler"}).State -Match "Running") {
              "SPOOLER service running..." | LogMe
              $tests.Spooler = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "SPOOLER service stopped" | LogMe -display -error
                $tests.Spooler = "ERROR","Error"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "SPOOLER service stopped" | LogMe -display -warning
                $tests.Spooler = "WARNING","Warning"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
            }

            if (($services | Where-Object {$_.Name -eq "cpsvc"}).State -Match "Running") {
              "Citrix Print Manager service running..." | LogMe
              $tests.CitrixPrint = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "Citrix Print Manager service stopped" | LogMe -display -error
                $tests.CitrixPrint = "ERROR","Error"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "Citrix Print Manager service stopped" | LogMe -display -warning
                $tests.CitrixPrint = "WARNING","Warning"
                $ErrorVDI = $ErrorVDI + 1
                $IsSeverityWarningLevel = $True
              }
            }
          }
          Catch {
            #"Error returned while checking the services" | LogMe -error; return 101
            #$ErrorVDI = $ErrorVDI + 1
            #$IsSeverityErrorLevel = $True
          }
        } Else {
          # Cannot connect via WMI
        }

        If ($IsWinRMAccessible) {
          $ProfileStatus = Get-ProfileAndUserEnvironmentManagementServiceStatus -ComputerName:$machineDNS
          "Profile Management and User Environment Management Status:" | LogMe -display -progress
          $FSLogixEnabled = $False
          If ($ProfileStatus.FSLogixInstalled) {
            "- FSLogix Installed: $($ProfileStatus.FSLogixInstalled)" | LogMe -display -progress
            "- FSLogix ServiceRunning: $($ProfileStatus.FSLogixServiceRunning)" | LogMe -display -progress
            "- FSLogix ProfileEnabled: $($ProfileStatus.FSLogixProfileEnabled)" | LogMe -display -progress
            "- FSLogix ProfileType: $($ProfileStatus.FSLogixProfileType)" | LogMe -display -progress
            "- FSLogix ProfileTypeDescription: $($ProfileStatus.FSLogixProfileTypeDescription)" | LogMe -display -progress
            "- FSLogix OfficeEnabled: $($ProfileStatus.FSLogixOfficeEnabled)" | LogMe -display -progress
            "- FSLogix CCDLocations: $($ProfileStatus.FSLogixCCDLocations)" | LogMe -display -progress
            "- FSLogix VHDLocations: $($ProfileStatus.FSLogixVHDLocations)" | LogMe -display -progress
            "- FSLogix LogFilePath: $($ProfileStatus.FSLogixLogFilePath)" | LogMe -display -progress
            "- FSLogix RedirectionType: $($ProfileStatus.FSLogixRedirectionType)" | LogMe -display -progress
            If ($ProfileStatus.FSLogixServiceRunning -AND $ProfileStatus.FSLogixProfileEnabled -eq 1) {
              $FSLogixEnabled = $True
            }
          } Else {
            "- FSLogix is not installed" | LogMe -display -progress
          }
          If ($FSLogixEnabled) {
            $tests.FSLogixEnabled = "SUCCESS", $FSLogixEnabled
          } Else {
            $tests.FSLogixEnabled = "NORMAL", $FSLogixEnabled
          }
          $UPMEnabled = $False
          If ($ProfileStatus.UPMInstalled) {
            "- UPM Installed: $($ProfileStatus.UPMInstalled)" | LogMe -display -progress
            "- UPM ServiceRunning: $($ProfileStatus.UPMServiceRunning)" | LogMe -display -progress
            "- UPM ServiceActive: $($ProfileStatus.UPMServiceActive)" | LogMe -display -progress
            "- UPM PathToLogFile: $($ProfileStatus.UPMPathToLogFile)" | LogMe -display -progress
            "- UPM PathToUserStore: $($ProfileStatus.UPMPathToUserStore)" | LogMe -display -progress
            If ($ProfileStatus.UPMServiceRunning -AND $ProfileStatus.UPMServiceActive -eq 1) {
              $UPMEnabled = $True
            }
          } Else {
            "- UPM is not installed" | LogMe -display -progress
          }
          If ($UPMEnabled) {
            $tests.UPMEnabled = "SUCCESS", $UPMEnabled
          } Else {
            $tests.UPMEnabled = "NORMAL", $UPMEnabled
          }
          $WEMEnabled = $False
          If ($ProfileStatus.WEMInstalled) {
            "- WEM Installed: $($ProfileStatus.WEMInstalled)" | LogMe -display -progress
            "- WEM ServiceRunning: $($ProfileStatus.WEMServiceRunning)" | LogMe -display -progress
            "- WEM ServiceRunning: $($ProfileStatus.WEMServiceRunning)" | LogMe -display -progress
            "- WEM AgentRegistered: $($ProfileStatus.WEMAgentRegistered)" | LogMe -display -progress
            "- WEM AgentConfigurationSets: $($ProfileStatus.WEMAgentConfigurationSets)" | LogMe -display -progress
            "- WEM AgentCacheSyncMode: $($ProfileStatus.WEMAgentCacheSyncMode)" | LogMe -display -progress
            "- WEM AgentCachePath: $($ProfileStatus.WEMAgentCachePath)" | LogMe -display -progress
            If ($ProfileStatus.WEMServiceRunning -AND $ProfileStatus.WEMAgentRegistered) {
              $WEMEnabled = $True
            }
          } Else {
            "- WEM is not installed" | LogMe -display -progress
          }
          If ($WEMEnabled) {
            $tests.WEMEnabled = "SUCCESS", $WEMEnabled
          } Else {
            $tests.WEMEnabled = "NORMAL", $WEMEnabled
          }
        } Else {
          # Cannot connect via WinRM
        }

        If ($IsWMIAccessible) {
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

          } # Close off $ShowGraphicsMode
        }

        ## EDT MTU (set by default.ica or MTUDiscovery)
        If ($UseWinRM) {
          # Execute the "C:\Program Files (x86)\Citrix\HDX\bin\CtxSession.exe" in the remote session
          $EDTMTU = ""
          Try {
            $ctxsession = Invoke-Command -ComputerName $machineDNS -ErrorAction Stop -ScriptBlock { ctxsession -v }
            # ctxsession returns no data or null, such as for RDP sessions, so we check for that instead of failing the test.
            If (($ctxsession | Measure-Object).Count -eq 0) {
              $EDTMTU = ($ctxsession | findstr "EDT MTU:" | select -Last 1)
              if (!([string]::IsNullOrWhiteSpace($EDTMTU))) {
                $EDTMTU = ($ctxsession).split(":")[1].trimstart()
              } else {
                $EDTMTU = "Not set"
              }
            }
            $tests.EDT_MTU = "NEUTRAL", $EDTMTU
            "EDT MTU Size is set to $EDTMTU" | LogMe -display -progress
          }
          Catch {
            $tests.EDT_MTU = "ERROR", "Failed"
            "EDT MTU Size failed to return data" | LogMe -display -progress
          }
        } Else {
          "EDT MTU Size cannot be checked" | LogMe -display -progress
        }

        # Check CrowdStrike State
        If ($ShowCrowdStrikeTests -eq 1) {
          If ($IsWinRMAccessible) {
            $return = Get-CrowdStrikeServiceStatus -ComputerName:$machineDNS
            If ($null -ne $return) {
              If ($return.CSFalconInstalled -AND $return.CSAgentInstalled) {
                "CrowdStrike Installed: True" | LogMe -display -progress
                "- CrowdStrike Windows Sensor Version: $($return.InstalledVersion)" | LogMe -display -progress
                $tests.CSVersion = "NORMAL", $return.InstalledVersion
                "- CrowdStrike Company ID (CID): $($return.CID)" | LogMe -display -progress
                $tests.CSCID = "NORMAL", $return.CID
                "- CrowdStrike Sensor Grouping Tags: $($return.SensorGroupingTags)" | LogMe -display -progress
                $tests.CSGroupTags = "NORMAL", $return.SensorGroupingTags
                "- CrowdStrike VDI switch: $($return.VDI)" | LogMe -display -progress
                If ($return.CSFalconServiceRunning -AND $return.CSAgentServiceRunning -AND (![string]::IsNullOrEmpty($return.AID))) {
                  $tests.CSEnabled = "SUCCESS", $True
                  "- CrowdStrike Agent ID (AID): $($return.AID)" | LogMe -display -progress
                  $tests.CSAID = "NORMAL", $return.AID
                } Else {
                  $tests.CSEnabled = "WARNING", $False
                  If ([string]::IsNullOrEmpty($return.AID)) {
                    "- CrowdStrike Agent ID (AID) is missing" | LogMe -display -warning
                    $tests.CSAID = "NORMAL", "Missing"
                  } Else {
                    "- CrowdStrike is installed, but not running" | LogMe -display -warning
                  }
                  $IsSeverityWarningLevel = $True
                }
              } else {
                "CrowdStrike Installed: False" | LogMe -display -progress
              }
            } else {
              "Unable to get the CrowdStrike Service Status" | LogMe -display -error
              $IsSeverityErrorLevel = $True
            }
          }
        }

      }#If can connect via WinRM or WMI
      else {
        "WinRM or WMI connection not possible" | LogMe -display -error
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityErrorLevel = $True
      } # Closing else cannot connect via WinRM or WMI

    } # Close off $Powered -eq "On", "Unknown", or "Unmanaged"

    # Column RegistrationState
    $RegState = $machine | ForEach-Object{ $_.RegistrationState }
    if ($RegState -ne "Registered") {
      if ($Powered -eq "Off") {
        "RegistrationState: $RegState" | LogMe -display -progress
        $tests.RegState = "NEUTRAL", $RegState
      } else {
        "RegistrationState: $RegState" | LogMe -display -error
        $tests.RegState = "ERROR", $RegState
        $ErrorVDI = $ErrorVDI + 1
        $IsSeverityErrorLevel = $True
      }
    } else {
      "RegistrationState: $RegState" | LogMe -display -progress
      $tests.RegState = "SUCCESS", $RegState
    }

    # Column MaintenanceMode
    $MaintenanceMode = $machine | ForEach-Object{ $_.InMaintenanceMode }
    "MaintenanceMode: $MaintenanceMode" | LogMe -display -progress
    if ($MaintenanceMode) {
      $objMaintenance = $null
      Try {
        $objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $machine.MachineName.ToUpper() } | Select-Object -First 1
      }
      Catch {
        # Avoid the error "The property 'TargetName' cannot be found on this object."
      }
      If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
      # The Get-LogLowLevelOperation cmdlet will tell us who placed a machine into maintanance mode. However, the Get-BrokerMachine cmdlet
      # will provide the MaintenanceReason, where the underlying reason, if manually entered, is stored in the MetadataMap property as part
      # of a Dictionary object.
      $MetadataMapDictionary = $machine | ForEach-Object{ $_.MetadataMap }
      foreach ($key in $MetadataMapDictionary.Keys) {
        if ($key -eq "MaintenanceModeMessage") {
          $MaintenanceModeOn = $MaintenanceModeOn + ", " + $MetadataMapDictionary[$key]
        }
      }
      "MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
      $tests.MaintMode = "WARNING", $MaintenanceModeOn
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityWarningLevel = $True
    }
    else { $tests.MaintMode = "SUCCESS", "OFF" }
  
    # Column HostedOn 
    $HostedOn = $machine | ForEach-Object{ $_.HostingServerName }
    "HostedOn: $HostedOn" | LogMe -display -progress
    $tests.HostedOn = "NEUTRAL", $HostedOn

    # Column VDAVersion AgentVersion
    $VDAVersion = $machine | ForEach-Object{ $_.AgentVersion }
    $Found = $False
    If ($SupportedVDAVersions.Count -gt 0) {
      If (!([String]::IsNullOrEmpty($SupportedVDAVersions[0]))) {
        ForEach ($SupportedVDAVersion in $SupportedVDAVersions) {
          If ($VDAVersion -Like $SupportedVDAVersion) {
            $Found = $True
            break
          }
        }
        If ($SupportedVDAVersions -contains $VDAVersion) {
          $Found = $True
        }
      }
    } Else {
      $Found = $True
    }
    If ($Found) {
      "VDAVersion: $VDAVersion" | LogMe -display -progress
      $tests.VDAVersion = "NEUTRAL", $VDAVersion
    } Else {
      "VDAVersion: $VDAVersion" | LogMe -display -warning
      $tests.VDAVersion = "WARNING", $VDAVersion
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityWarningLevel = $True
    }

    # Column AssociatedUserNames
    $AssociatedUserNames = $machine | ForEach-Object{ $_.AssociatedUserNames }
    "Assigned to $AssociatedUserNames" | LogMe -display -progress
    $tests.AssociatedUserNames = "NEUTRAL", $AssociatedUserNames

    # Column Tags 
    $Tags = $machine | ForEach-Object{ $_.Tags }
    "Tags: $Tags" | LogMe -display -progress
    $tests.Tags = "NEUTRAL", $Tags

    # Column MCSImageOutOfDate
    $ProvisioningType = $machine | ForEach-Object{ $_.ProvisioningType }
    # The Get-BrokerMachine cmdlet has a ProvisioningType property, but we can also get this from the Machine Catalogs we have already collected.
    # The Machine Catalogs should exist in the $Catalogs variable, so test this first before collecting them again using the Get-BrokerCatalog cmdlet.
    # Have left the next 13 lines of code in for future reference.
    # $ProvisioningType = ""
    # $GetCatalogs = $True
    # If (($Catalogs | Measure-Object).Count -gt 0) {
    #   $ProvisioningType = ($Catalogs | Where-Object {$_.Name -eq $CatalogName} | Select-Object -ExpandProperty ProvisioningType)
    #   $GetCatalogs = $False
    # }
    # If ($GetCatalogs) {
    #   If ($CitrixCloudCheck -ne 1) { 
    #     $ProvisioningType = (Get-BrokerCatalog -AdminAddress $AdminAddress -Name $CatalogName | Select-Object -ExpandProperty ProvisioningType)
    #   } Else {
    #     $ProvisioningType = (Get-BrokerCatalog -Name $CatalogName | Select-Object -ExpandProperty ProvisioningType)
    #   }
    # }
    If ($ProvisioningType -eq "MCS") {
      $MCSImageOutOfDate = $machine | ForEach-Object{ $_.ImageOutOfDate }
      if ($MCSImageOutOfDate -eq $true) {
        if ($Powered -eq "Off") {
          "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -progress
          $tests.MCSImageOutOfDate = "NEUTRAL", $MCSImageOutOfDate
        } else {
          "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -error
          $tests.MCSImageOutOfDate = "ERROR", $MCSImageOutOfDate
          $ErrorVDI = $ErrorVDI + 1
          $IsSeverityErrorLevel = $True
        }
      } else {
        "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -progress
        if ($Powered -eq "Off") {
          $tests.MCSImageOutOfDate = "NEUTRAL", $MCSImageOutOfDate
        } else {
          $tests.MCSImageOutOfDate = "SUCCESS", $MCSImageOutOfDate
        }
      }
    }

    # Column LastConnectionTime
    $yellow =((Get-Date).AddDays(-30).ToString('yyyy-MM-dd HH:mm:s'))
    $red =((Get-Date).AddDays(-90).ToString('yyyy-MM-dd HH:mm:s'))
    $machineLastConnectionTime = $machine | ForEach-Object{ $_.LastConnectionTime }
    if ([string]::IsNullOrWhiteSpace($machineLastConnectionTime))
    {
      $tests.LastConnectionTime = "NEUTRAL", "NO DATA"
    }
    elseif ($machineLastConnectionTime -lt $red)
    {
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -ERROR
      $tests.LastConnectionTime = "ERROR", $machineLastConnectionTime
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityErrorLevel = $True
    } 	
    elseif ($machineLastConnectionTime -lt $yellow)
    {
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -WARNING
      $tests.LastConnectionTime = "WARNING", $machineLastConnectionTime
      $ErrorVDI = $ErrorVDI + 1
      $IsSeverityWarningLevel = $True
    }
    else 
    {
      $tests.LastConnectionTime = "SUCCESS", $machineLastConnectionTime
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -progress
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    # Fill $tests into array if error occured OR $ShowOnlyErrorVDI = 0
    # Check if error exists on this vdi
    if ($ShowOnlyErrorVDI -eq 0 ) { $allResults.$machineDNS = $tests }
    else {
      if ($ErrorVDI -gt 0) { $allResults.$machineDNS = $tests }
      else { "$machineDNS is ok, no output into HTML-File" | LogMe -display -progress }
    }

    If ($tests.Count -gt 0) {
      If ($CheckOutputSyslog) {
        # Set up the severity of the log entry based on the output of each test.
        $Severity = "Informational"
        If ($IsSeverityWarningLevel) { $Severity = "Warning" }
        If ($IsSeverityErrorLevel) { $Severity = "Error" }
        # Setup the PSCustomObject that will become the Data within the Structured Data
        $Data = [PSCustomObject]@{
          'SingleSessionHost' = $machineDNS
        }
        $allResults.$machineDNS.GetEnumerator() | ForEach-Object {
          $MyKey = $_.Key -replace " ", ""
          $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
        }
        $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
        If ($SyslogFileOnly) {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -FileOnly
        } Else {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
        }
      }
    }
    " --- " | LogMe -display -progress

  } # Close off foreach $machine

} # Close off $ShowDesktopTable

else { "Desktop Check skipped because ShowDesktopTable = 0 " | LogMe -display -progress }

"####################### Check END  ##########################################################" | LogMe -display -progress

#==============================================================================================
# End of VDI (single-session) Check
#==============================================================================================

#==============================================================================================
# Start of XenApp/RDSH (multi-session) Check
#==============================================================================================

"Check XenApp/RDSH (multi-session) Servers ###################################################" | LogMe -display -progress

" " | LogMe -display -progress
  
# Check XenApp only if $ShowXenAppTable is 1

if($ShowXenAppTable -eq 1 ) {

  If (($ActualExcludedCatalogs | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Catalogs from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedCatalog in $ActualExcludedCatalogs) {
      "- $ActualExcludedCatalog" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedDeliveryGroups | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Delivery Groups from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedDeliveryGroup in $ActualExcludedDeliveryGroups) {
      "- $ActualExcludedDeliveryGroup" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedBrokerTags | Measure-Object).Count -gt 0) {
    "Excluding machines with the following Tags from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedBrokerTag in $ActualExcludedBrokerTags) {
      "- $ActualExcludedBrokerTag" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }

  $allXenAppResults = @{}

  If ($CitrixCloudCheck -ne 1) {
    $XAmachines = Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress -Filter "SessionSupport -eq 'MultiSession'"  | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)} | Sort-Object DNSName
  } Else {
    $XAmachines = Get-BrokerMachine -MaxRecordCount $maxmachines -Filter "SessionSupport -eq 'MultiSession'"  | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0 -and ($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)} | Sort-Object DNSName
  }

  If (($XAmachines | Measure-Object).Count -eq 0) {
    "There are no machines to process" | LogMe -display -progress
    " " | LogMe -display -progress
  }

  foreach ($XAmachine in $XAmachines) {
    $IsSeverityErrorLevel = $False
    $IsSeverityWarningLevel = $False
    $tests = @{}
    $ErrorXA = 0

    # Column Name of Machine
    $machineDNS = $XAmachine | ForEach-Object{ $_.DNSName }
    If (![string]::IsNullOrWhiteSpace($machineDNS)) {
      "Machine: $machineDNS" | LogMe -display -progress
    } Else {
      "Machine: This is an invalid computer object" | LogMe -display -error
      $ErrorXA = $ErrorXA + 1
      $IsSeverityErrorLevel = $True
    }

    # Column IPv4Address
    if (!([string]::IsNullOrWhiteSpace($machineDNS))) {
      Try {
        $IPv4Address = ([System.Net.Dns]::GetHostAddresses($machineDNS) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) -join ", "
        "IPv4Address: $IPv4Address" | LogMe -display -progress
        $tests.IPv4Address = "NEUTRAL", $IPv4Address
      }
      Catch [System.Net.Sockets.SocketException] {
        "Failed to lookup host in DNS: $($_.Exception.Message)" | LogMe -display -warning
        $ErrorXA = $ErrorXA + 1
        $IsSeverityWarningLevel = $True
      }
      Catch {
        "An unexpected error occurred: $($_.Exception.Message)" | LogMe -display -warning
        $ErrorXA = $ErrorXA + 1
        $IsSeverityWarningLevel = $True
      }
    }

    # Column CatalogNameName
    $CatalogName = $XAmachine | ForEach-Object{ $_.CatalogName }
    "Catalog: $CatalogName" | LogMe -display -progress
    $tests.CatalogName = "NEUTRAL", $CatalogName

    # Column DeliveryGroup
    $DeliveryGroup = $XAmachine | ForEach-Object{ $_.DesktopGroupName }
    if (!([string]::IsNullOrWhiteSpace($DeliveryGroup))) {
      "DeliveryGroup: $DeliveryGroup" | LogMe -display -progress
      $tests.DeliveryGroup = "NEUTRAL", $DeliveryGroup
    } Else {
      "DeliveryGroup: This machine is not assigned to a Delivery Group" | LogMe -display -warning
      $tests.DeliveryGroup = "WARNING", $DeliveryGroup
      $ErrorXA = $ErrorXA + 1
      $IsSeverityWarningLevel = $True
    }

    # Column Powerstate
    $Powered = $XAmachine | ForEach-Object{ $_.PowerState }
    if ($Powered -eq "On" -OR $Powered -eq "Unmanaged") {
      "PowerState: $Powered" | LogMe -display -progress
      $tests.PowerState = "SUCCESS", $Powered
    } ElseIf ($Powered -eq "Off") {
      "PowerState: $Powered" | LogMe -display -progress
      $tests.PowerState = "NEUTRAL", $Powered
    } else {
      "PowerState: $Powered" | LogMe -display -error
      $tests.PowerState = "ERROR", $Powered
      $ErrorXA = $ErrorXA + 1
      $IsSeverityErrorLevel = $True
    }

    if ($Powered -eq "On" -OR $Powered -eq "Unknown" -OR $Powered -eq "Unmanaged") {

      $IsPingable = Test-Ping -Target:$machineDNS -Timeout:100 -Count:3
      # Column Ping
      If ($IsPingable -eq "SUCCESS") {
        $tests.Ping = "SUCCESS", $IsPingable
      } Else {
        $tests.Ping = "NORMAL", $IsPingable
      }
      "Is Pingable: $IsPingable" | LogMe -display -progress

      $IsWinRMAccessible = IsWinRMAccessible -hostname:$machineDNS
      # Column WinRM
      If ($IsWinRMAccessible) {
        $tests.WinRM = "SUCCESS", $IsWinRMAccessible
      } Else {
        $tests.WinRM = "WARNING", $IsWinRMAccessible
        $ErrorXA = $ErrorXA + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via WinRM: $IsWinRMAccessible" | LogMe -display -progress

      $IsWMIAccessible = IsWMIAccessible -hostname:$machineDNS -timeoutSeconds:20
      # Column WMI
      If ($IsWMIAccessible) {
        $tests.WMI= "SUCCESS", $IsWMIAccessible
      } Else {
        $tests.WMI = "WARNING", $IsWMIAccessible
        $ErrorXA = $ErrorXA + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via WMI: $IsWMIAccessible" | LogMe -display -progress

      $IsUNCPathAccessible = (IsUNCPathAccessible -hostname:$machineDNS).Success
      # Column UNC
      If ($IsUNCPathAccessible) {
        $tests.UNC= "SUCCESS", $IsUNCPathAccessible
      } Else {
        $tests.UNC = "WARNING", $IsUNCPathAccessible
        $ErrorXA = $ErrorXA + 1
        $IsSeverityWarningLevel = $True
      }
      "Can connect via UNC: $IsUNCPathAccessible" | LogMe -display -progress

      if ($IsWinRMAccessible -OR $IsWMIAccessible) {

        $UseWinRM = $False
        If ($IsWinRMAccessible) {
          $UseWinRM = $True
        }

        #==============================================================================================
        #  Column Uptime
        #==============================================================================================
        $hostUptime = Get-UpTime -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $hostUptime) {
          if ($hostUptime.TimeSpan.days -gt ($maxUpTimeDays * 3)) {
            "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -error
            $tests.Uptime = "ERROR", $hostUptime.TimeSpan.days
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          } elseif ($hostUptime.TimeSpan.days -gt $maxUpTimeDays) {
            "reboot warning, last reboot: {0:D}" -f $hostUptime.LBTime | LogMe -display -warning
            $tests.Uptime = "WARNING", $hostUptime.TimeSpan.days
            $ErrorXA = $ErrorXA + 1
            $IsSeverityWarningLevel = $True
          } else {
            If ($hostUptime.TimeSpan.days -gt 0) {
              "Uptime: $($hostUptime.TimeSpan.days)"  | LogMe -display -progress
            } Else {
              "Uptime: $(ToHumanReadable($hostUptime.TimeSpan))" | LogMe -display -progress
            }
            $tests.Uptime = "SUCCESS", $hostUptime.TimeSpan.days
          }
        } else {
          "Unable to get host uptime" | LogMe -display -error
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }

        #==============================================================================================
        #  Get the CPU configuration and check the AvgCPU value for 5 seconds
        #==============================================================================================

        $CpuConfigAndUsage = Get-CpuConfigAndUsage -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $CpuConfigAndUsage) {
          $XAAvgCPUval = $CpuConfigAndUsage.CpuUsage
          if( [int] $XAAvgCPUval -lt 75) { "CPU usage is normal [ $XAAvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", "$XAAvgCPUval %" }
          elseif([int] $XAAvgCPUval -lt 85) { "CPU usage is medium [ $XAAvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", "$XAAvgCPUval %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityWarningLevel = $True }   	
          elseif([int] $XAAvgCPUval -lt 95) { "CPU usage is high [ $XAAvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$XAAvgCPUval %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
          elseif([int] $XAAvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
          else { "CPU usage is Critical [ $XAAvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", "$XAAvgCPUval %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
          $XAAvgCPUval = 0
          "CPU Configuration:" | LogMe -display -progress
          If ($CpuConfigAndUsage.LogicalProcessors -gt 1) {
            "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -progress
            $tests.LogicalProcessors = "NEUTRAL", $CpuConfigAndUsage.LogicalProcessors
          } ElseIf ($CpuConfigAndUsage.LogicalProcessors -eq 1) {
            "- LogicalProcessors: $($CpuConfigAndUsage.LogicalProcessors)" | LogMe -display -warning
            $tests.LogicalProcessors = "WARNING", $CpuConfigAndUsage.LogicalProcessors
            $ErrorXA = $ErrorXA + 1
            $IsSeverityWarningLevel = $True
          } Else {
            "- LogicalProcessors: Unable to detect." | LogMe -display -progress
          }
          If ($CpuConfigAndUsage.Sockets -gt 0) {
            "- Sockets: $($CpuConfigAndUsage.Sockets)" | LogMe -display -progress
            $tests.Sockets = "NEUTRAL", $CpuConfigAndUsage.Sockets
          } Else {
            "- Sockets: Unable to detect." | LogMe -display -progress
          }
          If ($CpuConfigAndUsage.CoresPerSocket -gt 0) {
            "- CoresPerSocket: $($CpuConfigAndUsage.CoresPerSocket)" | LogMe -display -progress
            $tests.CoresPerSocket = "NEUTRAL", $CpuConfigAndUsage.CoresPerSocket
          } Else {
            "- CoresPerSocket: Unable to detect." | LogMe -display -progress
          }
        } else {
          "Unable to get CPU configuration and usage" | LogMe -display -error
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }

        #==============================================================================================
        #  Get the Physical Memory size and usage 
        #==============================================================================================
        [int] $XAUsedMemory = CheckMemoryUsage -hostname:$machineDNS -UseWinRM:$UseWinRM

        If ($null -ne $XAUsedMemory) {
          if( [int] $XAUsedMemory -lt 75) { "Memory usage is normal [ $XAUsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", "$XAUsedMemory %" }
          elseif( [int] $XAUsedMemory -lt 85) { "Memory usage is medium [ $XAUsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", "$XAUsedMemory %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityWarningLevel = $True }   	
          elseif( [int] $XAUsedMemory -lt 95) { "Memory usage is high [ $XAUsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$XAUsedMemory %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
          elseif( [int] $XAUsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
          else { "Memory usage is Critical [ $XAUsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", "$XAUsedMemory %" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }   
          $XAUsedMemory = 0  
        } else {
          "Unable to get Memory usage" | LogMe -display -error
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }

        # Get the total Physical Memory
        $TotalPhysicalMemoryinGB = Get-TotalPhysicalMemory -hostname:$machineDNS -UseWinRM:$IsWinRMAccessible
        If ($TotalPhysicalMemoryinGB -ge 4) {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -progress
          $tests.TotalPhysicalMemoryinGB = "NEUTRAL", $TotalPhysicalMemoryinGB
        } ElseIf ($TotalPhysicalMemoryinGB -ge 2) {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -warning
          $tests.TotalPhysicalMemoryinGB = "WARNING", $TotalPhysicalMemoryinGB
          $ErrorXA = $ErrorXA + 1
          $IsSeverityWarningLevel = $True
        } Else {
          "Total Physical Memory: $($TotalPhysicalMemoryinGB) GB" | LogMe -display -error
          $tests.TotalPhysicalMemoryinGB = "ERROR", $TotalPhysicalMemoryinGB
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }

        #==============================================================================================
        #  Check Disk Usage
        #==============================================================================================
        foreach ($disk in $diskLettersWorkers)
        {
          "Checking free space on $($disk):" | LogMe -display
          $HardDisk = CheckHardDiskUsage -hostname:$machineDNS -deviceID:"$($disk):" -UseWinRM:$UseWinRM
          if ($null -ne $HardDisk) {	
            $XAPercentageDS = $HardDisk.PercentageDS
            $frSpace = $HardDisk.frSpace
            If ( [int] $XAPercentageDS -gt 15) { "Disk Free is normal [ $XAPercentageDS % ]" | LogMe -display; $tests."$($disk)Freespace" = "SUCCESS", "$frSpace GB" } 
            ElseIf ([int] $XAPercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests.CFreespace = "ERROR", "Err" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
            ElseIf ([int] $XAPercentageDS -lt 5) { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True } 
            ElseIf ([int] $XAPercentageDS -lt 15) { "Disk Free is Low [ $XAPercentageDS % ]" | LogMe -warning; $tests."$($disk)Freespace" = "WARNING", "$frSpace GB" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityWarningLevel = $True }     
            Else { "Disk Free is Critical [ $XAPercentageDS % ]" | LogMe -error; $tests."$($disk)Freespace" = "ERROR", "$frSpace GB" ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
            $XAPercentageDS = 0
            $frSpace = 0
            $HardDisk = $null
          } else {
            "Unable to get Hard Disk usage" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          }
        }
 
        # Column OSBuild 
        $WinEntMultisession = $False
        $return = Get-OSVersion -hostname:$machineDNS -UseWinRM:$UseWinRM
        If ($null -ne $return) {
          If ($return.Error -eq "Success") {
            $tests.OSCaption = "NEUTRAL", $return.Caption
            $tests.OSBuild = "NEUTRAL", $return.Version
            "OS Caption: $($return.Caption)" | LogMe -display -progress
            If ($return.Caption -like "*Enterprise*" -AND $return.Caption -like "*multi-session*") {
              $WinEntMultisession = $True
            }
            "OS Version: $($return.Version)" | LogMe -display -progress
          } Else {
            $tests.OSCaption = "ERROR", $return.Caption
            $tests.OSBuild = "ERROR", $return.Version
            "OS Test: $($return.Error)" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          }
        } else {
          "Unable to get OS Version and Caption" | LogMe -display -error
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }

        ################ Start PVS SECTION ###############

        If ($IsUNCPathAccessible) {
          $PersonalityInfo = Get-PersonalityInfo -computername:$machineDNS
          $wcdrive = "D"
          If ($PersonalityInfo.IsPVS -OR $PersonalityInfo.IsMCS) {
            If ($PersonalityInfo.IsPVS ) { $wcdrive = $PvsWriteCacheDrive }
            If ($PersonalityInfo.IsMCS ) { $wcdrive = $MCSIOWriteCacheDrive }
            $WriteCacheDriveInfo = Get-WriteCacheDriveInfo -computername:$machineDNS -IsPVS:$PersonalityInfo.IsPVS -IsMCS:$PersonalityInfo.IsMCS -wcdrive:$wcdrive -UseWinRM:$UseWinRM

            If ($PersonalityInfo.IsPVS) {
              $tests.IsPVS = "SUCCESS", $PersonalityInfo.IsPVS
              $PersonalityInfo.Output_To_Log1 | LogMe -display -progress
              $tests.WriteCacheType = $PersonalityInfo.Output_For_HTML1, $PersonalityInfo.PVSCacheType
              $tests.PVSvDiskName = "NORMAL", $PersonalityInfo.PVSDiskName
              "Image Type: PVS"  | LogMe -display -progress
              "PVS vDisk Name: $($PersonalityInfo.PVSDiskName)"  | LogMe -display -progress
            }
            If ($PersonalityInfo.IsMCS) {
              $tests.IsMCS = "SUCCESS", $PersonalityInfo.IsMCS
              "Image Type: MCS"  | LogMe -display -progress
            }
            If ($PersonalityInfo.IsPVS -OR $PersonalityInfo.IsMCS) {
              $tests.DiskMode = "NEUTRAL", $PersonalityInfo.DiskMode
              "DiskMode: $($PersonalityInfo.DiskMode)"  | LogMe -display -progress
            } Else {
              "This is a standalone VDA"  | LogMe -display -progress
            }

            If ($PersonalityInfo.IsMCS -AND $WriteCacheDriveInfo.Output_To_Log2 -Like "*Failed to connect") {
              "It is assumed this is not using MCSIO" | LogMe -display -progress
              # If this is an Azure VM, Machine Creation Services (MCS) supports using Azure Ephemeral OS disk for
              # non-persistent VMs. Ephemeral disks should be fast IO because it uses temp storage on the local host.
              # MCSIO is not compatible with Azure Ephemeral Disks, and is therefore not an available configuration.
              $tests.WCdrivefreespace = "NEUTRAL", "N/A"
            } Else {
              $tests.WCdrivefreespace = $WriteCacheDriveInfo.Output_For_HTML2, $WriteCacheDriveInfo.WCdrivefreespace
              if (($WriteCacheDriveInfo.Output_To_Log2 -like "*normal*") -OR ($WriteCacheDriveInfo.Output_To_Log2 -like "*does not exist")) {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -progress
              }
              elseif ($WriteCacheDriveInfo.Output_To_Log2 -like "*low*") {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -warning
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
              else {
                $WriteCacheDriveInfo.Output_To_Log2 | LogMe -display -error
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              }
            }

            If ($WriteCacheDriveInfo.vhdxSize_inMB -ne "N/A") {
              $CacheDiskMB = [long]($WriteCacheDriveInfo.vhdxSize_inMB)
              $CacheDiskGB = ($CacheDiskMB / 1024)
              "Write Cache file size: {0:n3} MB" -f($CacheDiskMB) | LogMe -display
              "Write Cache file size: {0:n3} GB" -f($CacheDiskMB / 1024) | LogMe -display
              "Write Cache max size: {0:n2} GB" -f($WriteCacheMaxSizeInGB) | LogMe -display
              if ($CacheDiskGB -lt ($WriteCacheMaxSizeInGB * 0.5)) {
                # If the cache file is less than 50% the max size, flag as all good
                "WriteCache file size is low" | LogMe
                 $tests.vhdxSize_inGB = "SUCCESS", ("{0:n3} GB" -f($CacheDiskGB))
              }
              elseif ($CacheDiskGB -lt ($WriteCacheMaxSizeInGB * 0.8)) {
                # If the cache file is less than 80% the max size, flag as a warning
                "WriteCache file size moderate" | LogMe -display -warning
                $tests.vhdxSize_inGB = "WARNING", ("{0:n3} GB" -f($CacheDiskGB))
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
              else {
                # Flag as an error when 80% or greater
                "WriteCache file size is high" | LogMe -display -error
                $tests.vhdxSize_inGB = "ERROR", ("{0:n3} GB" -f($CacheDiskGB))
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              }
            } Else {
              $tests.vhdxSize_inGB = $WriteCacheDriveInfo.Output_For_HTML3, $WriteCacheDriveInfo.vhdxSize_inMB
            }
            if (($WriteCacheDriveInfo.Output_To_Log3 -like "*size is*") -OR ($WriteCacheDriveInfo.Output_To_Log3 -like "*N/A")) {
              $WriteCacheDriveInfo.Output_To_Log3 | LogMe -display -progress
            }
            else {
              $WriteCacheDriveInfo.Output_To_Log3 | LogMe -display -error
              $ErrorXA = $ErrorXA + 1
              $IsSeverityErrorLevel = $True
            }

          }
        } Else {
          # Cannot access UNC path so cannot test for the Personality.ini or MCSPersonality.ini
        }

        ################ End PVS SECTION ###############

        ################ Start Nvidia License Check SECTION ###############

        $NvidiaDriverFound = $False
        $NvidiaLicensedDriver = $False
        If ($IsWMIAccessible) {
          $return = Get-NvidiaDetails -computername:$machineDNS

          If ($return.Display_Driver_Ver -ne "N/A") {
            "Nvidia Driver Version: $($return.Display_Driver_Ver)" | LogMe -display -progress
            $tests.NvidiaDriverVer = "NEUTRAL", $return.Display_Driver_Ver
            $NvidiaDriverFound = $True
            If ($return.Licensable_Product -ne "N/A") {
              $NvidiaLicensedDriver = $True
            }
          }
        } Else {
          # WMI is not accessible
        }
        If ($NvidiaLicensedDriver) {
          If ($IsUNCPathAccessible) {
            $return = Check-NvidiaLicenseStatus -computername:$machineDNS
            $tests.nvidiaLicense = $return.Output_For_HTML, $return.Licensed
            if (($return.Output_To_Log -like "*successfully") -OR ($return.Output_To_Log -like "*does not exist") -OR ($return.Output_To_Log -like "*N/A")) {
              $return.Output_To_Log | LogMe -display -progress
            } else {
              $return.Output_To_Log | LogMe -display -error
              $ErrorXA = $ErrorXA + 1
              $IsSeverityErrorLevel = $True
            }
          } Else {
            # UNC Path is not accessible
          }
        } Else {
          If ($NvidiaDriverFound) {
            $tests.nvidiaLicense = "NEUTRAL", "N/A"
            "This host does not contain an Nvidia licensable product" | LogMe -display -progress
          }
        }

        ################ End Nvidia License Check SECTION ###############

        ################ Start RDS Licensing Details Check SECTION ###############

        If ($WinEntMultisession -eq $False) {
          $return = Get-RDLicenseGracePeriodEventErrorsSinceBoot -ComputerName:$machineDNS -UseWinRM:$UseWinRM -LastBootTime:$hostUptime.LBTime
          $RDSGracePeriodExpired = $False
          If ($return.Found -AND $Return.Count -gt 0) {
            $RDSGracePeriodExpired = $True
            $tests.RDSGracePeriodExpired = "ERROR", $return.Found
            "RDSGracePeriodExpired: $($return.Found)" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          } Else {
            $tests.RDSGracePeriodExpired = "NEUTRAL", $return.Found
            "RDSGracePeriodExpired: $($return.Found)" | LogMe -display -progress
          }
          $return = Get-RDSLicensingDetails -computername:$machineDNS -UseWinRM:$UseWinRM
          If ($RDSGracePeriodExpired -eq $False) {
            $tests.RDSGracePeriod = $return.Output_For_HTML, $return.GracePeriod
            if (($return.Output_To_Log -like "*Good*") -OR ($return.Output_To_Log -like "*N/A")) {
              $return.Output_To_Log | LogMe -display -progress
            }
            elseif ($return.Output_To_Log -like "*Warning*") {
              $return.Output_To_Log | LogMe -display -warning
              $ErrorXA = $ErrorXA + 1
              $IsSeverityWarningLevel = $True
            }
            else {
              $return.Output_To_Log | LogMe -display -error
              $ErrorXA = $ErrorXA + 1
              $IsSeverityErrorLevel = $True
            }
          } Else {
            $tests.RDSGracePeriod = "ERROR", $return.GracePeriod
            "RDSGracePeriod: Expired [ $($return.GracePeriod) ]" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          }
          If ($return.TerminalServerMode -eq "AppServer") {
            $tests.TerminalServerMode = "SUCCESS",$return.TerminalServerMode
            "TerminalServerMode: $($return.TerminalServerMode)" | LogMe -display -progress
          } Else {
            $tests.TerminalServerMode = "ERROR",$return.TerminalServerMode
            "TerminalServerMode: $($return.TerminalServerMode)" | LogMe -display -Error
            If ($return.TerminalServerMode -eq "RemoteAdmin") {
              "RemoteAdmin mode does not have a GetGracePeriodDays method or path to query" | LogMe -display -Error
              $ErrorXA = $ErrorXA + 1
              $IsSeverityErrorLevel = $True
            }
          }
          If ([string]::IsNullOrEmpty($return.LicensingName) -OR $return.LicensingName -eq "Unknown") {
            $tests.LicensingName = "ERROR", $return.LicensingName
            "LicensingName: $($return.LicensingName)" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          } ELse {
            $tests.LicensingName = "NEUTRAL", $return.LicensingName
            "LicensingName: $($return.LicensingName)" | LogMe -display -progress
          }
          If ($return.LicensingType -eq "5" -OR $return.LicensingType -eq "Unknown") {
            $tests.LicensingType = "ERROR", $return.LicensingType
            "LicensingType: $($return.LicensingType)" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          } Else {
            $tests.LicensingType = "NEUTRAL", $return.LicensingType
            "LicensingType: $($return.LicensingType)" | LogMe -display -progress
          }
          If ($return.LicenseServerList -eq "Unknown") {
            $tests.LicenseServerList = "ERROR", $return.LicenseServerList
            "LicenseServerList: $($return.LicenseServerList)" | LogMe -display -error
            $ErrorXA = $ErrorXA + 1
            $IsSeverityErrorLevel = $True
          } Else {
            $tests.LicenseServerList = "NEUTRAL", $return.LicenseServerList
            "LicenseServerList: $($return.LicenseServerList)" | LogMe -display -progress
          }
        } Else {
          $tests.RDSGracePeriod = "NORMAL", "N/A"
          $tests.RDSGracePeriodExpired = "NORMAL", "N/A"
          "This is an Azure Windows 10/11 multi-session host. Traditional RDS CALs are not relevant" | LogMe -display -progress
          $tests.TerminalServerMode = "NORMAL","AppServer"
          "TerminalServerMode: AppServer" | LogMe -display -progress
          $tests.LicensingName = "NORMAL", "AAD Per User"
          "LicensingName: AAD Per User" | LogMe -display -progress
          $tests.LicensingType = "NORMAL", "6"
          "LicensingType: 6" | LogMe -display -progress
          $tests.LicenseServerList = "NORMAL", "N/A"
        }

        ################ End RDS Licensing Details Check SECTION ###############

        # Check services
        # The Get-Service command with -ComputerName parameter made use of DCOM and such functionality is
        # removed from PowerShell 7. So we use the Invoke-Command, which uses WinRM to run a ScriptBlock
        # instead.

        $ServicesChecked = $False
        If ($IsWinRMAccessible) {
          $ServicesChecked = $True
          Try {
            $services = Invoke-Command -ComputerName $machineDNS -ErrorAction Stop -ScriptBlock {Get-Service | where-object {$_.Name -eq 'Spooler' -OR $_.Name -eq 'cpsvc'}}

            if (($services | Where-Object {$_.Name -eq "Spooler"}).Status -Match "Running") {
              "SPOOLER service running..." | LogMe
              $tests.Spooler = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "SPOOLER service stopped" | LogMe -display -error
                $tests.Spooler = "ERROR","Error"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "SPOOLER service stopped" | LogMe -display -warning
                $tests.Spooler = "WARNING","Warning"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
            }

            if (($services | Where-Object {$_.Name -eq "cpsvc"}).Status -Match "Running") {
              "Citrix Print Manager service running..." | LogMe
              $tests.CitrixPrint = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "Citrix Print Manager service stopped" | LogMe -display -error
                $tests.CitrixPrint = "ERROR","Error"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "Citrix Print Manager service stopped" | LogMe -display -warning
                $tests.CitrixPrint = "WARNING","Warning"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
            }
          }
          Catch {
            #"Error returned while checking the services" | LogMe -error; return 101
            #$ErrorXA = $ErrorXA + 1
            #$IsSeverityErrorLevel = $True
          }
        } Else {
          # Cannot connect via WinRM
        }

        If ($IsWMIAccessible -AND $ServicesChecked -eq $False) {
          Try {
            $services = Get-WmiObject -ComputerName $machineDNS -Class Win32_Service -ErrorAction Stop | where-object {$_.Name -eq 'Spooler' -OR $_.Name -eq 'cpsvc'}

            if (($services | Where-Object {$_.Name -eq "Spooler"}).State -Match "Running") {
              "SPOOLER service running..." | LogMe
              $tests.Spooler = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "SPOOLER service stopped" | LogMe -display -error
                $tests.Spooler = "ERROR","Error"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "SPOOLER service stopped" | LogMe -display -warning
                $tests.Spooler = "WARNING","Warning"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
            }

            if (($services | Where-Object {$_.Name -eq "cpsvc"}).State -Match "Running") {
              "Citrix Print Manager service running..." | LogMe
              $tests.CitrixPrint = "SUCCESS","Success"
            }
            else {
              If ($MarkSpoolerAsWarningOnly -eq 0) {
                "Citrix Print Manager service stopped" | LogMe -display -error
                $tests.CitrixPrint = "ERROR","Error"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityErrorLevel = $True
              } Else {
                "Citrix Print Manager service stopped" | LogMe -display -warning
                $tests.CitrixPrint = "WARNING","Warning"
                $ErrorXA = $ErrorXA + 1
                $IsSeverityWarningLevel = $True
              }
            }
          }
          Catch {
            #"Error returned while checking the services" | LogMe -error; return 101
            #$ErrorXA = $ErrorXA + 1
            #$IsSeverityErrorLevel = $True
          }
        } Else {
          # Cannot connect via WMI
        }

        If ($IsWinRMAccessible) {
          $ProfileStatus = Get-ProfileAndUserEnvironmentManagementServiceStatus -ComputerName:$machineDNS
          "Profile Management and User Environment Management Status:" | LogMe -display -progress
          $FSLogixEnabled = $False
          If ($ProfileStatus.FSLogixInstalled) {
            "- FSLogix Installed: $($ProfileStatus.FSLogixInstalled)" | LogMe -display -progress
            "- FSLogix ServiceRunning: $($ProfileStatus.FSLogixServiceRunning)" | LogMe -display -progress
            "- FSLogix ProfileEnabled: $($ProfileStatus.FSLogixProfileEnabled)" | LogMe -display -progress
            "- FSLogix ProfileType: $($ProfileStatus.FSLogixProfileType)" | LogMe -display -progress
            "- FSLogix ProfileTypeDescription: $($ProfileStatus.FSLogixProfileTypeDescription)" | LogMe -display -progress
            "- FSLogix OfficeEnabled: $($ProfileStatus.FSLogixOfficeEnabled)" | LogMe -display -progress
            "- FSLogix CCDLocations: $($ProfileStatus.FSLogixCCDLocations)" | LogMe -display -progress
            "- FSLogix VHDLocations: $($ProfileStatus.FSLogixVHDLocations)" | LogMe -display -progress
            "- FSLogix LogFilePath: $($ProfileStatus.FSLogixLogFilePath)" | LogMe -display -progress
            "- FSLogix RedirectionType: $($ProfileStatus.FSLogixRedirectionType)" | LogMe -display -progress
            If ($ProfileStatus.FSLogixServiceRunning -AND $ProfileStatus.FSLogixProfileEnabled -eq 1) {
              $FSLogixEnabled = $True
            }
          } Else {
            "- FSLogix is not installed" | LogMe -display -progress
          }
          If ($FSLogixEnabled) {
            $tests.FSLogixEnabled = "SUCCESS", $FSLogixEnabled
          } Else {
            $tests.FSLogixEnabled = "NORMAL", $FSLogixEnabled
          }
          $UPMEnabled = $False
          If ($ProfileStatus.UPMInstalled) {
            "- UPM Installed: $($ProfileStatus.UPMInstalled)" | LogMe -display -progress
            "- UPM ServiceRunning: $($ProfileStatus.UPMServiceRunning)" | LogMe -display -progress
            "- UPM ServiceActive: $($ProfileStatus.UPMServiceActive)" | LogMe -display -progress
            "- UPM PathToLogFile: $($ProfileStatus.UPMPathToLogFile)" | LogMe -display -progress
            "- UPM PathToUserStore: $($ProfileStatus.UPMPathToUserStore)" | LogMe -display -progress
            If ($ProfileStatus.UPMServiceRunning -AND $ProfileStatus.UPMServiceActive -eq 1) {
              $UPMEnabled = $True
            }
          } Else {
            "- UPM is not installed" | LogMe -display -progress
          }
          If ($UPMEnabled) {
            $tests.UPMEnabled = "SUCCESS", $UPMEnabled
          } Else {
            $tests.UPMEnabled = "NORMAL", $UPMEnabled
          }
          $WEMEnabled = $False
          If ($ProfileStatus.WEMInstalled) {
            "- WEM Installed: $($ProfileStatus.WEMInstalled)" | LogMe -display -progress
            "- WEM ServiceRunning: $($ProfileStatus.WEMServiceRunning)" | LogMe -display -progress
            "- WEM ServiceRunning: $($ProfileStatus.WEMServiceRunning)" | LogMe -display -progress
            "- WEM AgentRegistered: $($ProfileStatus.WEMAgentRegistered)" | LogMe -display -progress
            "- WEM AgentConfigurationSets: $($ProfileStatus.WEMAgentConfigurationSets)" | LogMe -display -progress
            "- WEM AgentCacheSyncMode: $($ProfileStatus.WEMAgentCacheSyncMode)" | LogMe -display -progress
            "- WEM AgentCachePath: $($ProfileStatus.WEMAgentCachePath)" | LogMe -display -progress
            If ($ProfileStatus.WEMServiceRunning -AND $ProfileStatus.WEMAgentRegistered) {
              $WEMEnabled = $True
            }
          } Else {
            "- WEM is not installed" | LogMe -display -progress
          }
          If ($WEMEnabled) {
            $tests.WEMEnabled = "SUCCESS", $WEMEnabled
          } Else {
            $tests.WEMEnabled = "NORMAL", $WEMEnabled
          }
        } Else {
          # Cannot connect via WinRM
        }

        # Check CrowdStrike State
        If ($ShowCrowdStrikeTests -eq 1) {
          If ($IsWinRMAccessible) {
            $return = Get-CrowdStrikeServiceStatus -ComputerName:$machineDNS
            If ($null -ne $return) {
              If ($return.CSFalconInstalled -AND $return.CSAgentInstalled) {
                "CrowdStrike Installed: True" | LogMe -display -progress
                "- CrowdStrike Windows Sensor Version: $($return.InstalledVersion)" | LogMe -display -progress
                $tests.CSVersion = "NORMAL", $return.InstalledVersion
                "- CrowdStrike Company ID (CID): $($return.CID)" | LogMe -display -progress
                $tests.CSCID = "NORMAL", $return.CID
                "- CrowdStrike Sensor Grouping Tags: $($return.SensorGroupingTags)" | LogMe -display -progress
                $tests.CSGroupTags = "NORMAL", $return.SensorGroupingTags
                "- CrowdStrike VDI switch: $($return.VDI)" | LogMe -display -progress
                If ($return.CSFalconServiceRunning -AND $return.CSAgentServiceRunning -AND (![string]::IsNullOrEmpty($return.AID))) {
                  $tests.CSEnabled = "SUCCESS", $True
                  "- CrowdStrike Agent ID (AID): $($return.AID)" | LogMe -display -progress
                  $tests.CSAID = "NORMAL", $return.AID
                } Else {
                  $tests.CSEnabled = "WARNING", $False
                  If ([string]::IsNullOrEmpty($return.AID)) {
                    "- CrowdStrike Agent ID (AID) is missing" | LogMe -display -warning
                    $tests.CSAID = "NORMAL", "Missing"
                  } Else {
                    "- CrowdStrike is installed, but not running" | LogMe -display -warning
                  }
                  $IsSeverityWarningLevel = $True
                }
              } else {
                "CrowdStrike Installed: False" | LogMe -display -progress
              }
            } else {
              "Unable to get the CrowdStrike Service Status" | LogMe -display -error
              $IsSeverityErrorLevel = $True
            }
          }
        }

      }#If can connect via WinRM or WMI
      else {
        "WinRM or WMI connection not possible" | LogMe -display -error
        $ErrorXA = $ErrorXA + 1
        $IsSeverityErrorLevel = $True
      } # Closing else cannot connect via WinRM or WMI

    } # Close off $Powered -eq "On", "Unknown", or "Unmanaged"

    # Column Serverload
    $Serverload = $XAmachine | ForEach-Object{ $_.LoadIndex }
    "Serverload: $Serverload" | LogMe -display -progress
    if ($Serverload -ge $loadIndexError) { $tests.Serverload = "ERROR", $Serverload ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityErrorLevel = $True }
    elseif ($Serverload -ge $loadIndexWarning) { $tests.Serverload = "WARNING", $Serverload ; $ErrorXA = $ErrorXA + 1 ; $IsSeverityWarningLevel = $True }
    else { $tests.Serverload = "SUCCESS", $Serverload }
  
    # Column RegistrationState
    $RegState = $XAmachine| ForEach-Object{ $_.RegistrationState }
    if ($RegState -ne "Registered") {
      if ($Powered -eq "Off") {
        "RegistrationState: $RegState" | LogMe -display -progress
        $tests.RegState = "NEUTRAL", $RegState
      } else {
        "RegistrationState: $RegState" | LogMe -display -error
        $tests.RegState = "ERROR", $RegState
        $ErrorXA = $ErrorXA + 1
        $IsSeverityErrorLevel = $True
      }
    } else {
      "RegistrationState: $RegState" | LogMe -display -progress
      $tests.RegState = "SUCCESS", $RegState
    }

    # Column MaintMode
    $MaintMode = $XAmachine | ForEach-Object{ $_.InMaintenanceMode }
    "MaintenanceMode: $MaintMode" | LogMe -display -progress
    if ($MaintMode) {
      $objMaintenance = $null
      Try {
        $objMaintenance = $Maintenance | Where-Object { $_.TargetName.ToUpper() -eq $XAmachine.MachineName.ToUpper() } | Select-Object -First 1
      }
      Catch {
        # Avoid the error "The property 'TargetName' cannot be found on this object."
      }
      If ($null -ne $objMaintenance){$MaintenanceModeOn = ("ON, " + $objMaintenance.User)} Else {$MaintenanceModeOn = "ON"}
      # The Get-LogLowLevelOperation cmdlet will tell us who placed a machine into maintanance mode. However, the Get-BrokerMachine cmdlet
      # will provide the MaintenanceReason, where the underlying reason, if manually entered, is stored in the MetadataMap property as part
      # of a Dictionary object.
      $MetadataMapDictionary = $XAmachine | ForEach-Object{ $_.MetadataMap }
      foreach ($key in $MetadataMapDictionary.Keys) {
        if ($key -eq "MaintenanceModeMessage") {
          $MaintenanceModeOn = $MaintenanceModeOn + ", " + $MetadataMapDictionary[$key]
        }
      }
      "MaintenanceModeInfo: $MaintenanceModeOn" | LogMe -display -progress
      $tests.MaintMode = "WARNING", $MaintenanceModeOn
      $ErrorXA = $ErrorXA + 1
      $IsSeverityWarningLevel = $True
    }
    else { $tests.MaintMode = "SUCCESS", "OFF" }

    # Column VDAVersion AgentVersion
    $VDAVersion = $XAmachine | ForEach-Object{ $_.AgentVersion }
    $Found = $False
    If ($SupportedVDAVersions.Count -gt 0) {
      If (!([String]::IsNullOrEmpty($SupportedVDAVersions[0]))) {
        ForEach ($SupportedVDAVersion in $SupportedVDAVersions) {
          If ($VDAVersion -Like $SupportedVDAVersion) {
            $Found = $True
            break
          }
        }
        If ($SupportedVDAVersions -contains $VDAVersion) {
          $Found = $True
        }
      }
    } Else {
      $Found = $True
    }
    If ($Found) {
      "VDAVersion: $VDAVersion" | LogMe -display -progress
      $tests.VDAVersion = "NEUTRAL", $VDAVersion
    } Else {
      "VDAVersion: $VDAVersion" | LogMe -display -warning
      $tests.VDAVersion = "WARNING", $VDAVersion
      $ErrorXA = $ErrorXA + 1
      $IsSeverityWarningLevel = $True
    }

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
  
    # Column MCSImageOutOfDate
    $ProvisioningType = $XAmachine | ForEach-Object{ $_.ProvisioningType }
    # The Get-BrokerMachine cmdlet has a ProvisioningType property, but we can also get this from the Machine Catalogs we have already collected.
    # The Machine Catalogs should exist in the $Catalogs variable, so test this first before collecting them again using the Get-BrokerCatalog cmdlet.
    # Have left the next 13 lines of code in for future reference.
    # $ProvisioningType = ""
    # $GetCatalogs = $True
    # If (($Catalogs | Measure-Object).Count -gt 0) {
    #   $ProvisioningType = ($Catalogs | Where-Object {$_.Name -eq $CatalogName} | Select-Object -ExpandProperty ProvisioningType)
    #   $GetCatalogs = $False
    # }
    # If ($GetCatalogs) {
    #   If ($CitrixCloudCheck -ne 1) { 
    #     $ProvisioningType = (Get-BrokerCatalog -AdminAddress $AdminAddress -Name $CatalogName | Select-Object -ExpandProperty ProvisioningType)
    #   } Else {
    #     $ProvisioningType = (Get-BrokerCatalog -Name $CatalogName | Select-Object -ExpandProperty ProvisioningType)
    #   }
    # }
    If ($ProvisioningType -eq "MCS") {
      $MCSImageOutOfDate = $XAmachine | ForEach-Object{ $_.ImageOutOfDate }
      if ($MCSImageOutOfDate -eq $true) {
        if ($Powered -eq "Off") {
          "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -progress
          $tests.MCSImageOutOfDate = "NEUTRAL", $MCSImageOutOfDate
        } else {
          "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -error
          $tests.MCSImageOutOfDate = "ERROR", $MCSImageOutOfDate
          $ErrorXA = $ErrorXA + 1
          $IsSeverityErrorLevel = $True
        }
      } else {
        "MCSImageOutOfDate: $MCSImageOutOfDate" | LogMe -display -progress
        if ($Powered -eq "Off") {
          $tests.MCSImageOutOfDate = "NEUTRAL", $MCSImageOutOfDate
        } else {
          $tests.MCSImageOutOfDate = "SUCCESS", $MCSImageOutOfDate
        }
      }
    }

    # Column ConnectedUsers
    $ConnectedUsers = $XAmachine | ForEach-Object{ $_.AssociatedUserNames }
    $ConnectedUsersCount = 0
    Try {
      # The AssociatedUserNames property is a System.String[] (an array of strings) object, which is a type of System.Object[].
      # - With multiple users connected, the ForEach-Object converts it into a System.Object[] object, which is correct.
      # - With 1 user connected, the ForEach-Object casts it into a System.String object, which is incorrect.
      # - Therefore, with 1 user connected, the count is 0, which is incorrect output.
      # - So we use type casting for single string to array to address this.
      If ($ConnectedUsers.GetType().FullName -eq "System.String") {
        $ConnectedUsers = [System.Object[]]@($ConnectedUsers)
        # $ConnectedUsers will be a System.Object[] containing one string element.
      }
      $ConnectedUsersCount = $ConnectedUsers.Count
    }
    Catch {
      # The AssociatedUserNames property is of type "System.Object[]
      # Wrapping this in a Try/Catch to prevent errors like...
      # - "The property 'Count' cannot be found on this object."
      # - "You cannot call a method on a null-valued expression."
    }
    "$ConnectedUsersCount Connected users: $ConnectedUsers" | LogMe -display -progress
    $tests.ConnectedUsers = "NEUTRAL", $ConnectedUsers

    # Column ActiveSessions
    $ActiveSessions = $XAmachine | ForEach-Object{ $_.SessionCount }
    If ([int]$ActiveSessions -gt $ConnectedUsersCount) {
      $tests.ActiveSessions = "ERROR", $ActiveSessions
      "The Active Sessions count does match the Connected Users count: $ActiveSessions" | LogMe -display -error
      $ErrorXA = $ErrorXA + 1
      $IsSeverityErrorLevel = $True
    } Else {
      $tests.ActiveSessions = "NEUTRAL", $ActiveSessions
      "Active Sessions: $ActiveSessions" | LogMe -display -progress
    }

    # Column LastConnectionTime
    $yellow =((Get-Date).AddDays(-30).ToString('yyyy-MM-dd HH:mm:s'))
    $red =((Get-Date).AddDays(-90).ToString('yyyy-MM-dd HH:mm:s'))
    $machineLastConnectionTime = $XAmachine | ForEach-Object{ $_.LastConnectionTime }
    if ([string]::IsNullOrWhiteSpace($machineLastConnectionTime))
    {
      $tests.LastConnectionTime = "NEUTRAL", "NO DATA"
    }
    elseif ($machineLastConnectionTime -lt $red)
    {
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -ERROR
      $tests.LastConnectionTime = "ERROR", $machineLastConnectionTime
      $ErrorXA = $ErrorXA + 1
      $IsSeverityErrorLevel = $True
    } 	
    elseif ($machineLastConnectionTime -lt $yellow)
    {
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -WARNING
      $tests.LastConnectionTime = "WARNING", $machineLastConnectionTime
      $ErrorXA = $ErrorXA + 1
      $IsSeverityWarningLevel = $True
    }
    else 
    {
      $tests.LastConnectionTime = "SUCCESS", $machineLastConnectionTime
      "LastConnectionTime: $machineLastConnectionTime" | LogMe -display -progress
    }

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename
  
    # Fill $tests into array if error occured OR $ShowOnlyErrorXA = 0
    # Check if error exists on this vdi
    if ($ShowOnlyErrorXA -eq 0 ) { $allXenAppResults.$machineDNS = $tests }
    else {
      if ($ErrorXA -gt 0) { $allXenAppResults.$machineDNS = $tests }
      else { "$machineDNS is ok, no output into HTML-File" | LogMe -display -progress }
    }

    If ($tests.Count -gt 0) {
      If ($CheckOutputSyslog) {
        # Set up the severity of the log entry based on the output of each test.
        $Severity = "Informational"
        If ($IsSeverityWarningLevel) { $Severity = "Warning" }
        If ($IsSeverityErrorLevel) { $Severity = "Error" }
        # Setup the PSCustomObject that will become the Data within the Structured Data
        $Data = [PSCustomObject]@{
          'MultiSessionHost' = $machineDNS
        }
        $allXenAppResults.$machineDNS.GetEnumerator() | ForEach-Object {
          $MyKey = $_.Key -replace " ", ""
          $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
        }
        $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
        If ($SyslogFileOnly) {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -FileOnly
        } Else {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
        }
      }
    }
    " --- " | LogMe -display -progress

  } # Closing foreach $XAmachine

} # Closing if $ShowXenAppTable

else { "XenApp Check skipped because ShowXenAppTable = 0" | LogMe -display -progress }
  
"####################### Check END  ##########################################################" | LogMe -display -progress

#==============================================================================================
# End of XenApp/RDSH (multi-session) Check
#==============================================================================================

#==============================================================================================
# Start of Stuck Sessions Check
#==============================================================================================

#  Check Stuck Sessions only if $ShowStuckSessionsTable is 1
if($ShowStuckSessionsTable -eq 1 ) {

  "Check Stuck Sessions ########################################################################" | LogMe -display -progress
  " " | LogMe -display -progress

  $allStuckSessionResults = @{}
  $AllStuckSessions = @()

  If (($ActualExcludedCatalogs | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Catalogs from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedCatalog in $ActualExcludedCatalogs) {
      "- $ActualExcludedCatalog" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedDeliveryGroups | Measure-Object).Count -gt 0) {
    "Excluding machines from the following Delivery Groups from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedDeliveryGroup in $ActualExcludedDeliveryGroups) {
      "- $ActualExcludedDeliveryGroup" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }
  If (($ActualExcludedBrokerTags | Measure-Object).Count -gt 0) {
    "Excluding machines with the following Tags from these tests..." | LogMe -display -progress
    ForEach ($ActualExcludedBrokerTag in $ActualExcludedBrokerTags) {
      "- $ActualExcludedBrokerTag" | LogMe -display -progress
    }
    " " | LogMe -display -progress
  }

  # Get all sessions that are in a Connected state with no User with the logon in progress for more than 10 minutes
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { SessionState -eq 'Connected' -AND UserName -eq $null -AND LogonInProgress -eq $True -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { SessionState -eq 'Connected' -AND UserName -eq $null -AND LogonInProgress -eq $True -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Get all sessions where the app state of the session is in a PreLogon state for more than 10 minutes
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { AppState -eq 'PreLogon' -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { AppState -eq 'PreLogon' -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Get all non-Active Sessions that have not changed state for more than x hours
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { SessionState -ne 'Active'} | Where {$_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { SessionState -ne 'Active'} | Where {$_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Get all non-Active Sessions that have a Client Address of 127.0.0.1 and have been in this state for more than 10 minutes. They may be stuck and need logging off or rebooting.
  # This needs more assessment
  #If ($CitrixCloudCheck -ne 1) {
    #$AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { SessionState -ne 'Active' -AND ClientAddress -eq "127.0.0.1" -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #} Else {
    #$AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { SessionState -ne 'Active' -AND ClientAddress -eq "127.0.0.1" -AND SessionStateChangeTime -lt '0:10' } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #}

  # Get all unbrokered sessions that have been in a state for more than x hours
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { ConnectionMode -eq "Unbrokered"} | Where {$_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { ConnectionMode -eq "Unbrokered"} | Where {$_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Get all sessions where the protocol is Console. They may be stuck sessions and need logging off or rebooting.
  # This needs more assessment
  #If ($CitrixCloudCheck -ne 1) {
    #$AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { Protocol -eq "Console"} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #} Else {
    #$AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { Protocol -eq "Console"} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #}

  # Get all sessions where the Client Address is 127.0.0.1 where the session is not in a disconnected state. They may be stuck sessions and need logging off or rebooting.
  # This needs more assessment
  #If ($CitrixCloudCheck -ne 1) {
    #$AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { ClientAddress -eq '127.0.0.1' -AND SessionState -ne 'Disconnected'} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #} Else {
    #$AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { ClientAddress -eq '127.0.0.1' -AND SessionState -ne 'Disconnected'} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  #}

  # Get all sessions that have not been used in more than x hours
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { BrokeringTime -ne $null -AND BrokeringDuration -gt 0 -AND SessionStateChangeTime -ne $null} | where { $_.BrokeringTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours)) -AND $_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { BrokeringTime -ne $null -AND BrokeringDuration -gt 0 -AND SessionStateChangeTime -ne $null} | where { $_.BrokeringTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours)) -AND $_.SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))} | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Get all sessions where the BrokeringTime has not changed in x hours.
  If ($CitrixCloudCheck -ne 1) {
    $AllStuckSessions += Get-BrokerSession -AdminAddress $AdminAddress -MaxRecordCount $maxmachines -Filter { BrokeringTime -ne $null -AND BrokeringDuration -gt 0 -AND EstablishmentDuration -gt 0} | where { $_.BrokeringTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours)) } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  } Else {
    $AllStuckSessions += Get-BrokerSession -MaxRecordCount $maxmachines -Filter { BrokeringTime -ne $null -AND BrokeringDuration -gt 0 -AND EstablishmentDuration -gt 0} | where { $_.BrokeringTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours)) } | Where-Object {($_.catalogname -notin $ActualExcludedCatalogs -AND $_.desktopgroupname -notin $ActualExcludedDeliveryGroups)}
  }

  # Sort sessions to remove duplicates
  $AllStuckSessions = $AllStuckSessions | Sort-Object -Property MachineName -Unique

  # Filter out sessions that contain accounts names in the $ExcludedStuckSessionUserAccounts array.
  $Found = $False
  If ($ExcludedStuckSessionUserAccounts.Count -gt 0) {
    If (!([String]::IsNullOrEmpty($ExcludedStuckSessionUserAccounts[0]))) {
      ForEach ($ExcludedStuckSessionUserAccount in $ExcludedStuckSessionUserAccounts) {
        # Ensure there is no match for the UserName or the BrokeringUserName
        $AllStuckSessions = $AllStuckSessions | ForEach-Object {
          $Found = $False
          ForEach ($ExcludedStuckSessionUserAccount in $ExcludedStuckSessionUserAccounts) {
            If ($_.UserName -eq $ExcludedStuckSessionUserAccount -OR $_.UserName -like $ExcludedStuckSessionUserAccount -OR $_.BrokeringUserName -eq $ExcludedStuckSessionUserAccount -OR $_.BrokeringUserName -like $ExcludedStuckSessionUserAccount) {
              $Found = $True
              break
            }
          }
          If ($Found -eq $False) {
            $_
          }
        }
      }
    }
  }

  # Filter out machines that contain any $ActualExcludedBrokerTags. We need to do this here using the Get-BrokerMachine cmdlet because the Get-BrokerSession cmdlet does not have a tags property.
  If (($ActualExcludedBrokerTags | Measure-Object).Count -gt 0) {
    $nontaggedMachineIds = @()
    If ($CitrixCloudCheck -ne 1) {
      $nontaggedMachineIds = (Get-BrokerMachine -MaxRecordCount $maxmachines -AdminAddress $AdminAddress | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0}) | Select-Object -ExpandProperty DNSName
    } Else {
      $nontaggedMachineIds = (Get-BrokerMachine -MaxRecordCount $maxmachines | Where-Object {@(Compare-Object $_.tags $ActualExcludedBrokerTags -IncludeEqual | Where-Object {$_.sideindicator -eq '=='}).count -eq 0}) | Select-Object -ExpandProperty DNSName
    }
    $AllStuckSessions = $AllStuckSessions | Where-Object {$nontaggedMachineIds -contains $_.DNSName}
  }

  foreach ($StuckSession in $AllStuckSessions) {

    $tests = @{}

    # Column Name of Machine
    $machineDNS = $StuckSession | %{ $_.DNSName }
    "Machine: $machineDNS" | LogMe -display -progress

    # Column CatalogName
    $CatalogName = $StuckSession | %{ $_.CatalogName }
    "Catalog: $CatalogName" | LogMe -display -progress
    $tests.CatalogName = "NEUTRAL", $CatalogName

    # Column DesktopGroupName
    $DesktopGroupName = $StuckSession | %{ $_.DesktopGroupName }
    "DesktopGroupName: $DesktopGroupName" | LogMe -display -progress
    $tests.DesktopGroupName = "NEUTRAL", $DesktopGroupName

    # Column UserName
    $UserName = $StuckSession | %{ $_.UserName}
    $BrokeringUserName = $StuckSession | %{ $_.BrokeringUserName}
    if (!([String]::IsNullOrEmpty($UserName))) {
      "UserName: $UserName" | LogMe -display -progress
      $tests.UserName = "NEUTRAL", $UserName
    } elseif (!([String]::IsNullOrEmpty($BrokeringUserName))) {
      "UserName: $BrokeringUserName" | LogMe -display -progress
      $tests.UserName = "WARNING", $BrokeringUserName
    } else {
      "UserName: $UserName" | LogMe -display -progress
      $tests.UserName = "ERROR", $UserName
    }

    # Column SessionState
    $SessionState = $StuckSession | %{ $_.SessionState}
    "SessionState: $SessionState" | LogMe -display -progress
    if ($SessionState -eq "Connected") {
      $tests.SessionState = "WARNING", $SessionState
    } else { $tests.SessionState = "SUCCESS", $SessionState}

    # Column AppState
    $AppState = $StuckSession | %{ $_.AppState}
    "AppState: $AppState" | LogMe -display -progress
    if ($AppState -eq "PreLogon") {
      $tests.AppState = "WARNING", $AppState
    } else { $tests.AppState = "SUCCESS", $AppState}

    # Column SessionStateChangeTime
    $SessionStateChangeTime = $StuckSession | %{ $_.SessionStateChangeTime}
    $BrokeringTime = $StuckSession | %{ $_.BrokeringTime}
    "SessionStateChangeTime: $SessionStateChangeTime" | LogMe -display -progress
    if ($SessionState -eq "Disconnected" -AND $SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))) {
      $tests.SessionStateChangeTime = "ERROR", $SessionStateChangeTime
    } elseif ($BrokeringTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours)) -AND $SessionStateChangeTime -lt ((Get-Date).AddHours(-$MaxDisconnectTimeInHours))) {
      $tests.SessionStateChangeTime = "ERROR", $SessionStateChangeTime
    } elseif (($AppState -eq "PreLogon" -OR $SessionState -eq "Connected" ) -AND $SessionStateChangeTime -lt ((Get-Date).AddMinutes(-10))) {
      $tests.SessionStateChangeTime = "ERROR", $SessionStateChangeTime
    } else { $tests.SessionStateChangeTime = "SUCCESS", $SessionStateChangeTime }

    # Column LogonInProgress
    $LogonInProgress = $StuckSession | %{ $_.LogonInProgress}
    "LogonInProgress: $LogonInProgress" | LogMe -display -progress
    if ($LogonInProgress) {
      $tests.LogonInProgress = "ERROR", $LogonInProgress
    } else { $tests.LogonInProgress = "SUCCESS", $LogonInProgress }

    # Column LogoffInProgress
    $LogoffInProgress = $StuckSession | %{ $_.LogoffInProgress}
    "LogoffInProgress: $LogoffInProgress" | LogMe -display -progress
    if ($LogoffInProgress) {
      $tests.LogoffInProgress = "ERROR", $LogoffInProgress
    } else { $tests.LogoffInProgress = "SUCCESS", $LogoffInProgress }

    # Column ClientAddress
    $ClientAddress = $StuckSession | %{ $_.ClientAddress}
    "ClientAddress: $ClientAddress" | LogMe -display -progress
    if ($ClientAddress -eq "127.0.0.1") {
      $tests.ClientAddress = "WARNING", $ClientAddress
    } else { $tests.ClientAddress = "SUCCESS", $ClientAddress }

    # Column ConnectionMode
    $ConnectionMode = $StuckSession | %{ $_.ConnectionMode}
    "ConnectionMode: $ConnectionMode" | LogMe -display -progress
    if ($ConnectionMode -eq "Unbrokered") {
      $tests.ConnectionMode = "WARNING", $ConnectionMode
    } else { $tests.ConnectionMode = "SUCCESS", $ConnectionMode}

    # Column Protocol
    $Protocol = $StuckSession | %{ $_.Protocol}
    "Protocol: $Protocol" | LogMe -display -progress
    if ($Protocol -eq "Console") {
      $tests.Protocol = "ERROR", $Protocol
    } elseif ($Protocol -eq "RDP") {
      $tests.Protocol = "WARNING", $Protocol
    } else { $tests.Protocol = "SUCCESS", $Protocol }

    " --- " | LogMe -display -progress

    # Add the SiteName to the tests for the Syslog output
    $tests.SiteName = "NORMAL", $sitename

    $allStuckSessionResults.$machineDNS = $tests 

    If ($tests.Count -gt 0) {
      If ($CheckOutputSyslog) {
        # Set up the severity of the log entry based on the output of each test.
        $Severity = "Informational"
        If ($IsSeverityWarningLevel) { $Severity = "Warning" }
        If ($IsSeverityErrorLevel) { $Severity = "Error" }
        # Setup the PSCustomObject that will become the Data within the Structured Data
        $Data = [PSCustomObject]@{
          'StuckSession' = $machineDNS
        }
        $allStuckSessionResults.$machineDNS.GetEnumerator() | ForEach-Object {
          $MyKey = $_.Key -replace " ", ""
          $Data | Add-Member -MemberType NoteProperty $MyKey -Value $_.Value[1]
        }
        $sdString = ConvertTo-StructuredData -Id $StructuredDataID -Data $Data -AllowMoreParamChars
        If ($SyslogFileOnly) {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -FileOnly
        } Else {
          Write-IetfSyslogEntry -AppName "$SyslogAppName" -Severity $Severity -Message "$machineDNS" `
                                -StructuredData $sdString -MsgId "$SyslogMsgId" -CollectorType Syslog `
                                -LogFilePath "$resultsSyslog" -SyslogServer $SyslogServer
        }
      }
    }
  }
} # Close if $ShowStuckSessionsTable
else { "Stuck Sessions Check skipped because ShowStuckSessionsTable = 0" | LogMe -display -progress	}

"####################### Check END  ##########################################################" | LogMe -display -progress

#==============================================================================================
# End of Stuck Sessions Check
#==============================================================================================

# ======= Write all results to an html file =================================================
# Add Site Name and Version to EnvironmentName for e-mail subject
$EnvironmentNameOut = "$EnvironmentName for Site $sitename"
$XDmajor, $XDminor = $controllerversion.Split(".")[0..1]
$XDVersion = "$XDmajor.$XDminor"
$EnvironmentNameOut = "$EnvironmentNameOut v$XDVersion"
$emailSubject = ("$EnvironmentNameOut Report - " + $ReportDate)

Write-Host ("Saving results to html report: " + $resultsHTM)
writeHtmlHeader "$EnvironmentNameOut Report" $resultsHTM

# Write Table with the Failures #FUTURE !!!!
#"Adding Failures output to HTML" | LogMe -display -progress 
#writeTableHeader $resultsHTM $CTXFailureFirstheaderName $CTXFailureHeaderNames $CTXFailureTableWidth
#$ControllerResults | ForEach-Object{ writeData $CTXFailureResults $resultsHTM $CTXFailureFirstheaderName }
#writeTableFooter $resultsHTM

# Write Table with the Controllers
if ($CitrixCloudCheck -ne 1 ) {
  "Adding Controller output to HTML" | LogMe -display -progress 
  writeTableHeader $resultsHTM $XDControllerFirstheaderName $XDControllerHeaderNames $XDControllerTableWidth
  $ControllerResults | ForEach-Object{ writeData $ControllerResults $resultsHTM $XDControllerHeaderNames }
  writeTableFooter $resultsHTM
}
else { "No Controller output in HTML (CitrixCloud) " | LogMe -display -progress }

if ($CitrixCloudCheck -eq 1 -AND $ShowCloudConnectorTable -eq 1 ) {
  # Write Table with the CloudConnectorServers
  "Adding Cloud Connector output to HTML" | LogMe -display -progress 
  writeTableHeader $resultsHTM $CCFirstheaderName $CCHeaderNames $CCTableWidth
  $CCResults | ForEach-Object{ writeData $CCResults $resultsHTM $CCHeaderNames }
  writeTableFooter $resultsHTM
}
else { "No Cloud Connector output in HTML (CitrixCloud) " | LogMe -display -progress }

if ($ShowStorefrontTable -eq 1 ) {
  # Write Table with the StorefrontServers
  "Adding Storefront output to HTML" | LogMe -display -progress 
  writeTableHeader $resultsHTM $SFFirstheaderName $SFHeaderNames $SFTableWidth
  $SFResults | ForEach-Object{ writeData $SFResults $resultsHTM $SFHeaderNames }
  writeTableFooter $resultsHTM
}
else { "No Storefront output in HTML (CitrixCloud) " | LogMe -display -progress }

# Write Table with the License
If ($CitrixCloudCheck -ne 1) {
  If ($CTXLicResults.Count -gt 0) {
    "Adding License output to HTML" | LogMe -display -progress 
    writeTableHeader $resultsHTM $CTXLicFirstheaderName $CTXLicHeaderNames $CTXLicTableWidth
    $CTXLicResults | ForEach-Object{ writeData $CTXLicResults $resultsHTM $CTXLicHeaderNames }
    writeTableFooter $resultsHTM
  } Else { "The License output has not been added to the HTML as it contains no data. This may be because it is License Activation Service (LAS) enabled." | LogMe -display -progress }
}
else { "No License output in HTML (CitrixCloud) " | LogMe -display -progress }

# Write Table with the Machine Catalogs
"Adding Machine Catalog output to HTML" | LogMe -display -progress 
writeTableHeader $resultsHTM $CatalogHeaderName $CatalogHeaderNames $CatalogTablewidth
$CatalogResults | ForEach-Object{ writeData $CatalogResults $resultsHTM $CatalogHeaderNames}
writeTableFooter $resultsHTM

# Write Table with the Assignments (Delivery Groups)
"Adding Delivery Group output to HTML" | LogMe -display -progress 
writeTableHeader $resultsHTM $AssigmentFirstheaderName $vAssigmentHeaderNames $Assigmenttablewidth
$AssigmentsResults | ForEach-Object{ writeData $AssigmentsResults $resultsHTM $vAssigmentHeaderNames }
writeTableFooter $resultsHTM

# Write Table with the Connection Failures from Broker Connection Log
If($ShowBrokerConnectionFailuresTable -eq 1 ) {
  If ($BrokerConnectionLogResults.Count -gt 0) {
    "Adding Connection Failures output to HTML" | LogMe -display -progress 
    writeTableHeader $resultsHTM $BrkrConFailureFirstheaderName $BrkrConFailureHeaderNames $BrkrConFailureTableWidth
    $BrokerConnectionLogResults | ForEach-Object{ writeData $BrokerConnectionLogResults $resultsHTM $BrkrConFailureHeaderNames }
    writeTableFooter $resultsHTM
  } Else { "The Connection Failures output has not been added to the HTML as it contains no data" | LogMe -display -progress }
}
else { "No Connection Failures output in HTML " | LogMe -display -progress }

# Write Table with the Hypervisor Connections
If ($HypervisorConnectionResults.Count -gt 0) {
  "Adding Hypervisor Connection output to HTML" | LogMe -display -progress 
  writeTableHeader $resultsHTM $HypervisorConnectionFirstheaderName $HypervisorConnectionHeaderNames $HypervisorConnectiontablewidth
  $HypervisorConnectionResults | ForEach-Object{ writeData $HypervisorConnectionResults $resultsHTM $HypervisorConnectionHeaderNames }
  writeTableFooter $resultsHTM
} Else { "The Hypervisor Connection output has not been added to the HTML as it contains no data" | LogMe -display -progress }

# Write Table with all XenApp/RDSH (multi-session) Servers
if ($ShowXenAppTable -eq 1 ) {
  If ($allXenAppResults.Count -gt 0) {
    "Adding XenApp/RDSH (multi-session) output to HTML" | LogMe -display -progress 
    writeTableHeader $resultsHTM $XenAppFirstheaderName $XenAppHeaderNames $XenApptablewidth
    If ([string]::IsNullOrEmpty($SortXenAppTableByHeaderName)) {
      $allXenAppResults | ForEach-Object{ writeData $allXenAppResults $resultsHTM $XenAppHeaderNames }
    } Else {
      $allXenAppResults | ForEach-Object{ writeDataSortedByHeaderName $allXenAppResults $resultsHTM $XenAppHeaderNames $SortXenAppTableByHeaderName }
    }
    writeTableFooter $resultsHTM
  } Else { "The XenApp/RDSH (multi-session) output has not been added to the HTML as it contains no data" | LogMe -display -progress }
}
else { "No XenApp/RDSH (multi-session) output in HTML " | LogMe -display -progress }

# Write Table with all Desktops
if ($ShowDesktopTable -eq 1 ) {
  If ($allResults.Count -gt 0) {
    "Adding VDI (single-session) output to HTML" | LogMe -display -progress 
    writeTableHeader $resultsHTM $VDIFirstheaderName $VDIHeaderNames $VDItablewidth
    If ([string]::IsNullOrEmpty($SortDesktopTableByHeaderName)) {
      $allResults | ForEach-Object{ writeData $allResults $resultsHTM $VDIHeaderNames }
    } Else {
      $allResults | ForEach-Object{ writeDataSortedByHeaderName $allResults $resultsHTM $VDIHeaderNames $SortDesktopTableByHeaderName }
    }
    writeTableFooter $resultsHTM
  } Else { "The VDI (single-session) output has not been added to the HTML as it contains no data" | LogMe -display -progress }
}
else { "No VDI (single-session) output in HTML " | LogMe -display -progress }

# Write Table with all Stuck Sessions
if ($ShowStuckSessionsTable -eq 1 ) {
  If ($allStuckSessionResults.Count -gt 0) {
    "Adding Stuck Session output to HTML" | LogMe -display -progress 
    writeTableHeader $resultsHTM $StuckSessionsFirstheaderName $StuckSessionsHeaderNames $StuckSessionstablewidth
    $allStuckSessionResults | ForEach-Object{ writeData $allStuckSessionResults $resultsHTM $StuckSessionsHeaderNames }
    writeTableFooter $resultsHTM
  } Else { "The Stuck Session output has not been added to the HTML as it contains no data" | LogMe -display -progress }
}
else { "No Stuck Session output in HTML " | LogMe -display -progress }

"Adding footer output to HTML" | LogMe -display -progress
If ($CitrixCloudCheck -ne 1) {
  writeHtmlFooter -fileName $resultsHTM
} Else {
  writeHtmlFooter -fileName $resultsHTM -cloud
}

"HTML file created: $resultsHTM" | LogMe -display -progress

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
if (!([String]::IsNullOrEmpty($emailCC))) {
  $emailMessage.CC.Add( $emailCC )
}
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

    "Report sent via email" | LogMe -display -progress

}#end of IF CheckSendMail
else{

    "Report not sent via email because CheckSendMail = 0" | LogMe -display -progress

}#Skip Send Mail

} # Close off $scriptBlockExecute

If ($UseRunspace) {
  $PSinstance = [PowerShell]::Create().AddScript($scriptBlockExecute).AddParameter('ParameterFilePath', "$ParameterFilePath").AddParameter('ParameterFile', "$ParameterFile").AddParameter('outputpath', "$outputpath").AddParameter('snapins', $SnapIns)
  $PSinstance.RunspacePool = $RunspacePool
  $runspaces.Add(([pscustomobject]@{
        Id = $ParameterFile
        PowerShell = $PSinstance
        Handle = $PSinstance.BeginInvoke()
    })) | out-null
  write-verbose "$(Get-Date): Runspace created: $ParameterFile" -verbose
} Else {
  # It was super challenging to pass named parameters to Invoke-Command -ScriptBlock.
  # Whilst an answer can be found here, it is flawed, because it stringifies all parameters passed:
  # - https://stackoverflow.com/questions/27794898/powershell-pass-named-parameters-to-argumentlist
  # - $formattedParams = &{ $args } @hashArgs # this converts to strings and should not be used
  #   This does not preserve the hashtable as-is. Instead, it expands it into an array of key/value pairs, effectively turning:
  #   @{ snapins = $snapins }
  #   into:
  #   @("snapins", "<stringified snapins>")
  $sb = [scriptblock]::Create($scriptBlockExecute)
  # Use a hashtable with splatting — don't stringify anything
  $hashArgs = @{
    ParameterFilePath = $ParameterFilePath
    ParameterFile     = $ParameterFile
    OutputPath        = $OutputPath
    Snapins           = $Snapins
  }
  # Use the following for local execution.
  & $sb @hashArgs
  # For remoting, we typically use ArgumentList with a param block like this:
  # Invoke-Command -ScriptBlock $sb -ArgumentList @($ParameterFilePath, $ParameterFile, $OutputPath, $Snapins)
  # However, Invoke-Command -ArgumentList does not support named parameters. It's possitional only. Therefore, it passes arguments
  # by position to the param() block inside the scriptblock. To use named parameters here we need to pass an object of arguments to
  # the -ArgumentList.
  # Wrap all parameters in a single object (like a hashtable or PSCustomObject)
  $paramBundle = [PSCustomObject]@{
    ParameterFilePath = $ParameterFilePath
    ParameterFile     = $ParameterFile
    OutputPath        = $OutputPath
    Snapins           = $Snapins
  }
  # Then we use -ArgumentList to pass it as one object, which passes it to the first parameter of the scriptblock that much be defined
  # as a PSCustomObject.
  #Invoke-Command -ScriptBlock $sb -ArgumentList $paramBundle
}

} # Close off ForEach $ParameterFile

If ($UseRunspace) {
  $CountRunspaces = ($runspaces | Measure-Object).count
  Write-Verbose "$(Get-Date): $CountRunspaces runspaces were created." -verbose
  Write-Verbose "$(Get-Date): --------WAITING FOR JOBS TO COMPLETE---------" -verbose
  Write-Verbose "$(Get-Date): Check the output logs for progress." -verbose
  # Use a sleep timer to control CPU utilization
  $SleepTimer = 1000
  while ($runspaces.Handle -ne $null)
  {
    foreach ($runspace in $runspaces)
    {
       If ($runspace.Handle -ne $null) {
         If ($runspace.Handle.IsCompleted) {
          write-verbose "$(Get-Date): Runspace complete: $($runspace.Id)" -verbose
          $runspace.Handle = $null
         }
       }
    }
    Start-Sleep -Milliseconds $SleepTimer
  }
  $RunspacePool.Close()
  $RunspacePool.Dispose()
  Write-Verbose "$(Get-Date): Script complete!" -verbose
}

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
# - 1.4.5.1, Ping to Desktop or AppServer is no more red because not critical
# - 1.4.6, Removed HTML Output of Controller / License for CitrixCloud and added a proper exclude from check when catalog is excluded GitHub issue #56
# - 1.4.7, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Bugfix added emailCC to XML
#          - Added new functions to test for connectivity: Test-Ping, IsWinRMAccessible, IsWMIAccessible, IsUNCPathAccessible
#            These are fast and efficent for testing base connectivity to the machines before running each test, which helps prevent the functions from waiting for timeouts when connectivity
#            is a problem.
#          - Removed the old Ping function. There wasn't anything wrong with it. I was just attempting to improve its efficiency.
#          - Removed the $wmiOSBlock scriptblock. This is now in a function.
#          - Removed the OS version check using the version of C:\Windows\System32\hal.dll. This is inefficient, slow and can be inaccurate as the hal.dll version loosely correlates to OS builds,
#            but isn't guaranteed to match OS version of edition exactly. Replaced it with the Get-OSVersion function, which also returns the caption that can be useful for other checks.
#          - Added new functions for tests: Get-UpTime, Get-OSVersion, Get-NvidiaDetails, Check-NvidiaLicenseStatus, Get-RDSLicensingDetails, Get-PersonalityInfo, Get-WriteCacheDriveInfo
#            - These will help move the code to functions as per future improvements from version 1.5
#          - Enhanced the following functions to allow for WinRM: CheckCpuUsage, CheckMemoryUsage, CheckHardDiskUsage
#          - Fixed a bug in the CheckHardDiskUsage function so it won't fail if an ISO is left mounted as the D: drive.
#          - Added $OutputLog, $OutputHTML, $ShowStuckSessionsTable and $MaxDisconnectTimeInHours variables to XML file.
#          - Added new Stuck Sessions table and the logic that generates the information so you can quickly see any machines or sessions that may not look healthy.
#            - Some of this logic relies on the $MaxDisconnectTimeInHours variable being set correctly in the parameters XML file.
#            - The only limitation here is that the Get-BrokerSession cmdlet does not have a tags property. So it will not support filtering against the $ExcludedTags array.
#              Thereofre, we use the Get-BrokerMachine cmdlet to filter the $AllStuckSessions collection correctly against the $ExcludedTags array.
#          - Added 3 new script parameters $ParamsFile, $All and $UseRunspace so the script can be launched from a central server to health check multiple sites in parallel.
#          - Added Get-LogicalProcessorCount function for the Runspace pool.
#          - Reordered and improved some of the code so that it flows in and out of the scriptblock.
#          - Improved some checking on the Delivery Controllers so that it exits the scriptblock if invalid.
#          - Changed the $LicWMIQuery variable to $LicQuery and enhanced the code so it can use WinRM.
#          - Implemented Storefront checks from Naveen Arya's version of the script.
#          - Added ShowStorefrontTable and StoreFrontServers variables to XML file to facilitate the Storefront checks.
#          - Added ShowCloudConnectorTable and CloudConnectorServers variables to XML file to facilitate the Cloud Connector checks.
#          - Removed $ControllerVersion -lt 7 code, as this is old and should no longer be needed in this script.
#          - Added ExcludedDeliveryGroups variable to XML file, as a further option to ExcludedCatalogs and ExcludedTags.  Updated the Check Assigments (Delivery Group Check) with this variable.
#            Updated the code so they all support wildcards.
#          - Added more checks to ensure the first element of the ExcludedDeliveryGroups, ExcludedCatalogs and ExcludedTags arrays is not empty. This ensures the script functions if the values
#            in the XML file are left empty.
#          - Improved the Get-BrokerMachine filtering for SessionSupport, so less objects are returned. This reduces processing time for the Where-Object filtering.
#          - It's important to note that Cloud PCs are Azure AD joined only, and may be managed by a different group. So it's most likely that you would not have permissions
#            to run health checks against them. Therefore it's recommended to exclude the Catalogs and/or Delivery Groups from these checks.
#          - The Windows 10/11 Enterprise multi-session hosts are processed as multi-session hosts, but the RDS licensing is not relevant. So the RDS Grace period is marked as N/A
#            and we don't run the Get-RDSLicensingDetails function on those machines.
#          - Added the HypervisorConnection table and removed the HypervisorConnectionstate from the footer. It didn't make sense to have it there.
#          - As the HTML <td> width attribute is deprecated in HTML5 and now considered obsolete, uncommented the CSS from the writeHtmlHeader function and added "width: fit-content" to the td.
#            Have removed the $headerWidths requirement for the writeTableHeader function and from the remainder of the script.
#          - Whilst we can use "white-space: nowrap" to prevent a line-break at hyphen in the HTML, this will make the table too wide. Therefore modified the writeData function to replace
#            hyphens with a non-breaking hyphen before it writes the data to the table.
#          - Removed WriteCacheSize from tables in favor of vhdxSize_inMB. It makes more sense for PVS and MCSIO.
#          - Renamed MCSVDIImageOutOfDate to MCSImageOutOfDate and added this to the XenApp/RDS/Multisession host checks.
#          - Added the MCSIOWriteCacheDrive variable to the XML file to pass to the Get-WriteCacheDriveInfo function in line with PvsWriteCacheDrive.
#          - Renamed the PvsWriteMaxSize variable to WriteCacheMaxSizeInGB in the XML file, and removed the PvsWriteMaxSizeInGB variable altogether, which makes sense for both PVS, when
#            assessing the size of the vdiskdif.vhdx file, and MCSIO, when assessing the size of the mcsdif.vhdx file. This is used against the output of the Get-WriteCacheDriveInfo function.
#          - The output to the "vhdxSize_inGB" column now goes to 3 decimal points, which allows us to see the 4MB default value after a reboot. This way we know it's all good and healthy.
#          - Added the "MasterImageVMDate", "UseFullDiskClone", "UseWriteBackCache", "WriteBackCacheMemSize" columns to the Check Catalog table using the Get-ProvScheme cmdlet.
#          - If the MasterImageVMDate is older than 90 days, mark it as a warning as it's missed patching cycles.
#          - Check the Machine Catalog ProvisioningType that the machine belongs to is MCS before filing out the MCSImageOutOfDate column. This reduces confusion.
#          - Added ShowOnlyErrorXA and ErrorXA back into the logic like ShowOnlyErrorVDI and ErrorVDI.
#          - Added a new table for "ConnectionFailureOnMachine" that uses the Get-BrokerConnectionLog cmdlet.
#          - Added a new variable called BrokerConnectionFailuresinHours to the XML file used by the ConnectionFailureOnMachine check. If the report is run first thing in the morning, this will
#            help see any brokering issues that occurred overnight.
#          - Machine Creation Services (MCS) supports using Azure Ephemeral OS disk for non-persistent VMs. Ephemeral disks should be fast IO because it uses temp storage on the local host.
#            MCSIO is not compatible with Azure Ephemeral Disks, and is therefore not an available configuration.
#          - Added a Get-BrokerTag section to create the $ActualExcludedBrokerTags and $ActualIncludedBrokerTags arrays used for filtering
#          - The script will generate 6 new arrays as it processes the Machine Catalogs, Delivery Groups and Tags: $ActualExcludedCatalogs, $ActualIncludedCatalogs, $ActualExcludedDeliveryGroups,
#            $ActualIncludedDeliveryGroups, $ActualExcludedBrokerTags, $ActualIncludedBrokerTags
#            This helps us filter out the $ExcludedCatalogs, $ExcludedDeliveryGroups, $ExcludedTags by simply using wildcards. This works well if you have naming standards for these items such as
#            "*Remote PC*", "*Cloud PC*", etc.
#          - Added the MaintModeReason info to the MaintMode column for the VDI (singlesession) and XenApp/RDSH (multisession) tests. This is handy if an Administrator has recorded why they have
#            placed a machine into maintenance mode, such as draining the sessions, a change record, etc.
#          - Removed the Get-CitrixMaintenanceInfo function as we are now calling the Get-LogLowLevelOperation cmdlet directly, but still using the methodology the Stefan Beckmann created. Will
#            revisit PS Remoting for sections of the script in a future version. Enhanced it further so we can also get information for who placed Delivery Groups and Hypervisor Connections into
#            maintenance mode. Added PropertyName and TargetType to the output to help with debugging.
#          - Added the $MaintenanceModeActionsFromPastinDays variable to the XML file. It is used for the Get-LogLowLevelOperation cmdlet so we know how far back to go to get maintenance mode
#            actions.
#          - Added the Test-OpenPort function that uses the Test-NetConnection cmdlet to test for Port 80 and 443 to the Delivery Controllers, which is needed for the PowerShell SDK. This
#            replaces the Test-Connection cmdlet, which only sends ICMP echo request packets (pings). Not only is a ping a poor test to validate a Delivery Controller, OT environments are also
#            typically locked down where even pings are not allowed through. This could even be further enhanced to do an XDPing to check the Broker's Registrar service. That function can be
#            found here: https://www.jhouseconsulting.com/2019/05/16/xdping-powershell-function-1931
#          - Added minor updates to the New-XMLVariables function so it's easier to debug.
#          - Added the $ExcludedStuckSessionUserAccounts variable to the XML file. It is used to exclude certain sessions from the Stuck Sessions table. It supports wildcards. The purpose of
#            this is to exclude kiosk accounts that may stay logged in for long periods of time.
#          - Enhanced the Uptime VDI (single-session) and XenApp/RDSH (multi-session) tests so that they are marked as an error when Uptime is $maxUpTimeDays x 3.
# - 1.4.8, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Replaced the CheckCpuUsage function with the Get-CpuConfigAndUsage function. This now gets the total logical processor count, as well as number of sockets and cores per socket,
#            which helps identify misconfigurations. It will flag a warning if there is only 1 logical processor.
#          - Added LogicalProcessors, Sockets, CoresPerSocket to the tests and column outputs.
#          - Added the Get-TotalPhysicalMemory function to get the total physical memory, which helps identify misconfigurations.
#          - Added the TotalPhysicalMemoryinGB to the tests and column outputs. It will flag as warning if less than 4GB and an error is less than 2GB.
#          - Removed the error from the XenApp/RDS (multisession) report for AvgCPU, MemUsg and drive Freespace if there is no WinRM or WMI connectivity. It makes the report look untidy.
#          - For the RDS Licensing tests, LicensingType 5 and LicenseServerList Unknown should be marked as error (red). Improved the documentation for the Get-RDSLicensingDetails function.
#          - Fixed up a bug when the Hypervisor Connection FaultState was marked as ERROR when empty or null.
#          - Tidied up the ETD_MTU test being marked as unknown if the WinRM connection is not possible. It made the report look messy and misleading.
#          - The $displaymode section needs to be re-written so it uses WinRM. However, for the moment is will only run if the machine is accessable via WMI.
#          - The Ping timedout result is no longer marked as an error for the Delivery Controllers, Cloud Connectors and Storefront Servers. This just depends on your security policies and
#            relevant firewall rules for ICMP.
#          - Fixed an error in the logic so that the Cloud Connector tests and table will run when $CitrixCloudCheck variable is 1.
#          - If a VDI (single-session) or XenApp/RDSH (multi-session) machine is not assigned to a Delivery Group, it will be marked as a warning.
#          - Added OSCaption and OSBuild column to the Delivery Controllers, Cloud Connectors and Storefront Servers tests. This helps with OS lifecycle planning.
#          - Added OSCaption column to the VDI (single-session) or XenApp/RDSH (multi-session) tests. This helps with OS lifecycle planning.
#          - Added the ability to sort the VDI (single-session) or XenApp/RDSH (multi-session) tables by a specific header name. For example, you can sort by Delivery Group in ascending order,
#            instead of hostname, which can help to identify config differences within that Delivery Group where hostnames are not contiguous or naming standards have changed over time.
#            - To facilitate this have added...
#              - new $SortDesktopTableByHeaderName and $SortXenAppTableByHeaderName variables to the XML file.
#              - a new function called writeDataSortedByHeaderName that uses the $headerToSortBy parameter that we pass the new variables to. The function contains documentation on how we
#                convert the hashtable into a collection of key-value pairs (DictionaryEntry objects), sort it, and then store it in a new ordered dictionary before creating the output.
#          - Cleaned up the outputs and error capturing as much as possible throughout the script.
#          - Modified the writeTableHeader function to remove td width='6%', which helps better with the layout.
#          - Updated all table widths to help fit all the extra data collected without looking too squashed.
# - 1.4.9, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Uncommented a line in the writeData that I mistakenly left commented for testing under version 1.4.8.
#          - Added an Enabled column for the Delivery Group table as requested under issue #67.
#          - Added a DesktopsPowerStateUnknown column for the Delivery Group table as requested under issue #48.
#            The output of Get-BrokerDesktopGroup cmdlet will not provide this, so we use the Get-BrokerMachine cmdlet with the DesktopGroupName property and further filtering using the
#            PowerState property where it equals 'Unknown' and then piping that to the Measure-Object cmdlet to get the count. Anything above 0 is marked as an ERROR. The issue asked that the
#            PowerState Unknown machines to not count as an available machine. However, that's false logic. Machines that are registered with a PowerState of Unknown will cause issues, because
#            sessions will still attempt to be brokered to them. They do not reduce the available machines.
#          - A PowerState of Unknown for machines in the VDI (single-session) or XenApp/RDSH (multi-session) tables will now be marked as ERROR.
#          - If PowerState is Off, Unregistered and MCSImageOutOfDate are not marked as ERROR. Reduces unnecessary errors (red) in the report.
#          - Tidied up some coding for the write cache drive detection so that it doesn't flag as an error if N/A.
#          - Renamed the LastConnect test to LastConnectionTime in the VDI (single-session) section, so the name makes more sense, and added it to the XenApp/RDSH (multi-session) table too.
#            Changed the warning from 1 month to 30 days, and error from 3 months to 90 days.
#            This aligns with management reporting requirements helps to identify machines that haven't been used in a while that may be able to be decommissioned.
#          - Added a DesktopsNotUsedLast90Days column for the Delivery Group table, which helps understand at a high level if we can decommision some machines from that Delivery Group.
#            The output of Get-BrokerDesktopGroup cmdlet will not provide this, so we use the Get-BrokerMachine cmdlet with the DesktopGroupName property and further filtering using the
#            LastConnectionTime less than '-90' days and then piping that to the Measure-Object cmdlet to get the count. Anything above 0 is marked as a WARNING.
#          - Improved the IsWinRMAccessible function.
#          - Piped the output of the Get-BrokerMachine cmdlet for the VDI (single-session) and XenApp/RDSH (multi-session) tests to Sort-Object via DNSName so that there is an order to the
#            processing and logs for anyone that may be observing it.
# - 1.5.0, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Added a new function called Get-ProfileAndUserEnvironmentManagementServiceStatus, which gets the status of the Citrix UPM, WEM and Microsoft FSLogix services. It checks if the
#            services are installed, running and enabled.
#          - Added the UPMEnabled, FSLogixEnabled and WEMEnabled columns and tests to the tables.
#          - Added the Spooler and CitrixPrint columns and tests to the VDI (single-session) table.
#          - Added an IPv4Address column to all tests to help easily identify which subnets machines are on, which also help locate common firewall and routing issues.
#          - Fixed a bug with the date format for the MasterImageVMDate so that it compares accurately for the 90 day warning.
#          - Found that he Get-ProvScheme MachineProfile and WindowsActivationType properties can return null and throw an error, so wrapped them in a try/catch to manage this.
#          - Changed the way the scriptblock can be executed locally and via Invoke-Command so that we can pass an object as a named parameter in the arguments. This needed to be enhanced
#            so that an object can be passed into the runspace as well.
#          - Improved the order of the headers and grouped them so they can easily be excluded or commented out when not required. An enhancement can be to add more variables to the xml file.
# - 1.5.1, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - If the machine name is ghosted, don't lookup it's IP address. An empty variable passed to the Dns.GetHostAddresses method will lookup the IP Address of the host the script is
#            running on, which is of course incorrect. Also wrapped some further error checking around the Dns.GetHostAddresses method.
#          - Removed the pipe to Select-Object from the Get-Brokerhypervisorconnection cmdlet. This adds unnecessary processing.
#          - Further improved the Citrix UPM, WEM and Microsoft FSLogix outout.
#          - Further improved the Nvidia output and fixed a bug with the $returnLicensable_Product needed to be $return.Licensable_Product.
#          - Improved the Citrix Cloud auth process and output.
#          - Improved some documentation and log outputs.
#          - Added an XDPing test for the Delivery Controller and Cloud Connector checks. It was an obvious test for Delivery Controllers. Even though Rendezvous V2 allows Cloud Connector-less
#            direct VDA registration, standard AD domain joined machines still require Cloud Connectors for VDA registration and session brokering. Even ovcoming this, Cloud Connectors may
#            continue to be required for legacy VDA registration. And therefore should still perform an XDPing to check the health status of the CdsController Iregistrar service.
#          - Changed the EDT MTU test so that it doesn't flag as failed if no data is returned from the "C:\Program Files (x86)\Citrix\HDX\bin\CtxSession.exe" command. That is misleading.
# - 1.5.2, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Updated the Citrix Cloud (DaaS) authentication process allowing for the change in the way the AdminAddress (AKA XDSDKProxy) has been removed as a Global variable and now stored
#            in as a key/value pair under the NonPersistentMetadata property of the Get-XDCredentials cmdlet output.
#          - Updated the Check-NvidiaLicenseStatus to use a switch statement instead of an if/else block. Improved the logic and outputs to help track common issues when licensing fails.
#          - Added the Get-CrowdStrikeServiceStatus function to check and audit CrowdStrike Windows Sensor information.
#          - Added the CSEnabled and CSGroupTags columns to the tables. CSEnabled means that CrowdStrike is installed and running with an Agent ID. CSGroupTags are the Sensor group tags used
#            when installed. This gives us a nice way to audit where it's missing and which tags have been used to ensure we have consistency. This relies on WinRM being enabled across all
#            hosts.
#          - Added the $ShowCrowdStrikeTests variable to the XML file.
#          - Updated the XDPing function.
# - 1.5.3, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Updated the documentation in the IsWinRMAccessible function. This can be used to demonstrate to the Cyber team that WinRM will use Kerberos authentication and is encrypted, despite
#            using the default HTTP port of TCP 5985. The WinRM tests will only proceed if this function returns true.
#          - Further enhanced the XDPing function.
#          - Added an Unassigned column to the MachineCatalogs table, and a test for the AvailableUnassignedCount, which is the number of available machines (those not in any desktop group)
#            that are not assigned to users. It is marked as a warning if great than 0.
#          - Fixed 2 bugs with the Get-CrowdStrikeServiceStatus function.
#          - More coding tidy-ups to provide improved output.
# - 1.5.4, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Improved both the Get-ProfileAndUserEnvironmentManagementServiceStatus and Get-CrowdStrikeServiceStatus functions by making them more resilient.
#          - Fixed a bug with the ToHumanReadable function where it would error if uptime was less than 1 hour.
#          - If the multisession or singlesession host uptime is less than 1 day, output to the log in human readable format. This makes it easier to find how many hours or minutes ago it last
#            rebooted.
# - 1.5.5, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Fixed a bug in the Citrix Cloud Auth related to the NonPersistentMetadata property not existing with older PowerShell modules. Basically just wrapped some more error checking
#            around it using a Try/Catch.
# - 1.5.6, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Added XML variable MarkSpoolerAsWarningOnly. This can be used to flag a failed Spooler and CitrixPrint service as a warning only. It reduces the red (errors) in the reports in
#            environments where these services are not required for normal operation.
#          - For multi-session hosts, if ActiveSessions is greater than ConnectedUsers count, mark as an error. It should point to a ghosted session in the Stuck Sessions Table. Depending on
#            the size of your environment and the potential time difference between when the multi-session host test collects data via the Get-BrokerMachine cmdlet and the stuck sessions
#            collects data via the Get-BrokerSession cmdlet, the ghosted sessions may have cleared, and that discrepancy is no longer valid.
#          - Fixed a bug with the ActiveSessions Count property if the System.String[] (an array of strings) object has been cast into a System.String.
#          - Added XML variable SupportedVDAVersions. Any versions that do not match will be marked as a warning in the VDAVersion column and log. Leave the setting empty in the XML to ignore
#            this. This helps with planning for software lifecycle management and to flag when someone has deployed an unsupported version.
#          - Updated the Citrix Cloud Auth related code to expose the bearer token in a variable which can be used in the header for OData API Authorization.
#          - Improved the Get-PersonalityInfo function to test for more scenarios based on the history of VDA changes for the Personality.ini and MCSPersonality.ini files.
#          - Improved the logging for CrowdStrike output.
#          - Added a minor code improvement to the Get-WriteCacheDriveInfo function.
#          - Exdended the use of the $ErrorVDI and $ErrorXA variables to flag the machine for all warninings and errors.
#          - Added new functions ConvertTo-StructuredData and Write-IetfSyslogEntry for Syslog output.
#          - Added XML variables CheckOutputSyslog, OutputSyslog, PrivateEnterpriseNumber, StructuredDataIDPrefix, SyslogMsgId, SyslogFileOnly, SyslogServer to support Syslog output.
#          - Added new script variables $SyslogAppName, $StructuredDataID and $resultsSyslog to support Syslog output, which are derived from XML variables.
#          - Added the $IsSeverityWarningLevel and $IsSeverityErrorLevel script variables throughout the test sections of the script to support the Severity level for Syslog.
#          - Added the Get-CCSiteId function to get the Citrix Cloud Site Id, which can then be appended to the "cloudsite" name. The SiteName is then added to each test to provide uniqueness
#            across sites when bringing the data altogether into observability platforms.
#          - Added the $AppendCCSiteIdToName XML variable. The Site Name for Citrix Cloud is "cloudxdsite". If you are running checks against multiple Citrix Cloud sites, you may want to set
#            this to true to append the SiteId GUID for uniqueness.
#          - Also added the Get-BearerToken and Get-CCSiteDetails functions in preparation for leveraging the APIs.
#          - Fixed a bug with the getting site maintenance information using the Get-LogLowLevelOperation, where it may not return all records.
#          - Added the $NoProxy and $SkipCertificateCheck XML variables, which make it easier when working with the Invoke-RestMethod and Invoke-WebRequest cmdlets.
#          - Added further variables to the XML params file to support the new Logon Durations script called CitrixLogonDurations.ps1 that I will publish separately. This allows them to share
#            the same XML params files, avoiding duplication.
# - 1.5.7, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Added the Get-CCOrchestrationStatus function in preparation for leveraging the APIs. Using it initially to get the ProductVersion.
#          - Added further variables to the XML params file to support the new Failed Connections script called CitrixFailedConnections.ps1 that I will publish separately. This allows them to
#            share the same XML params files, avoiding duplication.
#          - Added XML variable ShowBrokerConnectionFailuresTable to disable the Connection Failure On Machine table. Using the output from the CitrixFailedConnections.ps1 script provides an
#            improved and thorough output aligned with Citrix Director.
# - 1.5.8, by Jeremy Saunders (jeremy@jhouseconsulting.com)
#          - Enhanced the Get-WriteCacheDriveInfo function by also allowing for the CacheDisk label from BISF.
#          - Fixed a bug with the output of the Get-ProfileAndUserEnvironmentManagementServiceStatus function.
#          - Added the RDSGracePeriodExpired column to the XenApp/RDS/Multisession host table.
#          - Added the Get-RDLicenseGracePeriodEventErrorsSinceBoot function to populate the test for RDSGracePeriodExpired, which is done via the Event Log. This will provide more accuracy
#            than the GetGracePeriodDays method from the Get-RDSLicensingDetails function. If RDSGracePeriodExpired is True, RDSGracePeriod is marked as error to remove confusion.
#          - Enhanced the Check-NvidiaLicenseStatus function to allow for "Platform detection successful" to check if it is licensed in Azure. When using specific Azure N-series VMs you do not
#            need a separate NVIDIA vGPU license server. The necessary licensing for the NVIDIA GRID software is included with the Azure service itself. Microsoft redistributes the Azure-
#            optimized NVIDIA GRID drivers which are pre-licensed for the GRID Virtual GPU Software in the Azure environment. The "Platform detection successful" message indicates that the
#            NVIDIA driver has correctly recognized it is running on a the supported Microsoft Azure virtual machine instance where a license is automatically provide through the platform.
#          - Enhanced the Get-RDLicenseGracePeriodEventErrorsSinceBoot function to wrap the Invoke-Command cmdlet in a Start-Job cmdlet so we can use a timeout. This will help to prevent the
#            Invoke-Command cmdlet getting stuck on unhealthy remote machines.
#          - A minor update to the Get-WriteCacheDriveInfo function for an additional drive label.
#          - Added RecommendedMinimumFunctionalLevel to both the DeliveryGroups and Catalogs tables and enhanced the code under the DeliveryGroups and Catalogs Check to derive the recommended
#            MinimumFunctionalLevel based on VDA Agent Versions in the Delivery Group and Machine Catalog respectively. This required adding the $ProductVersionValues table and
#            $ProductVersionHashTable hashtable that provides a mapping between the Marketing Product Version, Internal Product Version, and the MinimumFunctionalLevel used by the Delivery
#            Groups and Machine Catalogs. Used the Group-BrokerMachine cmdlet to help achieve this.
#            Added the Find-CitrixVersion, Convert-ToComparableVersion, Convert-FunctionalLevelToVersion and Convert-VersionToFunctionalLevel functions so that we can:
#            1) Get the Internal Product Version and MinimumFunctionalLevel based on the Marketing Product Version
#            2) Get the Marketing Product Version and MinimumFunctionalLevel based on the Internal Product Version
#            3) Get the Lowest Supported VDA Version based on the MinimumFunctionalLevel and the Highest VDA Version Before the Next FunctionalLevel
#            4) Convert the MinimumFunctionalLevel to a version for comparison, and back again for correct output.
#          - Added MachinesInMaintMode and PercentageOfMachinesInMaintMode to the DeliveryGroups table and enhanced the code under the DeliveryGroups Check to get the count and percentage of
#            machines in maintenance mode and mark as warning if greater than 0% and error if 50% or greater. This will help alert Admins to potential issues when accidentally placing too many
#            machines into maintenance mode, or simply forgetting to remove maintenance mode after a maintenance task and/or a reboot. Used the Group-BrokerMachine cmdlet to help achieve this.
#          - If using an on-prem License server and the output for the license table contains no data, it will be excluded from the HTML report. This may be because it is License Activation
#            Service (LAS) enabled. With LAS enabled we now have unlimited capacity and no way to get a license count. "LAS provides products with unlimited capacity for the duration of the
#            entitlement validity. Because LAS does not require a 'check-out' operation to track individual licenses, real-time usage metrics are not available within the License Server or
#            Citrix Cloud."
#            A reference to the "Unlimited capacity" point here: https://docs.citrix.com/en-us/licensing/current-release/license-activation-service.html
#          - Added further variables to the XML params file to support the new Insights script called CitrixInsights.ps1 that I will publish separately. This allows them to share the same XML
#            params files, avoiding duplication.
#
# ==CURRENT KNOWN ISSUES AND/OR LIMITATIONS ==
# - Any functions that use the Invoke-Command cmdlet "may" cause the script to wait indefinitely when run against an unhealthy machine. This is due to the known timeout issue with this cmdlet.
#   Refer to comments in future release for how to code around this.
#
# == FUTURE ==
# #  1.5.x
# Version changes by S.Thomet & J.Saunders
# - CREATE Proper functions
# - Change all functions to allow for Invoke-Command for PS Remoting where possible.
# - The Invoke-Command cmdlet doesn't have a -Timeout parameter. To force a timeout for the Invoke-Command cmdlet we can put it in a ScriptBlock and run it as background job using Start-Job.
#   Then use Wait-Job on it with -Timeout specified. It will wait the amount of time we specify and then terminate the job. Refer to the Get-RDLicenseGracePeriodEventErrorsSinceBoot function
#   to see how this has been implemented.
# - Look to merge the Singlesession and Multisession sections using a for loop to process, as there is quite a bit of code duplication in those two sections.
# - Look to merge the Delivery Controllers, Cloud Connectors and Storefront Servers sections using a for loop to process, as there is quite a bit of code duplication in those three sections.
# - Enhance the MCSIO tests, where possible.
#   It is standard practice in Citrix Consulting to set MCSIO drive cache to be the same size as the C drive of the image. This reduces the risk of introducing failures that can impact the
#   business. They do this to error on the side of safety and reduce risk. Matching the disk cache size to image drive size ensures that isn't a problem. We could therefore do a further test
#   for this, by comparing the size of the drives.
# - Consider adding the Get-BrokerMachine LastAssignmentTime property as a column, which was introduced from 2209.
# - $displaymode section needs to be re-written so it uses WinRM.
# - Combine functions where possible to collect data more efficiently.
# - Implement DaaS APIs (from 2209 and above) to remove reliance on PowerShell SDK. It should be able to be slotted in quite easily with a variable to switch between them, which will allow
#   new and legacy to work side-by-side.
# - If Syslog output is not enough, then further convert the output for ingestion into tools like Splunk, Azure Monitor Logs, OpenSearch, Elasticsearch, Cribl, etc.
# - Implement improved licensing checks, especially for Citrix Cloud licensing, keep this as a separate script.
#
# #  1.5.x
# Version changes by S.Thomet
# - Implement Idea #27 from GitHub: Fail-Rate in %
#
#=========== History END ===========================================================================
