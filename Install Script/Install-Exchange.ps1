#requires -version 2.0

<#
    .SYNOPSIS
        Installs Exchange 2010 and configures to DoD EE specifications post-install.
    .DESCRIPTION
        This script is used to install Exchange 2010 in the DEE 1.0 environment.

        THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
        KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
        IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
        PARTICULAR PURPOSE.
        
        Author: James E. Miller
        Version: 2.2.20130828 
    .PARAMETER SourceDir
        Location of the Exchange source files in the installation package.
    .PARAMETER ScriptsDir
        Location of the Scripts folder in the installation package which contains
        supplementary installation files.
    .PARAMETER PrereqsDir
        Location of prerequisite packages for the install like the Office Filter Pack.
    .PARAMETER TargetDir
        Target location of the Exchange install.
    .PARAMETER Recover
        Puts the script into recovery mode instead of performing a fresh install.
    .EXAMPLE
        .\Install-Exchange.ps1

        This uses the script in its default install mode which will configure Exchange on each
        server in the manifest file.
    .EXAMPLE
        .\Install-Exchange.ps1 -SourceDir "Z:\Exchange2010SP3"

        This uses the script in its default install mode but overrides the default location of
        the source installation files.
    .EXAMPLE
        .\Install-Exchange.ps1 -Recover

        This tells the script to go into recovery mode which will reinstall Exchange on servers
        in a disaster recovery situation.
#>

param( 
    #Folder location of the Exchange source files in the installation package.
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$SourceDir = "D:\software\Exchange2010SP3",
    
    #Location of the Scripts folder in the installation package.
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$ScriptsDir = "D:\software\scripts",
    
    #Location of prerequisite packages for Exchange 2010 install.
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$PrereqsDir = "D:\software\PreReqPatches2008R2",
    
    #Target location of the Exchange install.
    [string]$TargetDir = "M:\program files\exchsrvr",

    #Location of the wildcard certificate.
    [ValidateScript({Test-Path $_ -PathType Leaf~})]
    [string]$WildcardCert = "Z:\Software\Exchange\Wildcard Certificate\WildcardCert\WildcardCert.pfx",
    
    #Sets script to recover mode.
    [switch]$Recover
)

#region Global Constants
#DO NOT CHANGE ANYTHING IN THIS SECTION
$MACHINE = Get-WMIObject "Win32_ComputerSystem"
$MACHINE_NAME = $MACHINE.Name
$MACHINE_FQDN = $MACHINE_NAME + "." + $MACHINE.Domain
$DOMAIN = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
$AD_SITE = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()
$SITE_NAME = $AD_SITE.Name
$GC_SERVER = ($DOMAIN.DomainControllers | ? { $_.SiteName -eq $AD_SITE } | Get-Random).Name
$ErrorActionPreference = "Continue"
#endregion

#region Install Constants
#Folder location for the logging of this script
$INSTALL_LOG_DIR = "D:\ExchAutoInstall"

#Signed certificate request file folder
$CERT_FILE = $INSTALL_LOG_DIR + "\$MACHINE_NAME-cert.req"

#Exchange 2010 Product Keys
$ENT_PROD_KEY = "GVMTV-GMXWH-C234M-8FMWP-TFPFP"
$STD_PROD_KEY = "XJG6B-4D4YV-4M338-Q42H6-39VT2"
#endregion
 
#region Client Access Constants
#DO NOT CHANGE ANYTHING IN THIS SECTION
$CAS_OWA_VDIR = "$MACHINE_NAME\OWA (Default Web Site)"
$CAS_ECP_VDIR = "$MACHINE_NAME\ECP (Default Web Site)"
$CAS_EAD_VDIR = "$MACHINE_NAME\Autodiscover (Default Web Site)"
$CAS_EAS_VDIR = "$MACHINE_NAME\Microsoft-Server-ActiveSync (Default Web Site)"
$CAS_EWS_VDIR = "$MACHINE_NAME\EWS (Default Web Site)"
$CAS_OAB_VDIR = "$MACHINE_NAME\OAB (Default Web Site)"
$CAS_RPC_PORT = 59531 # Internal RPC port for LB config
$CAS_AB_PORT = "59532" # Internal AddressBook port for LB config
$CAS_EXT_URL = "web.mail.mil" # Common external OWA URL
$CAS_EAD_URL = "autodiscover.mail.mil" 
$CAS_MRS_MAX_MOVES_PER_DB = "50" # MRS - maxActiveMovesPerTargetMDB"
$CAS_MRS_MAX_MOVES_PER_SERVER = "50" # MRS - maxActiveMovesPerTargetServer"

#endregion

#region Hub Transport Constants
$HUB_LOG_BASE_DIR = "D:\syslogs\TransportLogs" # Log Base Directory
$HUB_CONN_LOG_DIR = $HUB_LOG_BASE_DIR + "\Connect" # Connectivity Log Directory
$HUB_TRACK_LOG_DIR = $HUB_LOG_BASE_DIR + "\Track" # Message Tracking Log Directory
$HUB_PIPE_LOG_DIR = $HUB_LOG_BASE_DIR + "\Pipeline" # Pipeline Tracing Log Directory
$HUB_RCV_LOG_DIR = $HUB_LOG_BASE_DIR + "\Receive" # Receive Protocol Log Directory
$HUB_SND_LOG_DIR = $HUB_LOG_BASE_DIR + "\Send" # Send Protocol Log Director
$HUB_ROUTE_LOG_DIR = $HUB_LOG_BASE_DIR + "\Route" # Routing Log Directory
$HUB_QUEUE_DATADIR = "D:\QueueData" # Queue Database Directory
$HUB_QUEUE_LOGSDIR = "D:\QueueLogs" # Queue Log Directory
$HUB_LOGFILE_MAX_SIZE = "10MB" # Max Size per Protocol Log File
$HUB_LOGDIR_MAX_SIZE = "5GB" # Max Size Protocol Log Directory
$HUB_LOG_MAX_AGE = "7.00:00:00" # Max Age per Protocol Log File
$POSTMASTER = "postmaster@mail.mil" # Postmaster address
$MSG_TRK_LOG_AGE = "60.00:00:00" # Max age of tracking log file
$MSG_TRK_LOG_DIR_SIZE = "30GB" # Max size of tracking log directory FIFO
$MSG_TRK_LOG_FILE_SIZE = "500MB" # Max size of single tracking log
$MSG_EXPIRE_AGE = "2.00:00:00" # Max duration in queue
$MSG_MAX_SIZE = "28MB" # Max message size
$DUMPSTER_SIZE = "42MB" # Transport Dumpster Max Size

#Collect the GUID of the Production network adapter
[string]$NET_GUID = Get-WMIObject "Win32_NetworkAdapter" | ? { $_.NetConnectionId -like "*production*"} | Select-Object -Property GUID
$NET_GUID = $NET_GUID.Replace("@{GUID={","").Replace("}}","")

#endregion

#region Edge Transport Constants
$EDGE_SUB_FILE = $INSTALL_LOG_DIR + "\$MACHINE_NAME-edge_subscription.xml" #Name and location of Edge subscription file
$QUARANTINE_EMAIL = "svc.ex.quarantine@mail.mil" # SMTP address of internal quarantine mailbox
#endregion

#region IIS Constants
#IIS Logging Directory and Fields to log
$IIS_LOG_DIR = "D:\syslogs\IISLogs" 
$IIS_LOG_FIELDS = "Date,Time,ClientIP,UserName,ServerIP,ServerPort,Method,UriStem,UriQuery,HttpStatus,HttpSubStatus,Win32Status,TimeTaken,UserAgent,Referer,BytesSent,BytesRecv,ProtocolVersion"
#endregion

#region Site-SpecificConstants
#Matches the computer AD site to the info in sitelist.csv and assigns proper FQDNS to client access urls and SMTP FQDNs.
if ((Test-Path $ScriptsDir\sitelist.csv ) -eq $false) {
	Write-Host "Cannot find sitelist.csv in the proper location.`n`nPress any key to exit..." -ForegroundColor Red
	$Host.UI.RawUI.ReadKey() | Out-Null
	exit
}
else {
    $siteListFile = Import-CSV $ScriptsDir\sitelist.csv

    $siteInfo = $siteListFile | ? { $_.siteName.toLower() -eq $SITE_NAME.toLower() }

    if ($siteInfo.siteName) {
	    $INTERNAL_NAME = $siteInfo.InternalName
	    $EXT_OA_NAME = $siteInfo.ExternalOAName
	    $EXT_OWA_NAME = $siteInfo.ExternalOWAName
	    $INT_APPS_NAME = $siteInfo.AppsInternalName
	    $EXT_APPS_OA_NAME = $siteInfo.AppsExternalOAName
	    $EXT_APPS_OWA_NAME = $siteInfo.AppsExternalOWAName	
	    $SMTP_NAME = $siteInfo.smtpName
        
        $edgesite = ($siteInfo.siteName.Substring(1, $siteInfo.siteName.length - 1)).ToLower()
        $EDGE_NAME = "edge-$edgesite.mail.mil"
    } 
    else {
	    Write-Host "This computer's site membership does not match any of the sites named in the sitelist.csv.`nPlease check the site name and sitelist.csv.`n`nPress any key to exit..." -foregroundcolor Red
	    $Host.UI.RawUI.ReadKey() | Out-Null
	    exit
    }
}

#endregion

#region Script Functions
#Installs the FilterPack Prerequisite
function InstallFilterPack {
	Push-Location
	#Install Microsoft Filter Pack
	Write-Host "Installing Microsoft Filter Pack..."
	$expression = "$PrereqsDir\FilterPack64bit.exe /quiet /norestart"
	Invoke-Expression $expression
	Start-Sleep -Seconds 10
	Write-Host "Filter Pack installation done!" -ForegroundColor Green
	Pop-Location
}

#Loads Exchange Management Shell into the console.  Used for post-installation Exchange configuration.
function Connect-Exchange {
	. $TargetDir\bin\RemoteExchange.ps1
	Connect-ExchangeServer -serverFQDN $MACHINE_NAME
	Start-Sleep 30
}

# Add the Telnet client to this server role if it is missing.  The feature is not required for Exchange but is useful during troubleshooting.
function Install-TelnetClient {
	Import-Module ServerManager
	$tc = Get-WindowsFeature telnet-client
	if ($tc.Installed -eq $false) {
		Add-WindowsFeature telnet-client
	}
}

#Disable IE on the machine if it is enabled
function Disable-InternetExplorer {
	$ie = dism /online /get-featureinfo /featurename:Internet-Explorer-Optional-amd64
	if ($ie -contains "State : Enabled") {
		dism /online /disable-feature /featurename:Internet-Explorer-Optional-amd64 /NoRestart
	}
}

function Complete-Install {
    #Reboot computer
    Write-Host "Install Complete, restarting..."
    Stop-Transcript
	Restart-Computer -Confirm:$false -Force
}

#endregion

#region Install Functions
function Install-Exchange2010 {
	if ((Test-Path $ScriptsDir\manifest.csv ) -eq $false) { 
		Write-Host "Cannot find manifest.csv in the proper location.`n`nPress any key to exit..." -ForegroundColor Red
		$Host.UI.RawUI.ReadKey() | Out-Null
		exit
	}
    else {
	    $manifestFile = Import-CSV $ScriptsDir\manifest.csv
	    $server = $manifestFile | ? { $_.servername.toLower() -eq $MACHINE_NAME.toLower() }
	    if ($server) {
		    $role = $server.role
		    switch ($role) {
			    "CH" {Install-CASHTS}
                "HT" {Install-HTS}
			    "MB" {Install-MBX}
			    "ET" {Install-ETS}
			    "EDSP" {Install-CASNoCAC}
                "AIO" {Install-AIO}
			    "MT" {Install-MGMT}
			    default {
				    Write-Host "$role is not a valid role code for Exchange installation.`n`nPress any key to exit..." -ForegroundColor Red
				    $Host.UI.RawUI.ReadKey() | Out-Null
			    }
		    }
	    }
	    else {
		    Write-Host "$MACHINE_NAME does not appear in the manifest file.`n`nPress any key to exit..." -ForegroundColor Red
		    $Host.UI.RawUI.ReadKey() | Out-Null
	    }
    }
}

#Installs and configures the CAS/HTS roles
function Install-CASHTS {
	Install-TelnetClient
    Disable-InternetExplorer

	Push-Location

	#Install Exchange 2010 with CAS and Hub roles.
	Set-Location $SourceDir
    if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
    else {
	    ./setup.com `/m:Install `/r:C`,H `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
	Pop-Location
	. Connect-Exchange
	
	#License the Exchange server.
	Set-ExchangeServer -Identity $MACHINE_NAME -ProductKey $STD_PROD_KEY -Confirm:$false

	#Generate a Certificate Request file if fresh install
    if (!($Recover)) {
	    $certData = New-ExchangeCertificate -GenerateRequest -SubjectName "CN=$MACHINE_FQDN, OU=DISA, OU=PKI, OU=DoD, O=U.S. Government, C=US" -DomainName $MACHINE_FQDN, $MACHINE_NAME, web.mail.mil, *.easf.csd.disa.mil, *.mail.mil -PrivateKeyExportable $true
	    Set-Content -Path $CERT_FILE -Value $certData
    }

	Configure-HTS
	Configure-CAS
	Configure-CACAuthForOA
	Configure-CommonIIS
	Complete-Install
}

#Installs and configures the Hub Transport role
function Install-HTS {
	Install-TelnetClient
    Disable-InternetExplorer

	Push-Location

	#Install Exchange 2010 with the Hub role.
	Set-Location $SourceDir
	if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
    else {
        ./setup.com `/m:Install `/r:H `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
	Pop-Location
	. Connect-Exchange
	
	#License the Exchange server.
	Set-ExchangeServer -Identity $MACHINE_NAME -ProductKey $STD_PROD_KEY -Confirm:$false

	#Generate a Certificate Request file if fresh install
    if (!($Recover)) {
        $certData = New-ExchangeCertificate -GenerateRequest -SubjectName "CN=$MACHINE_FQDN, OU=DISA, OU=PKI, OU=DoD, O=U.S. Government, C=US" -DomainName $MACHINE_FQDN, $MACHINE_NAME -PrivateKeyExportable $true
	    Set-Content -Path $CERT_FILE -Value $certData
    }

	Configure-HTS
	Configure-CommonIIS
	Complete-Install
}

function Install-CASNoCAC {
	Install-TelnetClient
    Disable-InternetExplorer

	Push-Location

	#Install Exchange 2010 with the CAS role.
	Set-Location $SourceDir
	if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/dc:$GC_SERVER
    }
    else {
        ./setup.com `/m:Install `/r:C `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/dc:$GC_SERVER
    }
	Pop-Location
	. Connect-Exchange
	
	#License the Exchange server.
	Set-ExchangeServer -Identity $MACHINE_NAME -ProductKey $STD_PROD_KEY -Confirm:$false

	Configure-CAS
	Configure-CommonIIS
	Complete-Install
}

#Installs and configures the Mailbox role
function Install-MBX {
    Disable-InternetExplorer

    Push-Location

	#Install Exchange 2010 with the MBX role.
	Set-Location $SourceDir
	if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/dc:$GC_SERVER
    }
    else {
        ./setup.com `/m:Install `/r:M `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/dc:$GC_SERVER
    }
	Pop-Location
	. Connect-Exchange

	#License the Exchange server.
	Set-ExchangeServer -Identity $MACHINE_NAME -ProductKey $ENT_PROD_KEY -Confirm:$false

	Configure-MBX
	Configure-CommonIIS
	Complete-Install
}

function Install-AIO {
	Install-TelnetClient
    Disable-InternetExplorer

	Push-Location

	#Install Exchange 2010 with CAS, Hub and MBX roles.
	Set-Location $SourceDir
    if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
    else {
	    ./setup.com `/m:Install `/r:C`,H`,M `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates" `/DoNotStartTransport `/dc:$GC_SERVER
    }
	Pop-Location
	. Connect-Exchange
	
	#License the Exchange server.
	Set-ExchangeServer -Identity $MACHINE_NAME -ProductKey $ENT_PROD_KEY -Confirm:$false

	#Generate a Certificate Request file if fresh install
    if (!($Recover)) {
	    $certData = New-ExchangeCertificate -GenerateRequest -SubjectName "CN=$MACHINE_FQDN, OU=DISA, OU=PKI, OU=DoD, O=U.S. Government, C=US" -DomainName $MACHINE_FQDN, $MACHINE_NAME, web.mail.mil, *.easf.csd.disa.mil, *.mail.mil -PrivateKeyExportable $true
	    Set-Content -Path $CERT_FILE -Value $certData
    }

	Configure-HTS
	Configure-CAS
	Configure-CACAuthForOA
    Configure-MBX
	Configure-CommonIIS
	Complete-Install
}

#Installs and configures the Edge Transport role
function Install-ETS {
    Install-TelnetClient
    Disable-InternetExplorer

	Push-Location

	#Install Exchange 2010 with the Edge Transport role.
	Set-Location $SourceDir
	if ($Recover) {
        ./setup.com `/m:RecoverServer `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates"
    }
    else {
        ./setup.com `/m:Install `/r:E `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates"
    }
    Pop-Location
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
		
	#License the Exchange server.
	Set-ExchangeServer -ProductKey $STD_PROD_KEY -Confirm:$false

	#Generate a Certificate Request file if fresh install
    if (!($Recover)) {
        $certData = New-ExchangeCertificate -GenerateRequest -SubjectName "CN=$MACHINE_FQDN, OU=DISA, OU=PKI, OU=DoD, O=U.S. Government, C=US" -DomainName $MACHINE_FQDN, $MACHINE_NAME, $EDGE_NAME, $SMTP_NAME, mail.mil -PrivateKeyExportable $true
	    Set-Content -Path $CERT_FILE -Value $certData
    }

    Configure-ETS
    Complete-Install
}

#Installs the Exchange Management tools
function Install-MGMT {
    Set-Location $SourceDir
	./setup.com `/m:Install `/r:MT `/InstallWindowsComponents `/t:$TargetDir `/u:"$SourceDir\updates"
}

#endregion

#region Role Configuration Functions
#Configures the Mailbox role
function Configure-MBX {
	#Grab the CAS Array assigned to the site
	$casArray = (Get-ClientAccessArray -site $SITE_NAME).Name

	#Disable Subject logging on message tracking.  Rename the default database and set quotas.
	Set-MailboxServer -Identity $MACHINE_NAME -MessageTrackingLogSubjectLoggingEnabled $false -AutoDatabaseMountDial BestAvailability -DatabaseCopyAutoActivationPolicy IntrasiteOnly 
	Get-MailboxDatabase -Server $MACHINE_NAME | Set-MailboxDatabase -Name "Health-Admin-DB-$MACHINE_NAME" -ProhibitSendQuota 2GB -IssueWarningQuota 1.9GB -ProhibitSendReceiveQuota 2.3GB -Confirm:$false -RpcClientAccessServer $casArray -RetainDeletedItemsUntilBackup $false -DomainController $GC_SERVER
}

#Configures the Hub Transport role
function Configure-HTS {
	Configure-TransportSettings
	
	#Configure the Hub Server with various settings
	Set-TransportServer -Identity $MACHINE_NAME -ConnectivityLogEnabled $true -ConnectivityLogMaxAge $HUB_LOG_MAX_AGE -ConnectivityLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -ConnectivityLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -ConnectivityLogPath $HUB_CONN_LOG_DIR -MessageTrackingLogEnabled $true -MessageTrackingLogMaxAge $MSG_TRK_LOG_AGE -MessageTrackingLogMaxDirectorySize $MSG_TRK_LOG_DIR_SIZE -MessageTrackingLogMaxFileSize $MSG_TRK_LOG_FILE_SIZE -MessageTrackingLogPath $HUB_TRACK_LOG_DIR -PipelineTracingPath $HUB_PIPE_LOG_DIR -ReceiveProtocolLogMaxAge $HUB_LOG_MAX_AGE -ReceiveProtocolLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -ReceiveProtocolLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -ReceiveProtocolLogPath $HUB_RCV_LOG_DIR -RoutingTableLogMaxAge $HUB_LOG_MAX_AGE -RoutingTableLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -RoutingTableLogPath $HUB_ROUTE_LOG_DIR -SendProtocolLogMaxAge $HUB_LOG_MAX_AGE -SendProtocolLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -SendProtocolLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -SendProtocolLogPath $HUB_SND_LOG_DIR -MessageExpirationTimeout $MSG_EXPIRE_AGE -ExternalDNSAdapterGUID $NET_GUID -InternalDNSAdapterGUID $NET_GUID -DomainController $GC_SERVER -Confirm:$false | Out-Null

	#Remove the default client receive connector
	Get-ReceiveConnector "$MACHINE_NAME\Client $MACHINE_NAME" | Remove-ReceiveConnector -Confirm:$false -DomainController $GC_SERVER

	#Update all receive connectors with Max message size, protocol logging, and STIG-required banner
	Get-ReceiveConnector -Server $MACHINE_NAME | Set-ReceiveConnector -Banner "220 SMTP Server Ready" -MaxMessageSize $MSG_MAX_SIZE -ProtocolLoggingLevel Verbose -DomainController $GC_SERVER

	#Add server to Internal Relay Send Connector
	$internalRelaySendConnector = Get-SendConnector ("Internal Send Connector $SITE_NAME") -ErrorAction SilentlyContinue
	if ($internalRelaySendConnector) {
		Set-SendConnector $internalRelaySendConnector -SourceTransportServers ($internalRelaySendConnector.SourceTransportServers + $MACHINE_NAME) -domaincontroller $GC_SERVER
	} else {
		New-SendConnector -Name "Internal Send Connector $SITE_NAME" -Enabled:$true -Usage 'Custom' -AddressSpaces 'SMTP:us.army.mil;10' -IsScopedConnector $false -DNSRoutingEnabled $false -UseExternalDNSServersEnabled $false -ProtocolLoggingLevel Verbose -MaxMessageSize 20480KB -SmartHosts "143.69.243.34", "143.69.251.34" -SourceTransportServers $MACHINE_NAME -domaincontroller $GC_SERVER
	}
}

#Configures the Edge Transport role
function Configure-ETS {
	Configure-TransportSettings

    #Generate an Edge Subscription file.  
	New-EdgeSubscription -Filename $EDGE_SUB_FILE -Confirm:$false -force

	#Configure the Edge Server with various settings
	Set-TransportServer -ConnectivityLogEnabled $true -ConnectivityLogMaxAge $HUB_LOG_MAX_AGE -ConnectivityLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -ConnectivityLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -ConnectivityLogPath $HUB_CONN_LOG_DIR -MessageTrackingLogEnabled $true -MessageTrackingLogMaxAge $MSG_TRK_LOG_AGE -MessageTrackingLogMaxDirectorySize $MSG_TRK_LOG_DIR_SIZE -MessageTrackingLogMaxFileSize $MSG_TRK_LOG_FILE_SIZE -MessageTrackingLogPath $HUB_TRACK_LOG_DIR -MessageTrackingLogSubjectLoggingEnabled $false -PipelineTracingPath $HUB_PIPE_LOG_DIR -ReceiveProtocolLogMaxAge $HUB_LOG_MAX_AGE -ReceiveProtocolLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -ReceiveProtocolLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -ReceiveProtocolLogPath $HUB_RCV_LOG_DIR -RoutingTableLogMaxAge $HUB_LOG_MAX_AGE -RoutingTableLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -RoutingTableLogPath $HUB_ROUTE_LOG_DIR -SendProtocolLogMaxAge $HUB_LOG_MAX_AGE -SendProtocolLogMaxDirectorySize $HUB_LOGDIR_MAX_SIZE -SendProtocolLogMaxFileSize $HUB_LOGFILE_MAX_SIZE -SendProtocolLogPath $HUB_SND_LOG_DIR -MessageExpirationTimeout $MSG_EXPIRE_AGE -ExternalDNSAdapterGUID $NET_GUID -InternalDNSAdapterGUID $NET_GUID -Confirm:$false | Out-Null

	#Update the receive connector with Max message size, protocol logging, and STIG-required banner
	Get-ReceiveConnector | Set-ReceiveConnector -Banner "220 SMTP Server Ready" -MaxMessageSize $MSG_MAX_SIZE -ProtocolLoggingLevel Verbose

    #Create a Transport Rule that redirects messages with Blank Servers to a quarantine mailbox
	New-TransportRule "Quarantine Messages From Blank Senders" -FromAddressMatchesPattern "^$" -RedirectMessageTo $QUARANTINE_EMAIL -Enabled $true -Comments "Quarantines messages with blank senders by redirecting them to a quarantine mailbox per STIG requirements."

    #Check for and remove IIS if installed
    Import-Module ServerManager

    $iisRole = Get-WindowsFeature Web-Server
	if ($iisRole.Installed -eq $true) {
		Remove-WindowsFeature Web-Server
	}
}

#Configures common transport settings for ETS and HTS roles
function Configure-TransportSettings {
	# Check for logging directories and create as needed.
	if((Test-Path $HUB_LOG_BASE_DIR) -eq $false) {
		Write-Host "       Creating $HUB_LOG_BASE_DIR"
		New-Item $HUB_LOG_BASE_DIR -ItemType directory
	}
	if((Test-Path $HUB_CONN_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_CONN_LOG_DIR"
		New-Item $HUB_CONN_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_TRACK_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_TRACK_LOG_DIR"
		New-Item $HUB_TRACK_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_PIPE_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_PIPE_LOG_DIR"
		New-Item $HUB_PIPE_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_RCV_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_RCV_LOG_DIR"
		New-Item $HUB_RCV_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_SND_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_SND_LOG_DIR"
		New-Item $HUB_SND_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_ROUTE_LOG_DIR) -eq $false) {
		Write-Host "       Creating $HUB_ROUTE_LOG_DIR"
		New-Item $HUB_ROUTE_LOG_DIR -ItemType directory
	}
	if((Test-Path $HUB_QUEUE_DATADIR) -eq $false) {
		Write-Host "       Creating $HUB_QUEUE_DATADIR"
		New-Item $HUB_QUEUE_DATADIR -ItemType directory
	}
	if((Test-Path $HUB_QUEUE_LOGSDIR) -eq $false) {
		Write-Host "       Creating $HUB_QUEUE_LOGSDIR"
		New-Item $HUB_QUEUE_LOGSDIR -ItemType directory
	}


	#Update the edgetransport.exe.config file to move the queue DB and log paths.
	$configFilePath = $TargetDir + "\bin\edgetransport.exe.config"
	$xml = [XML](Get-Content $configFilePath)
	$queueDB = $xml.configuration.SelectSingleNode("appSettings/add[@key='QueueDatabasePath']")
	$queueLog = $xml.configuration.SelectSingleNode("appSettings/add[@key='QueueDatabaseLoggingPath']")
	$queueDB.Value = $HUB_QUEUE_DATADIR
	$queueLog.Value = $HUB_QUEUE_LOGSDIR
	$xml.Save($configFilePath)
}

function Configure-CAS {
    Import-Module ServerManager
    Import-Module WebAdministration
	    
    #Add File Server role for file witness
    $fileServer = Get-WindowsFeature FS-FileServer
	if ($fileServer.installed -eq $false) {
		Add-WindowsFeature FS-FileServer
	}
    
    #CAS Configurations for virtual directories
	Set-OwaVirtualDirectory -Identity $CAS_OWA_VDIR -ExternalAuthenticationMethods NTLM -BasicAuthentication $false -FormsAuthentication $false -WindowsAuthentication $true -ExternalUrl "https://$EXT_OWA_NAME/owa" -ChangePasswordEnabled $false -DomainController $GC_SERVER -Confirm:$false
	Set-EcpVirtualDirectory -Identity $CAS_ECP_VDIR -ExternalAuthenticationMethods NTLM -BasicAuthentication $false -FormsAuthentication $false -WindowsAuthentication $true -ExternalUrl "https://$EXT_OWA_NAME/ecp" -DomainController $GC_SERVER -confirm:$false
	Set-ActiveSyncVirtualDirectory -Identity $CAS_EAS_VDIR -BasicAuthEnabled:$false -WindowsAuthEnabled:$true -InternalUrl "https://$INTERNAL_NAME/Microsoft-Server-ActiveSync" -ExternalUrl "https://$EXT_OWA_NAME/Microsoft-Server-ActiveSync" -DomainController $GC_SERVER -Confirm:$false
	Set-OABVirtualDirectory -Identity $CAS_OAB_VDIR -InternalUrl "https://$INTERNAL_NAME/oab" -ExternalUrl "https://$EXT_OA_NAME/oab"  -DomainController $GC_SERVER -Confirm:$false
	Set-WebServicesVirtualDirectory -Identity $CAS_EWS_VDIR -WindowsAuthentication $true -InternalUrl "https://$INTERNAL_NAME/ews/exchange.asmx" -ExternalUrl "https://$EXT_OA_NAME/ews/exchange.asmx" -DomainController $GC_SERVER -Confirm:$false
	set-AutodiscoverVirtualDirectory -Identity $CAS_EAD_VDIR -BasicAuthentication $false -WindowsAuthentication $true -DomainController $GC_SERVER -Confirm:$false
	Set-ClientAccessServer -Identity $MACHINE_NAME -AutoDiscoverServiceInternalUri "https://$INTERNAL_NAME/autodiscover/autodiscover.xml" -DomainController $GC_SERVER -Confirm:$false
	
	#Enable Outlook Anywhere
	Enable-OutlookAnywhere -Server $MACHINE_NAME -ExternalHostName $EXT_OA_NAME -SSLOffloading:$false -DefaultAuthenticationMethod NTLM -DomainController $GC_SERVER -Confirm:$false

	#Set static ports for RCA and AddressBook to facilitate load balancer configuration
	New-Item -Path HKLM:\SYSTEM\CurrentControlSet\services\MSExchangeRPC\ParametersSystem
	New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\MSExchangeRPC\ParametersSystem -Name "TCP/IP Port" -PropertyType DWORD -Value $CAS_RPC_PORT
	New-Item -Path HKLM:\SYSTEM\CurrentControlSet\services\MSExchangeAB\Parameters
	New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\MSExchangeAB\Parameters -Name "RpcTcpPort" -PropertyType string -Value $CAS_AB_PORT

	#Modifications to the OWA web.config file to enable ToU cookie checking 
	$xml = New-Object XML
	$xml.PreserveWhitespace = $true
    $xml.Load("$TargetDir\ClientAccess\Owa\web.config")
	$oldappSettings = @($xml.configuration.appsettings.add)[0]
	$oldmodules = @($xml.configuration.location."system.webServer".modules.add)[0]

	if (!($xml.SelectSingleNode('/configuration/location/system.webServer/modules/add[@type="CheckForCookie"]'))) {
		$newmodules = $oldmodules.clone()
		$nmcomment = $xml.CreateComment("TOU Cookie Check Enable - Ensure this module is at the top of any other modules in the section")
		$newmodules.name = "CheckForCookie"
		$newmodules.type = "CheckForCookie"
		$xml.configuration.location."system.webServer".modules.PrependChild($newmodules)
		$xml.configuration.location."system.webServer".modules.PrependChild($nmcomment)
	} 

	$xml.save("$TargetDir\ClientAccess\Owa\web.config")
	

	#Modifications to the MRS config file to enable more mailbox imports simultaneously
	$xml = New-Object XML
	$xml.PreserveWhitespace = $true
    $xml.Load("$TargetDir\Bin\MSExchangeMailboxReplication.exe.config")
	$oldMRSSettings = @($xml.Configuration.MRSConfiguration)[0]
	
	$newMRSSettings = $oldMRSSettings.clone()
	
	$newMRSSettings.SetAttribute("MaxActiveMovesPerTargetMDB",$CAS_MRS_MAX_MOVES_PER_DB)
	$newMRSSettings.SetAttribute("MaxActiveMovesPerTargetServer",$CAS_MRS_MAX_MOVES_PER_SERVER)
	
	$xml.configuration.ReplaceChild($newMRSSettings, $oldMRSSettings)

	$xml.save("$TargetDir\Bin\MSExchangeMailboxReplication.exe.config")
	

	#Modifications to .NET config to increase the requestQueueLimit
    $xml = New-Object XML
    $xml.PreserveWhitespace = $true
    $xml.Load("C:\Windows\Microsoft.NET\Framework64\v2.0.50727\CONFIG\machine.config")
    if ($xml.configuration.'system.web'.processModel.requestQueueLimit -ne "100000") {
        $xml.configuration.'system.web'.processModel.requestQueueLimit = "100000"
    }
    $xml.Save("C:\Windows\Microsoft.NET\Framework64\v2.0.50727\CONFIG\machine.config")
	

    #Set the appConcurrentRequestLimit to match the requestQueueLimit
    Set-WebConfigurationProperty /system.webserver/serverruntime -Name appConcurrentRequestLimit -Value 100000

    #Set the DeafultAppPool Queue Length
    $appPool = Get-Item IIS:\AppPools\DefaultAppPool
    If ($appPool.QueueLength –ne 10000) {
        $appPool.QueueLength = 10000
        $appPool | Set-Item
    }

	
	#Check for BES Throttling Policy and create if missing
	$throttle = Get-ThrottlingPolicy "BESPolicy"
	if (!($throttle)) {
		New-ThrottlingPolicy BESPolicy
		Set-ThrottlingPolicy BESPolicy -RCAMaxConcurrency $null -RCAPercentTimeInAD $null -RCAPercentTimeInCAS $null -RCAPercentTimeInMailboxRPC $null -EWSMaxConcurrency $null -EWSPercentTimeInAD $null -EWSPercentTimeInCAS $null -EWSPercentTimeInMailboxRPC $null -EWSMaxSubscriptions $null -EWSFastSearchTimeoutInSeconds $null -EWSFindCountLimit $null
	}
	
	#Check for Client Access Array in the AD Site and create if missing
	$casArray = Get-ClientAccessArray -Site $SITE_NAME
	if (!($casArray)) {
		New-ClientAccessArray -Name $INTERNAL_NAME -Fqdn $INTERNAL_NAME -site $SITE_NAME -DomainController $GC_SERVER
	}

	#Copy files used for OWA ToU Banner to WWWRoot directory. Add App_Code and checkforcookie.cs file
	Copy-Item $ScriptsDir\dod.gif -Destination c:\inetpub\wwwroot\dod.gif
	Copy-Item $ScriptsDir\web.config.tou -Destination c:\inetpub\wwwroot\web.config
	Copy-Item $ScriptsDir\TermsOfUse.aspx -Destination c:\inetpub\wwwroot\TermsOfUse.aspx


	if((Test-Path $TargetDir\App_Code) -eq $false) {
		New-Item $TargetDir\App_Code -ItemType directory
	}

	Copy-Item $ScriptsDir\CheckForCookie.cs -Destination "$TargetDir\App_Code\CheckForCookie.cs"
    Copy-Item $TargetDir\App_Code $TargetDir\ClientAccess\Owa -recurse

    #Install and configure the wildcard certificate.
    $password = (Get-Credential).password
    $fileData = ([Byte[]]$(Get-Content -Path $WildcardCert -Encoding byte -ReadCount 0))
    Import-ExchangeCertificate -FileData $fileData -Password $password | Enable-ExchangeCertificate –Services IIS,POP,IMAP -Confirm:$false

}

#Enables CAC authentication
function Configure-CACAuthForOA {
	Push-Location
	#Run script and commands to enable certificate-based authentication for Outlook Anywhere			
	. $TargetDir\Scripts\Enable-OutlookCertificateAuthentication.ps1
	Set-Location $ScriptsDir
	cscript adsutil.vbs set w3svc/1/SSLAlwaysNegoClientCert true
	Pop-Location
}

#Update IIS configuration - Move inetpub and logging directories, update logging fields, remove BAK files
#Warning, this restarts IIS, so should be done after all Exchange commands are run since it will kill the Exchange remote powershell
function Configure-CommonIIS {
    Push-Location

	Import-Module WebAdministration

	Set-Location $env:windir\system32\inetsrv
	.\appcmd add backup beforeRootMove
	iisreset /stop
	xcopy $env:systemdrive\inetpub D:\inetpub /O /E /I /Q

	# Check for IIS logging directory and create as needed.
	if((Test-Path $IIS_LOG_DIR) -eq $false) {
		Write-Host "       Creating $IIS_LOG_DIR"
		New-Item $IIS_LOG_DIR -ItemType directory
	}
	
    if (-not (Get-Item -Path HKLM:\System\CurrentControlSet\Services\WAS\Parameters)) { New-Item -Path HKLM:\System\CurrentControlSet\Services\WAS\Parameters }
	New-ItemProperty -Path HKLM:\System\CurrentControlSet\Services\WAS\Parameters -Name "ConfigIsolationPath" -PropertyType string -Value "D:\inetpub\temp\appPools"

	.\appcmd set config `/section:system.applicationHost/sites `/siteDefaults.traceFailedRequestsLogging.directory:"$IIS_LOG_DIR"
	.\appcmd set config `/section:system.applicationHost/sites `/siteDefaults.logfile.directory:"$IIS_LOG_DIR"
	.\appcmd set config `/section:system.applicationHost/log `/centralBinaryLogFile.directory:"$IIS_LOG_DIR"
	.\appcmd set config `/section:system.applicationHost/log `/centralW3CLogFile.directory:"$IIS_LOG_DIR"
	.\appcmd set config `/section:system.applicationhost/configHistory -path:d:\inetpub\history
	.\appcmd set config `/section:system.webServer/asp `/cache.disktemplateCacheDirectory:"d:\inetpub\temp\ASP Compiled Templates"
	.\appcmd set config `/section:system.webServer/httpCompression `/directory:"d:\inetpub\temp\IIS Temporary Compressed Files"
	.\appcmd set vdir "Default Web Site/" `/physicalPath:d:\inetpub\wwwroot
	.\appcmd set config `/section:httpErrors `/"[statusCode='401'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='403'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='404'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='405'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='406'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='412'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='500'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='501'].prefixLanguageFilePath:d:\inetpub\custerr"
	.\appcmd set config `/section:httpErrors `/"[statusCode='502'].prefixLanguageFilePath:d:\inetpub\custerr"
    .\appcmd set config "Default Web Site/" /section:serverruntime /alternateHostName:webmail /commit:apphost
    .\appcmd set config "Default Web Site/" /section:machineKey /validation:SHA1

	Get-WebConfiguration /system.applicationhost/sites/*/logfile | % {Set-WebConfigurationProperty $_.ItemXPath -Name logExtFileFlags -Value $IIS_LOG_FIELDS}
    
    if (!(Get-Item -Path HKLM:\Software\Microsoft\inetstp)) { New-Item -Path HKLM:\Software\Microsoft\inetstp }
	New-ItemProperty -Path HKLM:\Software\Microsoft\inetstp -Name "PathWWWRoot" -PropertyType string -Value "D:\inetpub\wwwroot"
    New-ItemProperty -Path HKLM:\Software\Microsoft\inetstp -Name "PathFTPRoot" -PropertyType string -Value "D:\inetpub\ftproot"

    Add-Content D:\inetpub\wwwroot\robots.txt "User-agent: *`r`nDisallow: /"

	#Remove C:\Inetpub
	takeown `/F C:\inetpub\custerr\en-us\500-100.asp `/A
	icacls C:\inetpub\custerr\en-us\500-100.asp `/grant administrators:F
	Remove-Item C:\inetpub -recurse -force
    
    Pop-Location
	
}

#endregion

#region Main
Clear-Host

#Create the directory for the automation logs
if((Test-Path $INSTALL_LOG_DIR) -eq $false) {
	Write-Host "Creating Logging Directory at $INSTALL_LOG_DIR"
	New-Item $INSTALL_LOG_DIR -ItemType directory
}

Start-Transcript -Path $INSTALL_LOG_DIR\ExchInstaller.txt -Append

#Disable UAC if necessary
$NoUAC = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin
if ($NoUAC.ConsentPromptBehaviorAdmin -ne 0) {
	Write-host "Disabling UAC"
	#Disables UAC Default Settings
	Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name ConsentPromptBehaviorAdmin -Value 0
	Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name EnableLUA -Value 0
	Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System -Name PromptOnSecureDesktop -Value 0
}

Install-Exchange2010

#endregion