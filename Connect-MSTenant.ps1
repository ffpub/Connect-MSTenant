<#
.SYNOPSIS
Microsoft Service Delegated Access Connection Helper for Interactive Logins
.DESCRIPTION
Wrapper for connection commands to streamline:
	Connecting multiple services at once.
	Switching betweeen tenancies an account has delegated access to.

This saves a lot of time entering MFA when switching between tenancies using delegated access

Connecting MSOnline first will avoid needing to re-use MFA for delegated tenant logins
ExchangeOnlineManagement will re-use MSOnline's authentication
AzAccounts will require MFA on first use, then remains quiet when changing subscriptions
Most others might want MFA on each use

On the first run with -Find, or if the list is too old, MSOnline & AzAccounts will connect to prepare a CSV of Tenant Info 
If CSV File is not clean, delete it

Todo: 
	Thoroughly test Sharepoint & Graph 
	Allow Sharepoint to use CSV for TenantInfo & -Find
	Find works with sharepoint I just broke it by changing the params
	Add Access Token and Cert Auth Commands (Cert Thumbnail From Params)

.PARAMETER Find
Tenant or AzSubrciption Search Term
Using Find when connecting will use matching tenant info from Tenant Info CSV to connect
If Find is not used, service prompts for login, use their admin credentials (delegated access and tenant list not checked if Find is not used)
.PARAMETER DebugLogging
Enable Verbose Logging
.PARAMETER ExchangeOnlineManagement
Connect ExchangeOnlineManagement
.PARAMETER MicrosoftGraph
Connect MicrosoftGraph
.PARAMETER MSOnline
Connect MSOnline
.PARAMETER AzureAD
Connect AzureAD
.PARAMETER AzAccounts
Connect Az.Accounts
.PARAMETER PnPPowershell
Connect PnP.Powershell
Cannot be used in combination with other service connection params
.PARAMETER SharepointAdminURL
If connecting with PnP.Powershell, SharepointAdminURL is required
.PARAMETER DisconnectAll
Completely Disconnect All Services
This overrides any other arguments and exits when done
.PARAMETER Help
Show Help

.EXAMPLE
Connect-MSTenant.ps1 -ExchangeOnlineManagement 

Connect specified service with provided credentials
.EXAMPLE
Connect-MSTenant -ExchangeOnlineManagement -Find "TenantName"
Connect-MSTenant -ExchangeOnlineManagement -AzAccounts -Find "TenantName"


Search for "TenantName" in tenant info CSV and connect specified service/s
.EXAMPLE
Connect-MSTenant -ExchangeOnlineManagement -Find "TenantName" -DisconnectAll

Disconnect everything completely, don't check if connected first
-DisconnectAll overrides any other arguments and exits when done
.EXAMPLE
Connect-MSTenant -MSOnline -ExchangeOnlineManagement -AzureAD -Find "TenantName" -DebugLogging

Search for "TenantName" in tenant info CSV and connect specified services
Vebose logging enabled
.EXAMPLE
Connect-MSTenant -PnPPowershell -SharepointAdminURL "https://tenant-admin.sharepoint.com" -DebugLogging

Connect to specified Sharepoint tenancy
Vebose logging enabled
.INPUTS
System.String
.NOTES
	FunctionName : Connect-MSTenant
	Created by   : Fraser Fitzgerald
	Last Update  : v4 2/7/22
.LINK
	https://github.com/ffitz-public/Connect-MSTenant
#>

param (
	[CmdletBinding(DefaultParametersetName='ShowHelp')]
	[Parameter(ParameterSetName='Services',Mandatory=$false)][String]$Find,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$DebugLogging,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$ExchangeOnlineManagement,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$MicrosoftGraph,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$MSOnline,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$AzureAD,
	[Parameter(ParameterSetName='Services',Mandatory=$false)][Switch]$AzAccounts,
	[Parameter(ParameterSetName='Sharepoint',Mandatory=$false)][switch]$PnPPowershell,      
	[Parameter(ParameterSetName='Sharepoint',Mandatory=$true)][string]$SharepointAdminURL,
	[Parameter(ParameterSetName='DisconnectAll',Mandatory=$false)][Switch]$DisconnectAll,
	[Parameter(ParameterSetName='ShowHelp',Mandatory=$false)][Switch]$Help

)

$DateStarted = Get-Date
$TenantInfoCSVFilePath = (get-childitem $PROFILE).Directory.FullName + "\365TenantInfo.csv"


#TODO Add
#See bit with Lync for security & compliance connections
#https://www.michev.info/Blog/Post/1771/hacking-your-way-around-modern-authentication-and-the-powershell-modules-for-office-365
#
#Prefill Usernames?
#   Connect-AzureAD -AccountId $globalAdminAcct
#	Connect-ExchangeOnline -UserPrincipalName $globalAdminAcct
#	Connect-IPPSSession -UserPrincipalName $globalAdminAcct
#TODO Get Access Token from Az Module to use with REST API calls - See also: Get-AzAccessToken.ps1 
#Get-AzAccessToken -ResourceTypeName MSGraph
#Get-AzAccessToken -ResourceTypeName Arm

#Also Possible using ADAL Binaries? May be shorter duration
#$cache = [Microsoft.IdentityModel.Clients.ActiveDirectory.TokenCache]::DefaultShared
#$cache.ReadItems()


if ($Help) {
	$ThisScriptName = $MyInvocation.MyCommand.Name
	Get-Help $ThisScriptName -Detailed
	exit
}

function Log {
	param (
		[Parameter(mandatory=$true)]
			[String]$Text,
		[Switch]$Warning,
		[Switch]$Success
	)

	$LogTime = Get-date -Format yyMMdd-HH:mm:ss-fff 
	$Preface = "$LogTime | "

	if ($Success) {
		$Colour = "Green"
		$Content = $Preface + "`t" + $Text
	}
	elseif ($Warning) {
		$Colour = "Red"
		$Content = $Preface + "`t" + $Text 
	}
	elseif ($Text -like "|*") {
		$Colour = "Gray"
		$Content = $Preface + "`t" + $Text
	}
	else {
		$Colour = "White"
		$Content = $Preface + "`t" + $Text
	}

	if ( $DebugLogging -or $Success -or $Warning) {
		write-host -ForegroundColor $Colour $Content
	}
}


$IgnoreParams = @(
	"Find"
	"DebugLogging"
	"SharepointAdminURL"
)

#TODO Check Inputs & prepare a list of the specified services to connect
#Specified Params; Ignore params unrelated to user
$Script:Params = $PSBoundParameters
If ($Params) { $Script:Services_Params = $Params.Keys | % { if ( -not ( $IgnoreParams -contains $_ )  ) { $_ } } }


#TODO Environment Checks
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$Script:RunningAsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if( $Global:PSVersionTable.PSVersion.Major -eq 7 -and $AzureAD ) { Log -Warning "[Compatibility] AzureAD Doesn't Work in PS7" }
if( $Global:PSVersionTable.PSVersion.Major -eq 7 -and $MSOnline ) { Log -Warning "[Compatibility] MSOnline MFA Prompt Doesn't Work in PS7" }
#if ( $Global:PSVersionTable.PSVersion.Major -eq 5 -and $AzAccounts) { Log -Warning "[Compatibility] Az Accounts Doesn't Work in PS5" }
if ($AzureAD -and $AzAccounts) { Log -Warning "[Conflict] AzureAD And AzAccounts Cannot be Used Together" }
if (-not ($Script:RunningAsAdmin) ) { Log "[CurrentUser] Not running as admin; If module requires installation, this will be done from user context and require installation again next time." }


function CheckModulesInstalled {
	param (
		[hashtable]$SpecifiedService
	)

	#TODO Limit Graph & AZ Modules

	if ( -not (Get-Module -Name ($SpecifiedService.ModuleName)) ) { 
		Log "$($SpecifiedService.ModuleName): Not Found. Installing & Importing..."
        if (-not ($Script:RunningAsAdmin) ) { install-module $($SpecifiedService.ModuleName) -Scope CurrentUser -Force } 
        else { install-module $($SpecifiedService.ModuleName) -Scope AllUsers -Force }
        import-module $($SpecifiedService.ModuleName) 
	} 

	if (Get-Module -Name $($SpecifiedService.ModuleName)) {
		return $true
	}
	else { return $false}
}


function CheckServiceConnection {
	param (
		[hashtable]$SpecifiedService,
		[PSCustomObject]$Connection,
		[Switch]$TenantConnection
	)

	#TenantConnection switch is only used for header

	#TODO GET CSP TenantID to use for MSOL switching - Add this at the CSV creation

	Log "Checking Connection State..."

	#Log "|-------------------------------------------------------------------------------------"
	#Log "| [$($SpecifiedService.ModuleName)] CheckServiceConnection [TenantConnection? $tenantConnection]"
	#Log "| DestinationTenant: $($Connection.DestinationTenant)" 
	#Log "|-------------------------------------------------------------------------------------"

	$ConnectionTestCommand = $SpecifiedService.IsConnectedCommand 
	$ConnectionCheckResult = try { Invoke-Command $ConnectionTestCommand  } 
	catch [System.InvalidOperationException] {
		if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "The current connection holds no SharePoint context.") {
			Log "`t$($SpecifiedService.ModuleName) Not Connected"
		}
	}
	catch { 
		#TODO RETURN TO THIS WHEN IT BREAKS
		$_.Exception | % { 
			Log -Warning "`tConnection Check Error Getting Current Connection Result"
			Log -Warning "`tName: $($_.gettype().fullname)" 
			Log -Warning "`tMessage: $($_.Message)"
			Log -Warning "`tError { $($_)"
		}
	}
	
	$CurrentTenancy = try { Invoke-Command $SpecifiedService.GetConnectedTenantNameCommand -ErrorAction SilentlyContinue  } catch [System.Management.Automation.CommandNotFoundException] { 
		#Catch error missing command
		if ($SpecifiedService.ModuleName -eq 'ExchangeOnlineManagement') {
			$ExchangeGetOrganizationConfigNotFound = (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "Get-OrganizationConfig")
			if ($ExchangeGetOrganizationConfigNotFound -and $ConnectionCheckResult) {
				#TODO Connected to ExchangeOnline but no Get-OrgConfig command

				Log "`t`tGet-OrganizationConfig command not loaded"

				$Script:CSPName
			}
		}
	}
	catch [System.Management.Automation.RuntimeException] {
		if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "You cannot call a method on a null-valued expression.") {
			Log "`t$($SpecifiedService.ModuleName) Not Connected"
		}
	}
	catch [System.InvalidOperationException] {
		if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "The current connection holds no SharePoint context.") {
			Log "`t$($SpecifiedService.ModuleName) Not Connected"
		}
	}
	catch {
		#TODO RETURN TO THIS WHEN IT BREAKS
		$_.Exception | % { 
			Log -Warning "`tConnection Check Error Getting Current Tenant Name"
			Log -Warning "`t`tName: $($_.gettype().fullname)" 
			Log -Warning "`t`tMessage: $($_.Message)"
			Log -Warning "`t`tError { $($_)"
		}
	}

	
	if ($ConnectionCheckResult) {

		$Connection.Connected = $True

		if (-not $Find -and $ExchangeGetOrganizationConfigNotFound) {
			$Connection.CurrentTenant = if ($CurrentTenancy) { $CurrentTenancy } else { (Get-AcceptedDomain | where {$_.Default -eq $True}).DomainName }
			Log "`tConnected to: $($Connection.CurrentTenant)"
			$Connection.Connected = $True
		}
		elseif ($CurrentTenancy -eq $Script:CSPName -or $CurrentTenancy -eq $Script:CSPAzSubscriptionName -or ($CurrentTenancy -and -not $Script:CSPName -and -not $Connection.DestinationTenant -and $Find)) {
			Log "`tCurrently Connected to CSP: $CurrentTenancy"
			$Connection.CurrentTenant = $CurrentTenancy
			$Connection.ConnectedToCSP = $True
		}
		elseif ($CurrentTenancy) {
			Log "`tCurrently Connected to $CurrentTenancy"
			$Connection.CurrentTenant = $CurrentTenancy
			$Connection.ConnectedToCSP = $False
		}
		else {
			Log "`tCurrently Connected. No Tenant Name? $CurrentTenancy"
			$Connection.CurrentTenant = $null
		}
	}
	else {
		Log "`tNot Currently Connected"
		$Connection.CurrentTenant = $null
		$Connection.Connected = $False
	}

	#Log "| End of Connection Check [Connected: $($Connection.Connected) Current: $CurrentTenancy]"
	#Log "|-------------------------------------------------------------------------------------"

	return $Connection
}

function ConnectToSpecifiedService {
	param (
		[hashtable]$SpecifiedService,
		[PSCustomObject]$Connection,
		[Switch]$TenantConnection
		
	)

	
	if (-not $TenantConnection) { 
		$ConnectionCommand = $SpecifiedService.ConnectCommand
	}
	else {
		$ConnectionCommand = $SpecifiedService.ConnectTenantCommand
	}

	$HeaderCheck = if ($Connection.ConnectedToCSP -and $Connection.ReturningToCSP) { "Connected To CSP: $($Connection.CurrentTenant)" } elseif ($Connection.ConnectedToCSP -and -not $Connection.ReturningToCSP) { "Leaving CSP: $($Connection.CurrentTenant)" } elseif ($Connection.ReturningToCSP) { "ReturningToCSP: $($Connection.ReturningToCSP)" } 
	Log "|-------------------------------------------------------------------------------------"
	Log "| [$($SpecifiedService.ModuleName)] ConnectToSpecifiedService $(if ($Find) { "[ TenantConnection (Search: $Find) ]" } )"
	Log "| CurrentTenant: $($Connection.CurrentTenant) $(if ($Connection.CurrentDestination) { "| DestinationTenant: $($Connection.CurrentDestination)" } ) $(if ($HeaderCheck) { "| $HeaderCheck"})"
	Log "|-------------------------------------------------------------------------------------"
	#Connect acter successful or unsuccessful check; If Not a new connection and check after connection failed; failse else connectioncheck has ensured we're in the right place 

	$NewConnectionAttempt = try { Invoke-Command $ConnectionCommand  } catch [System.Management.Automation.Remoting.PSRemotingTransportException] { 
		#TODO RETURN TO THIS WHEN IT BREAKS
			$MaxRunspaces = (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "ActiveRunspaces").Line
			if ($MaxRunspaces) { 
				Log -Warning "You may have too many concurrent sessions. `n`t`t`t`t$MaxRunspaces `n`t`t`t`tIf you close the powershell window they'll expire in 15 minutes."	
			}
		} 
		catch [Microsoft.Identity.Client.MsalClientException] {
			if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "User canceled authentication.") {
				Log "`t$($SpecifiedService.ModuleName): User canceled authentication"
			}
		}
		catch { $_.Exception | % { 
			Log -Warning "`tNew Connection Attempt Error Connecting"
			Log -Warning "`tName: $($_.gettype().fullname)" 
			Log -Warning "`tMessage: $($_.Message)"
			Log -Warning "`tError { $($_)"
		}	
	}

	Log "New Connection:"
	$NewConnectionState = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection:$TenantConnection -Connection $Connection

	if ($NewConnectionState.Connected -eq $True) {
		Log "New Connection Successful [ $($NewConnectionState.CurrentTenant) ]"

		$Connection.Connected = $True
	}
	else {
		Log "New Connection Failed [ $NewConnectionState ]"

		$Connection.Connected = $False
		$NewConnectionState = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection:$TenantConnection -Connection $Connection
	}

	

	#Log "| [$($SpecifiedService.ModuleName)] End of ConnectToSpecifiedService [Output: $($Connection.Connected)"
	#Log "|-------------------------------------------------------------------------------------"

	return $NewConnectionState
}

function DisconnectService {
	param (
		[hashtable]$SpecifiedService,
		[PSCustomObject]$Connection
	)

	#Log "|-------------------------------------------------------------------------------------"
	#Log "| [$($SpecifiedService.ModuleName)] DisconnectService [All: $DisconnectAll ]"
	#Log "| ReturningToCSP: $($Connection.ReturningToCSP)"
	#Log "|-------------------------------------------------------------------------------------"

	#TODO If Returning to CSP, these need to switch instead of disconnecting; If Disconnecting Everything, do the full disconnect
	if ($SpecifiedService.ModuleName -eq "MSOnline" -and $DisconnectAll) {
		Log "`t Disconnecting $($SpecifiedService.ModuleName) completely. Clearing user session state"
		[void] [Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()	
	} 
	elseif ($SpecifiedService.ModuleName -eq "Az.Accounts" -and $DisconnectAll) {
			Log "`t Disconnecting $($SpecifiedService.ModuleName) completely. Clearing user session state"
			Disconnect-AzAccount -Scope Process
	}
	else {

		$DisconnectionResult = try { Invoke-Command $SpecifiedService.DisconnectCommand } 
		catch [System.NotSupportedException] {
			Log "`t Suppressing disconnection error for ExchangeOnline"
			$Connection.Connected = $False

		}
		catch [System.NullReferenceException] {
			if ($SpecifiedService.ModuleName -eq 'AzureAD') {
				if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "Object reference not set to an instance of an object.") {
					Log "`t AzureAD Not Connected"
				}
			}
		}
		catch [System.Exception] {
			if ($SpecifiedService.ModuleName -eq 'Microsoft.Graph') {
				if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "No application to sign out from.") {
					Log "`t Microsoft.Graph Not Connected"
				}
			}
			if ($SpecifiedService.ModuleName -eq 'PnP.Powershell') {
				if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "No connection to disconnect") {
					Log "`t PnP.Powershell Not Connected"
				}
			}
			if ($SpecifiedService.ModuleName -eq 'AzAccounts') {
				if (($_.Exception.Message).Split("`n") | Select-String -SimpleMatch "Cannot validate argument on parameter 'Subscription'") {
					#TODO Connected to ExchangeOnline but no Get-OrgConfig command
					Log "`t No Azure Subscription Connected"
				}
			}
		}
		catch { $_.Exception | % { 
				Log -Warning "`tDisconnect: Error Disconnecting [$($SpecifiedService.ModuleName)]"
				Log -Warning "`tName: $($_.gettype().fullname)" 
				Log -Warning "`tMessage: $($_.Message)"
				Log -Warning "`tError { $($_)"
			} 
		}

		if ($DisconnectionResult) { 	
			$Connection.Connected = $False

			Log "[Disconnect] $($SpecifiedService.ModuleName) has been disconnected"
		} else {
			if ( -not $SpecifiedService.ModuleName -eq 'ExchangeOnlineManagement') {
				Log -Warning "[Disconnect] $($SpecifiedService.ModuleName) failed?" 
			}
		}
	}
	#Log "| [$($SpecifiedService.ModuleName)] End of DisconnectService [Output: $($Connection.Connected)]"
	#Log "|-------------------------------------------------------------------------------------"

	return $Connection
}


# Connect-MgGraph -TenantId "tenant.onmicrosoft.com" -Scope "Sites.FullControl.All", "Directory.ReadWrite.All"


$Script:ServiceList = @(
	@{
		ModuleName = "MSOnline" 
		ConnectCommand =   [scriptblock]::Create(" Connect-MsolService")
		ConnectTenantCommand = [scriptblock]::Create(" `$Global:PSDefaultParameterValues=@{`"*-Msol`*:TenantID`"=`$PartnerTenantID}" )
		GetConnectedTenantNameCommand = [scriptblock]::Create(" (Get-MsolPartnerInformation -ErrorAction SilentlyContinue).PartnerCompanyName ")
		IsConnectedCommand =   [scriptblock]::Create(" Get-MsolCompanyInformation -ErrorAction SilentlyContinue " )
		IsTenantConnectedCommand =   [scriptblock]::Create(" (Get-MsolCompanyInformation -ErrorAction SilentlyContinue).DisplayName " )
		DisconnectCommand = [scriptblock]::Create(" if (`$Global:PSDefaultParameterValues) { `$Global:PSDefaultParameterValues.Clear() } ") 
	},
	@{
		ModuleName = "Az.Accounts" 
		ConnectCommand =   [scriptblock]::Create(" Connect-AZAccount -InformationAction SilentlyContinue ")
		ConnectTenantCommand = [scriptblock]::Create(" Select-AzSubscription -SubscriptionId `$Script:AzSubscriptionId " )
		GetConnectedTenantNameCommand = [scriptblock]::Create(" (Get-AzContext -ErrorAction SilentlyContinue | select -ExpandProperty Name).Split(' ')[0] ")
		IsConnectedCommand =   [scriptblock]::Create(" Get-AzContext -ErrorAction SilentlyContinue " )
		IsTenantConnectedCommand =   [scriptblock]::Create(" if ( `$((Get-AzContext).SubscriptionName -like `"*`$Find*`") ) { (Get-AzContext | select -ExpandProperty Name).Split(' ')[0] } " )
		DisconnectCommand = [scriptblock]::Create(" Select-AzSubscription -SubscriptionId `$Script:CSPAzSubscriptionID ") 
	},
	@{
		ModuleName = "AzureAD" 
		ConnectCommand =   [scriptblock]::Create(" Connect-AzureAD -InformationAction SilentlyContinue ")
		ConnectTenantCommand = [scriptblock]::Create(" Connect-AzureAD -TenantId `$PartnerTenantID -AccountId `$Script:AADAccountID " )
		GetConnectedTenantNameCommand = [scriptblock]::Create(" try { (Get-AzureADTenantDetail).DisplayName } catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { `$False } ")
		IsConnectedCommand =   [scriptblock]::Create(" try { Get-AzureADTenantDetail } catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { `$False } " )
		IsTenantConnectedCommand =   [scriptblock]::Create(" try { Get-AzureADTenantDetail } catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] { `$False } " )
		DisconnectCommand = [scriptblock]::Create(" Disconnect-AzureAD -ErrorAction SilentlyContinue") 
	},
	@{
		ModuleName = "Microsoft.Graph" 
		ConnectCommand =   [scriptblock]::Create(" Connect-MgGraph -Scope 'Directory.ReadWrite.All'")
		ConnectTenantCommand = [scriptblock]::Create(" Connect-MgGraph -TenantId `$PartnerTenantId -Scope 'Directory.ReadWrite.All' " )
		GetConnectedTenantNameCommand = [scriptblock]::Create("  (Get-MgOrganization -ErrorAction SilentlyContinue).DisplayName ")
		IsConnectedCommand =   [scriptblock]::Create(" Get-MgContext -ErrorAction SilentlyContinue ")
		IsTenantConnectedCommand =   [scriptblock]::Create(" Get-MgContext -ErrorAction SilentlyContinue " )
		DisconnectCommand =  [scriptblock]::Create(" Disconnect-Graph ")
	},
	@{
		ModuleName = "PnP.Powershell" 
		ConnectCommand =   [scriptblock]::Create(" Connect-PnPOnline -Interactive -Url `$Script:SharepointAdminUrl ")
		ConnectTenantCommand = [scriptblock]::Create(" Connect-PnPOnline -Interactive -Url `$DelegatedTenantSharepointAdminUrl " )
		GetConnectedTenantNameCommand = [scriptblock]::Create(" (Get-pnpContext).Url ")
		IsConnectedCommand =   [scriptblock]::Create(" Get-PnPConnection -ErrorAction SilentlyContinue ")
		IsTenantConnectedCommand =   [scriptblock]::Create(" Get-PnpConnection " )
		DisconnectCommand =  [scriptblock]::Create(" Disconnect-PnPOnline ")
	},
	@{
		ModuleName = "ExchangeOnlineManagement" 
		ConnectCommand =   [scriptblock]::Create(" Connect-ExchangeOnline -ShowBanner:`$False -WarningAction SilentlyContinue ")
		ConnectTenantCommand = [scriptblock]::Create(" Connect-ExchangeOnline -DelegatedOrganization `$PartnerDomain -ShowBanner:`$False -WarningAction SilentlyContinue  " )
		GetConnectedTenantNameCommand = [scriptblock]::Create(" (Get-OrganizationConfig -Erroraction SilentlyContinue).DisplayName ")
		IsConnectedCommand =   [scriptblock]::Create(" Get-Command -Name 'Get-Mailbox' -ErrorAction SilentlyContinue ")
		IsTenantConnectedCommand = [scriptblock]::Create(" Get-Command -Name 'Get-OrganizationConfig' -ErrorAction SilentlyContinue " )
		DisconnectCommand =  [scriptblock]::Create(" Disconnect-ExchangeOnline -Confirm:`$False ")
	}
)


function ServiceHandler {
	param (
		[Switch]$CSVUpdate,
		[String]$SpecifiedServiceName,
		[Switch]$ConnectService,
		[Switch]$ConnectToTenantService,
		[pscustomobject]$TenantDetails,
		[Switch]$CheckConnected,
		[Switch]$DisconnectService
	)

	#For each input param if not the first, add a separator before adding it to the string
	$InputParams = ($PSBoundParameters.Keys | % { $i = 0; $Out = $_ ; if ($i -gt 0) { " | " + $Out ; $i++ } } )

	#Log "|-------------------------------------------------------------------------------------"
	#Log "| [$SpecifiedServiceName] ServiceHandler [TenantDetails: $($TenantDetails.Name)]"
	#Log "| ServiceHandler Params: $InputParams"
	#Log "|-------------------------------------------------------------------------------------"

	$Success = $False

	#TODO TENANT DETAILS
	$DelegatedTenantSharepointAdminUrl = "GET THIS FROM CSV"
	$PartnerName = $TenantDetails.Name
	$PartnerDomain = $TenantDetails.DefaultDomain
	$PartnerTenantID = $TenantDetails.TenantID
	$Script:SharepointAdminUrl = if (-not $Find -and $SharepointAdminURL) { $SharepointAdminURL } else { $TenantDetails.SharepointAdminURL }
	$Script:AzSubscriptionName = $TenantDetails.AzSubscriptionName
	$Script:AzSubscriptionId = $TenantDetails.AzSubscriptionId 

	$CSP = $Script:TenantInfoList | Where { $_.IsCSP -eq $True }
	$Script:CSPName = $CSP.Name
	$CSPTenantId = $CSP.TenantId
	$Script:CSPAzSubscriptionName = $CSP.AzSubscriptionName
	$Script:CSPAzSubscriptionId = $CSP.AzSubscriptionId

	$ConnectionState = [PSCustomObject]@{
		CurrentTenant = $null
		CurrentDestination = $null
		DestinationTenant = if ($Find -and $SpecifiedServiceName -eq 'AzAccounts') {$Script:AzSubscriptionName} else { $PartnerName }
		ConnectedToCSP = $null
		ReturningToCSP = $null
		Connected = $null
	}


	$Service = if ($SpecifiedServiceName) { $ServiceList | where { $SpecifiedServiceName -eq $( ($_.ModuleName).Replace(".","") ) } }
	if (-not $Service) { Continue }

	$Success = $null
	$Service | foreach-object {

		
		$SpecifiedService = $_ 

		if ($CheckConnected) {
			#Log "[ServiceHandler] CheckConnected"
			
			$CheckIfServiceIsConnected = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection -Connection $ConnectionState

			if ($CheckIfServiceIsConnected.Connected -eq $True) {
				$Success = $CheckIfServiceIsConnected
				Log "`tInitial connection check succeeded [ $CheckIfServiceIsConnected ]"

			}
			else {
				Log "`tInitial connection check failed [ $CheckIfServiceIsConnected ]"

			}
			#Log "[ServiceHandler] End CheckConnected"
		}

		if ($ConnectService) {
			Log "Checking and installing module for $($SpecifiedService.ModuleName)"

			$RequiredModulesAreInstalled = CheckModulesInstalled -SpecifiedService $SpecifiedService
			if ($RequiredModulesAreInstalled) {
				Log "Module ready"

				Log "[$SpecifiedServiceName] Connection Specified with Credentials"
				
				if ($CSVUpdate -or -not $Find) {
					$CurrentConnection = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection -Connection $ConnectionState

					if ($CurrentConnection.ConnectedToCSP -or $CurrentConnection.Connected) {
						#If Updating CSV, Connected to something already
						ServiceHandler -SpecifiedServiceName $SpecifiedServiceName -DisconnectService -Connection $CurrentConnection | out-null
					}
				}

				$CurrentConnection = ConnectToSpecifiedService -SpecifiedService $SpecifiedService -Connection $ConnectionState

				if ($CurrentConnection.CurrentTenant) {
					Log -Success "[$SpecifiedServiceName]> New Connection Successful: $($CurrentConnection.CurrentTenant)"
					$Success = $True
				}
				else {
					Log -Warning "[$SpecifiedServiceName]> New Connection Failed:  $($CurrentConnection.CurrentTenant)"
				}

			}
			else { 
				Log -Warning "Failed to Install or Check for Module: $($SpecifiedService.ModuleName)" 
			}
		}

		if ($ConnectToTenantService) {

			#TODO If servicename is msol, connect to CSP then set TenantID in tenant stage
			#TODO If not msol, connect to Tenant
			#TODO If not correct Tenant, disconnect and connect service to Tenant (Not to CSP; Only MSOL ought to connect to CSP?)

			#Log "[ServiceHandler] ConnectToTenantService"

			Log "Checking module for $($SpecifiedService.ModuleName)"
			$RequiredModulesAreInstalled = CheckModulesInstalled -SpecifiedService $SpecifiedService
			if ($RequiredModulesAreInstalled) {
				Log "Module ready"

				Log "[$SpecifiedServiceName] Connecting Tenant. Destination: $(if ($SpecifiedServiceName -eq 'AzAccounts') { $Script:CSPAzSubscriptionName } else { $Script:AzSubscriptionName })"

				$CurrentConnection = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection -Connection $ConnectionState

				#TODO NOTE: Run disconnection for these, disconnection command not needed for others
				if (($SpecifiedServiceName -eq "MSOnline" -or $SpecifiedServiceName -eq "AzureAd" -or $SpecifiedServiceName -eq "ExchangeOnlineManagement") -and $CurrentConnection.Connected -eq $True -and $CurrentConnection.CurrentTenant -ne $CurrentConnection.DestinationTenant) {
					# MSOnline Connected, Not in Destination Tenant; Clear PSDefaultParam TenantID
					Log "Returning to CSP With Disconnect. Current: [$($CurrentConnection.CurrentTenant)] Destination: [$($CurrentConnection.CurrentDestination)]"

					ServiceHandler -SpecifiedServiceName $SpecifiedServiceName -DisconnectService -Connection $CurrentConnection
					$CurrentConnection.ConnectedToCSP = $True
					$CurrentConnection.CurrentDestination = if ($SpecifiedServiceName -eq 'AzAccounts') { $Script:CSPAzSubscriptionName } else { $Script:AzSubscriptionName }
				}
				elseif ($SpecifiedServiceName -eq "AzAccounts" -and $CurrentConnection.Connected -and $CurrentConnection.CurrentTenant -ne $CurrentConnection.DestinationTenant) {
					$CurrentConnection.ConnectedToCSP = $True
					$CurrentConnection.ReturningToCSP = $False
					$CurrentConnection.CurrentDestination = $Script:AzSubscriptionName
					
					Log "Switching Context: [$($CurrentConnection.CurrentTenant) > $($CurrentConnection.CurrentDestination)]"
					$CurrentConnection = ConnectToSpecifiedService -SpecifiedService $SpecifiedService -Connection $CurrentConnection -TenantConnection
				}
				elseif ($CurrentConnection.ConnectedToCSP -eq $True) {
					$CurrentConnection.ReturningToCSP = $False
					$CurrentConnection.CurrentDestination = $CurrentConnection.DestinationTenant
					
					Log "Already connected to CSP Tenancy [ $($CurrentConnection.CurrentTenant) ]"
				}
				elseif ($CurrentConnection.Connected -eq $True -and $CurrentConnection.CurrentTenant -ne $CurrentConnection.DestinationTenant -and -not $SpecifiedServiceName -eq "AzAccounts" ) {
					$CurrentConnection.ReturningToCSP = $True

					Log "Returning to CSP Tenancy"
					
					Log -Warning $Script:CSPAzSubscriptionName
					#TODO ReturningToCSP is set in script scope and used in DisconnectService function
					ServiceHandler -SpecifiedServiceName $SpecifiedServiceName -DisconnectService -Connection $CurrentConnection
					$CurrentConnection = ConnectToSpecifiedService -SpecifiedService $SpecifiedService -Connection $CurrentConnection		
				}
				elseif ($CurrentConnection.Connected -eq $False) {
					$CurrentConnection.CurrentDestination = $Script:CSPName
					Log "New Connection to CSP Tenancy [ $($CurrentConnection.CurrentTenant) ]"
					$CurrentConnection = ConnectToSpecifiedService -SpecifiedService $SpecifiedService -Connection $CurrentConnection

					
				}
				elseif ($CurrentConnection.CurrentTenant -eq $CurrentConnection.DestinationTenant) {
					Log -Success "[$SpecifiedServiceName] Already connected to [ $($CurrentConnection.CurrentTenant) ] "
					$Success = $True
				}

				#TODO Connect after csp connection
				if (-not $Success -eq $True ) {

					$CurrentConnection = CheckServiceConnection -SpecifiedService $SpecifiedService -TenantConnection -Connection $CurrentConnection

					if ($CurrentConnection.ConnectedToCSP ) {
						Log -Success "[$SpecifiedServiceName] Connected to CSP: $($CurrentConnection.CurrentTenant)"
						$CurrentConnection.ReturningToCSP = $False
						$CurrentConnection.CurrentDestination = $CurrentConnection.DestinationTenant

						if ($SpecifiedServiceName -eq "AzureAd") {
							$Script:AADAccessToken = [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens['AccessToken']
							if ($Script:AADAccessToken) {
								Log -Success "[$SpecifiedServiceName] Acquired Access Token $($Script:AADAccessToken)"
								#TODO return for use? ($Script:AADAccessToken).AccessToken
								#$Script:AADAccountID = (Get-AzureADCurrentSessionInfo).Account
							}
							else { Log "[$SpecifiedServiceName] Failed to acquire Access Token."}
						}

						$CurrentConnection = ConnectToSpecifiedService -SpecifiedService $SpecifiedService -TenantConnection -Connection $CurrentConnection
					}
					
					if ($CurrentConnection.CurrentTenant -eq $CurrentConnection.DestinationTenant) {
						Log -Success "[$SpecifiedServiceName] New Connection to tenant Successful: $($CurrentConnection.CurrentTenant)"
						$Success = $True
					}
					else {
						Log -Warning "[$SpecifiedServiceName] New Connection to tenant Failed:  $($CurrentConnection.CurrentDestination)"
					}
				}
							
			}
			else { 
				Log "[ModulesIssue] Failed to Install or Check for Module: $($SpecifiedService.ModuleName)" 
			}
			Log "[ServiceHandler] End ConnectToTenantService"
		}
		
		if ($DisconnectService) {
			#TODO Note ReturningToCSP is set in script scope and used in function

			$ServiceDisconnection = DisconnectService -SpecifiedService $SpecifiedService -Connection $ConnectionState
			$Success = if ($ServiceDisconnection.Connected -eq $False) { $True } else { $False}

			#Log "[ServiceHandler] End DisconnectService"
		}
	}

	#Log "| [$SpecifiedServiceName] End of ServiceHandler [Output: $Success]"
	#Log "|-------------------------------------------------------------------------------------"
	return $Success
}

function CheckTenantInfoCSV {
	Log "[Checking Tenant Info CSV]"
	$TenantInfoCSVFileExists = Test-Path $TenantInfoCSVFilePath -ErrorAction SilentlyContinue
	if ($TenantInfoCSVFileExists) {
		Log "TenantInfoCSVFile Exists: $TenantInfoCSVFilePath"
		
		$TenantInfoCSVFile = Get-Item $TenantInfoCSVFilePath
		$TenantInfoCSVFileLastUpdate = ($TenantInfoCSVFile).LastWriteTime 
		$FileTooOldDays = 7
		$FileAge = New-TimeSpan -Start $TenantInfoCSVFileLastUpdate -End $DateStarted

		$CheckCSV_TestCSPAz = ConvertFrom-Csv $( Get-Content $TenantInfoCSVFilePath )

		$CSP = $CheckCSV_TestCSPAz | Where { $_.IsCSP -eq $True }
		$CSPName = $CSP.Name
		$CSPAzSubscription = $CSP.AzSubscriptionId

		if (-not $CSPAzSubscription) {
			$ForceUpdate = $True
			Log -warning "CSV Check Didn't find Az Sub for CSP in TenantInfoCSVFile"
		}
		else {
			Log "CSV Check verified Az Sub for CSP in TenantInfoCSVFile"

		}

		if ($TenantInfoCSVFileLastUpdate -lt ($DateStarted).AddDays(-$FileTooOldDays)) {
			Log "TenantInfoCSVFile Is Old: $TenantInfoCSVFilePath [$( [math]::Round($FileAge.TotalDays,1)) Days]"
			
			$TenantInfoCSVFileIsOld = $True
		}
	}
	
	if (-not $TenantInfoCSVFileExists -or $TenantInfoCSVFileIsOld -or $ForceUpdate) {
		Log -warning "TenantInfoCSVFile Does Not Exist: $TenantInfoCSVFilePath"
		Log -warning "About to sign into CSP Account with MSOL & Az modules"

		$Script:TenantInfoArray = New-Object System.Collections.ArrayList($null)

		$ConnectedToMSOnline = ServiceHandler -SpecifiedServiceName "MSOnline" -ConnectService -CSVUpdate
		
		if ($ConnectedToMSOnline) {
			ServiceHandler -SpecifiedServiceName "AzAccounts" -ConnectService -CSVUpdate
		
			$CSPTenancy = Get-MsolPartnerInformation
			$CSPAzTenant = Get-AzTenant
			$AzSubscriptions = Get-AzSubscription
			$MsolCompanyInfo = Get-MsolCompanyInformation

			Log "Connected to MSOL for $($MsolCompanyInfo.DisplayName)"

			if ( $CSPTenancy.CompanyType -eq "CompanyTenant") {
				Log "Not CSP? Company Type: $($CSPTenancy.CompanyType)"
			}
			else {
				$MsolPartnerList = Get-MsolPartnerContract -ErrorAction SilentlyContinue | select *
				#$AzSubscriptions = Get-AzSubscription 
				if ($MsolPartnerList) {
					
					$CSPTenancyInfo = [PSCustomObject]@{
						IsCSP = $True
						Name = $MsolCompanyInfo.DisplayName
						DefaultDomain = (Get-MsolDomain | Where {$_.IsDefault -eq $True}).Name
						TenantId = $CSPAzTenant.Id
						SharepointAdminURL = (Get-MsolDomain | Where {$_.Name -like "*.onmicrosoft.com"}).Name
						AzSubscriptionName = (Get-AzContext | select -ExpandProperty Name).Split(' ')[0]
						AzSubscriptionId = (Get-AzContext).Name.Split("()")[1]
						
					}

					[void] $TenantInfoArray.Add($CSPTenancyInfo)

					#Test - Check EmailDomain
					$MsolPartnerList | foreach-object {

						$MsolPartner = $_
						$MsolPartnerInfo = [PSCustomObject]@{
							IsCSP = $False
							Name = $MsolPartner.Name
							DefaultDomain = $MsolPartner.DefaultDomainName
							TenantId = $MsolPartner.TenantId
							SharepointAdminURL = "https://" + $MsolPartner.DefaultDomainName.Split(".")[0] + "-admin.sharepoint.com"
							AzSubscriptionName = ""
							AzSubscriptionID = ""

						}
						[void] $TenantInfoArray.Add($MsolPartnerInfo)
					}

					$AzSubscriptions | foreach-object {

						$AzSub = $_
						$AzSubInfo = [PSCustomObject]@{
							IsCSP = $False
							Name = ""
							DefaultDomain = ""
							TenantId = ""
							PartnerContext = ""
							AzSubscriptionName = $AzSub.Name
							AzSubscriptionID = $AzSub.Id
						}
						[void] $TenantInfoArray.Add($AzSubInfo)
					}

					$TenantInfoArray | ConvertTo-Csv -NoTypeInformation | out-file $TenantInfoCSVFilePath

					Log -Success "Updated TenantInfo CSV"
					return $true

				}
				#End IF PARTNER LIST
			}
		}
		else { Log -Warning "TenantInfoCheck Couldn't Connect to MSOL" ; return $False }
	}
	else {
		return $true
	}
}

function DisconnectAll {
		Log "Disconnect All"
		$Servicelist.ModuleName.Replace(".","") | % {		
			Log "[$_] ServiceHandler Disconnecting"
			[void] ( ServiceHandler -SpecifiedServiceName $_ -DisconnectService )
		}
		Log "Disconnected All"
		exit
}

if ($DisconnectAll) {
	DisconnectAll
}

#TODO SS Standalone

function itemizeList {
	param (
		[PSCustomObject]$ListEntry,
		[int]$Position,
		[String[]]$PropertyNames
	)

	if ($PropertyNames) {
		
		if (($PropertyNames).Count -gt 1) {
			$PropertySelectList = New-Object System.Collections.ArrayList($null)
			[Void] $PropertySelectList.Add( @{Label='#'; e={ ( $Position ) }} )

			$PropertyNames | % {
				[Void] $PropertySelectList.Add($_)
			}
		}
		else { $PropertySelectList = $PropertyNames	}

	}

	$ListEntry | select-object -Property $PropertySelectList

}

function SelectFromSearch { 
	param (
		[String[]]$PropertyNames
	)
	$list = @($input)

	$Position = 0
	$niceList = foreach ($ListEntry in $list) {
		itemizeList -ListEntry $ListEntry -Position $Position -PropertyNames $PropertyNames
		$Position++
	}

	$presentResults = $niceList | ft -AutoSize | Out-String
	
	write-host $presentResults
	$Choice = Read-Host "Enter Selection # to return object (0,1,4 for multi) or CTRL+C to quit"
	
	$selected = $list[$Choice]	| Sort
	return $selected

}

#Log "Beginning Actions" 
#TODO This will prevent continuing on with services
if ($Script:Services_Params -and -not $Find) { 
	$Services_Params | % { ServiceHandler -SpecifiedServiceName $_ -ConnectService | out-null
	#Log "If `$Services_Params (Arg input) True and not `$Find (Searchterm arg)" 
	}}
elseif ($Find -and $Script:Services_Params) {
	#Log "Elseif `$Services_Params True and `$Find " 
	Log "Checking if TenantInfo File exists or needs updating" 
	if (CheckTenantInfoCSV) {
		$TenantInfoCSVFileExists = Test-Path $TenantInfoCSVFilePath -ErrorAction SilentlyContinue
		if ($TenantInfoCSVFileExists) {
			Log "Confirmed that TenantInfo File is ready; Getting CSV Contents" 
			$Script:TenantInfoList = ConvertFrom-Csv $( Get-Content $TenantInfoCSVFilePath )
		}
		else { 
			Log -Warning "TenantInfo CSV Is not Ready; Try Reconnecting"
			exit
		}

		#Log "Preparing selected partner from Find input and TenantInfoList Var from CSV File" 

		#Log "Feeding Arg input list of services (as names) to ServiceHandler" 
		$Services_Params | % { 
			if ($_ -eq "AzAccounts") {

				if ($Find) { $Script:AzSubscriptionInfo = $Script:TenantInfoList | Where {$_.AzSubscriptionId -and $_.AzSubscriptionId -ne "" -and $_.IsCSP -ne $True -and $_.AzSubscriptionName -like "*$Find*" -or ($_.AzSubscriptionName).Replace(" ","") -like "*$Find*"} | Sort-Object AzSubscriptionId -Unique } 
				else {
					$Script:AzSubscriptionInfo = $Script:TenantInfoList | Where {$_.AzSubscriptionId -and $_.AzSubscriptionId -ne "" -and $_.IsCSP -ne $True  } | Sort-Object AzSubscriptionName -Unique
				}

				if (($Script:AzSubscriptionInfo).Count -gt 1 ) { 
					$Script:AzSubscriptionInfo = $($AzSubscriptionInfo | sort AzSubscriptionName ) | SelectFromSearch -PropertyNames "AzSubscriptionName", "AzSubscriptionId"
					if (-not $Script:AzSubscriptionInfo) { 
						Log -Warning "No search results..." 
					}
				}
				ServiceHandler -SpecifiedServiceName $_ -ConnectToTenantService -TenantDetails $AzSubscriptionInfo | out-null 
			} else {
				#If first run, ask the question
				if (-not $Script:PartnerTenantInfo) {
					$Script:PartnerTenantInfo = $Script:TenantInfoList | where {$_.Name -like "*$Find*" -or $_.DefaultDomain -like "*$Find*" -or ($_.Name).Replace(" ","") -like "*$Find*"} | Sort-Object Name -Unique

					if (($PartnerTenantInfo).Count -gt 1 ) { 
						$Script:PartnerTenantInfo = $($PartnerTenantInfo | sort Name ) | SelectFromSearch -PropertyNames "DefaultDomain", "Name"
						if (-not $PartnerTenantInfo) { 
							Log -Warning "No search results..." 
						}
					}
				}
				ServiceHandler -SpecifiedServiceName $_ -ConnectToTenantService -TenantDetails $PartnerTenantInfo | out-null 
			}
		}
	}
}
