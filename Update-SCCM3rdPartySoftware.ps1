<#
.SYNOPSIS
    Update 3rd party update files in MDT Apllication section
.DESCRIPTION
    Parses third party updates  CliXML list generated from Get-3rdPartySoftware.ps1. Then check MDT for applications similiar to the list using filters. 
.PARAMETER 
    NONE
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Update-SCCM3rdPartySoftware.ps1"
.NOTES
    Script name: Update-SCCM3rdPartySoftware
    Version:     1.0
    Author:      Richard Tracy
    DateCreated: 2019-04-22
#>

#==================================================
# FUNCTIONS
#==================================================
Function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}

Function Get-ScriptPath {
    If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }

    # Makes debugging from ISE easier.
    if ($PSScriptRoot -eq "")
    {
        if (Test-IsISE)
        {
            $psISE.CurrentFile.FullPath
            #$root = Split-Path -Parent $psISE.CurrentFile.FullPath
        }
        else
        {
            $context = $psEditor.GetEditorContext()
            $context.CurrentFile.Path
            #$root = Split-Path -Parent $context.CurrentFile.Path
        }
    }
    else
    {
        #$PSScriptRoot
        $PSCommandPath
        #$MyInvocation.MyCommand.Path
    }
}


Function Format-ElapsedTime($ts) {
    $elapsedTime = ""
    if ( $ts.Minutes -gt 0 ){$elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec.", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );}
    else{$elapsedTime = [string]::Format( "{0:00}.{1:00} sec.", $ts.Seconds, $ts.Milliseconds / 10 );}
    if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){$elapsedTime = [string]::Format("{0:00} ms.", $ts.Milliseconds);}
    if ($ts.Milliseconds -eq 0){$elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);}
    return $elapsedTime
}

Function Format-DatePrefix{
    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
    $CombinedDateTime = "$LogDate $LogTime"
    return ($LogDate + " " + $LogTime)
}

Function Write-LogEntry{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false,Position=2)]
		[string]$Source = '',

        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4)]
        [int16]$Severity,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $Global:LogFilePath,

        [parameter(Mandatory=$false)]
        [switch]$Outhost = $Global:OutToHost
    )
    ## Get the name of this function
    [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
	[int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
	[string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
    #  Get the file name of the source script

    Try {
	    If ($script:MyInvocation.Value.ScriptName) {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
	    }
	    Else {
		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
	    }
    }
    Catch {
	    $ScriptSource = ''
    }
    
    
    If(!$Severity){$Severity = 1}
    $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$Severity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"
    
    # Add value to log file
    try {
        Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
    }
    catch {
        Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
    }
    If($Outhost){
        If($Source){
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$Source,$Message)
        }
        Else{
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message)
        }

        Switch($Severity){
            0       {Write-Host $OutputMsg -ForegroundColor Green}
            1       {Write-Host $OutputMsg -ForegroundColor Gray}
            2       {Write-Warning $OutputMsg}
            3       {Write-Host $OutputMsg -ForegroundColor Red}
            4       {If($Global:Verbose){Write-Verbose $OutputMsg}}
            default {Write-Host $OutputMsg}
        }
    }
}


# Function to get properties from an MSI package
function Get-MsiProperty
{
	param(
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        [ValidateNotNullOrEmpty()]
		[string]$Property
	)
	    
	function Get-Property($Object, $PropertyName, [object[]]$ArgumentList)
	{
		return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
	}
	 
	function Invoke-Method($Object, $MethodName, $ArgumentList)
	{
		return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
	}
	 
	$ErrorActionPreference = 'Stop'
	Set-StrictMode -Version Latest
	 
	$msiOpenDatabaseModeReadOnly = 0
	$Installer = New-Object -ComObject WindowsInstaller.Installer
	 
	$Database = Invoke-Method $Installer OpenDatabase @($Path, $msiOpenDatabaseModeReadOnly)
	 
	$View = Invoke-Method $Database OpenView  @("SELECT Value FROM Property WHERE Property=$Property")
	 
	Invoke-Method $View Execute
	 
	$Record = Invoke-Method $View Fetch
	if ($Record)
	{
		Write-Output(Get-Property $Record StringData 1)
	}
	 
	Invoke-Method $View Close @( )
	Remove-Variable -Name Record, View, Database, Installer
	 
}


##*===========================================================================
##* VARIABLES
##*===========================================================================
# Use function to get paths because Powershell ISE and other editors have differnt results
$scriptPath = Get-ScriptPath
[string]$scriptDirectory = Split-Path $scriptPath -Parent
[string]$scriptName = Split-Path $scriptPath -Leaf
[string]$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($scriptName)

#Get required folder and File paths
[string]$ConfigPath = Join-Path -Path $scriptDirectory -ChildPath 'Configs'

[string]$SCCMXMLFile = (Get-Content "$ConfigPath\sccm_configs.s3i.xml" -ReadCount 0) -replace '&','&amp;'
[xml]$SCCMConfigs = $SCCMXMLFile

#get the list of aoftware
[string]$SCCMServer = $SCCMConfigs.SCCMConfigs.server.host

[string]$SCCMRepo = $SCCMConfigs.SCCMConfigs.server.repository
[string]$SCCMLocalRepo = $SCCMConfigs.SCCMConfigs.server.localrepository
[string]$SCCMSiteCode = $SCCMConfigs.SCCMConfigs.server.sitecode

#build sccm source path
$SCCMSharePath = "\\" + $SCCMServer + "\" + $SCCMRepo

#Test directories before continuing
If(-not(Test-Path $SCCMSharePath) ){
    write-host ("Unable to connect to SCCM Source Repository: {0}" -f $SCCMSharePath) -ForegroundColor Red
    Break
}
Else{
    New-PSDrive -Name SCCMREPO -PSProvider FileSystem -Root $SCCMSharePath
    write-host ("Connected to SCCM Source Repository: {0}" -f $SCCMSharePath)
}

#get the list of software
[string]$AppendToName = $SCCMConfigs.SCCMConfigs.softwareBuild.appendName
[string]$3rdSoftwareRootPath = $SCCMConfigs.SCCMConfigs.softwareBuild.software.rootPath -replace '&amp;','&'
[string]$3rdSoftwareListPath = $SCCMConfigs.SCCMConfigs.softwareBuild.software.listPath -replace '&amp;','&'
[string]$ExcludedSoftware = $SCCMConfigs.SCCMConfigs.softwareBuild.ExcludeSoftware -replace ',','|'



If((Test-Path $3rdSoftwareRootPath) -or (Test-Path $3rdSoftwareListPath)){
    $SoftwarePath = New-PSDrive -Name SOFTREPO -PSProvider FileSystem -Root $3rdSoftwareRootPath
}
Else{
    write-host ("Unable to find software repository path or xml file: {0}" -f $3rdSoftwareRootPath) -ForegroundColor Red
    Exit
}

#Import ConfigurationManager powershell Module
If($Env:SMS_ADMIN_UI_PATH){
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Verbose
    #Connect to the site drive if not already mapped
    If(-not(Get-PSDrive -Name $SCCMSiteCode -PSProvider CMSite)){
        New-PSDrive -Name $SCCMSiteCode -PSProvider CMSite -Root $SCCMServer
    }
}
ElseIf($RemoteProvider){
    [System.Management.Automation.PSCredential]$SCCMCreds = Import-Clixml ($scriptDirectory + "\" + $SCCMConfigs.SCCMConfigs.server.remoteAuthFile)

    #remote into MDT server to import MDT module
    Try{
        Enter-PSSession $SCCMServer -Credential $SCCMCreds -EnableNetworkAccess -ErrorAction Stop
        Start-sleep 5
        Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
        if(Get-module *ConfigurationManager*){
            Set-Location -path "$(Get-PSDrive -PSProvider CMSite):\"
        }

        $session = New-PSSession $SCCMServer -Credential $SCCMCreds -EnableNetworkAccess -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock {
            Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
            if(Get-module *ConfigurationManager*){
                Set-Location -path "$(Get-PSDrive -PSProvider CMSite):\"
            }

            $CMsite = (Get-CMSite).SiteCode
            $CMApplications = Get-CMApplication
            $CMPackages = Get-CMPackage
            $CCMTaskSequences = Get-CMTaskSequence


        } -ArgumentList 

    }
    Catch{
        Write-Host "Failed to remote into: $($_.Exception.Message)" -ForegroundColor Red
        Exit
    }
}
Else{
    Write-Host "Unable to load Configuration Managers powershell module" -ForegroundColor Red
    Exit
}


##* ==============================
##* MAIN - DO ACTION
##* ==============================
If(Get-PSDrive -Name $SCCMSiteCode -PSProvider CMSite){
    
    #set the working location to sccm
    Set-Location "$($SCCMSiteCode):\"
    
    #once Connected, lets grab some information
    $CMsite = (Get-CMSite).SiteCode
    $CMApplications = Get-CMApplication
    $CMPackages = Get-CMPackage
    $CCMTaskSequences = Get-CMTaskSequence

    #since mapping ps drive; must grab software list xml from share
    $NewSoftwareList = Import-Clixml "$($SoftwarePath):\softwarelist.xml"
    #$NewSoftwareList = Import-Clixml "$3rdSoftwareListPath"

    #filter list using excludded list
    $FilteredNewSoftList = $NewSoftwareList | Where-Object {$_.Publisher -notmatch $ExcludedSoftware}

    #loop through new software (filtered)
    foreach($NewSoftware in $FilteredNewSoftList)
    {
        #append name to applications to indetify what the script built
        If($AppendToName){$append = " (" + $AppendToName + ")"}Else{$append = ""}
        $GeneratedName = $NewSoftware.Publisher + " " + $NewSoftware.Product + $append
        
        #Check there is already an application found
        If(-not(Compare-object $GeneratedName $CMApplications.LocalizedDisplayName -IncludeEqual | Where{$_.SideIndicator -eq "=="})){

        
        } 

    } # end loop
    
}

#set schedule
New-CMSchedule -Start(Random-StartTime) –RecurInterval Days –RecurCount 1

$AppCollection = Get-CMCollection -Name $CollectionName

$AppCollection = New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $DeviceLimitingCollection -RefreshType Both -RefreshSchedule $Schedule
# If an AD group was specified, add a query membership rule based on that group
Add-CMDeviceCollectionQueryMembershipRule -Collection $AppCollection -QueryExpression "select *  from  SMS_R_System where SMS_R_System.SystemGroupName = ""$DomainNetbiosName\\$ADGroupName""" -RuleName "Members of AD group $ADGroupName"
#Created user collection
$AppCollection = New-CMUserCollection -Name $CollectionName -LimitingCollectionName $UserLimitingCollection -RefreshType Both -RefreshSchedule $Schedule

# Create application
New-CMApplication -Name $ApplicationName -Publisher $Publisher -AutoInstall $true -SoftwareVersion $ApplicationVersion -LocalizedName $TextBoxAppName.Text
# Move the application to folder
$ApplicationObject = Get-CMApplication -Name $ApplicationName
Move-CMObject -FolderPath $ApplicationFolderPath -InputObject $ApplicationObject


# CREATE MSI DEPLOYMENT TYPE
Add-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $MSIFile -LogonRequirementType WhereOrNotUserLoggedOn -Force
# Update the deployment type
$NewDeploymentType = Get-CMDeploymentType -ApplicationName $ApplicationName
# Set the installation program
Set-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -InstallCommand $InstallationProgram
# Set the uninstallation program
Set-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -UninstallCommand $UninstallationProgram
# Set behavior for running installation as 32-bit process on 64-bit systems
Set-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -Force32Bit $true
 # Set the content source path
Set-CMDeploymentType -MsiOrScriptInstaller -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -ContentLocation $ContentSourcePath -WarningAction SilentlyContinue
# Set the option for fallback source location
Set-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -ContentFallback $true
# Set the behavior for clients on slow networks
Set-CMMsiDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -SlowNetworkDeploymentMode Download
# Distribute content to DP group
Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointGroupName $DPGroup
# Deploy the application
Start-CMApplicationDeployment -CollectionName $AppCollection.Name -Name $ApplicationName -DeployPurpose $DeployPurpose


# CREATE MANUAL DEPLOYMENT TYPE
Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -ScriptLanguage PowerShell -ScriptText 'if (Test-Path C:\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}' -InstallationBehaviorType InstallForSystem -UserInteractionMode Normal -LogonRequirementType WhereOrNotUserLoggedOn
# Update the deployment type
$NewDeploymentType = Get-CMDeploymentType -ApplicationName $ApplicationName               