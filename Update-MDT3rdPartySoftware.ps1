<#
.SYNOPSIS
    Update 3rd party update files in MDT Apllication section
.DESCRIPTION
    Parses third party updates  CliXML list generated from Get-3rdPartySoftware.ps1. Then check MDT for applications similiar to the list using filters. 
.PARAMETER 
    NONE
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -file "Update-MDT3rdPartySoftware.ps1"
.NOTES
    Script name: Update-MDT3rdPartySoftware
    Version:     1.0
    Author:      Richard Tracy
    DateCreated: 2018-11-02
#>

#==================================================
# FUNCTIONS
#==================================================

Function Test-IsISE {
    # try...catch accounts for:
    # Set-StrictMode -Version latest
    try {    
        return ($null -ne $psISE);
    }
    catch {
        return $false;
    }
}

Function Get-ScriptPath {
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


Function Get-SMSTSENV {
    param(
        [switch]$ReturnLogPath,
        [switch]$NoWarning
    )
    
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    }
    Process{
        try{
            # Create an object to access the task sequence environment
            $Script:tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
        }
        catch{
            If(${CmdletName}){$prefix = "${CmdletName} ::" }Else{$prefix = "" }
            If(!$NoWarning){Write-Warning ("{0}Task Sequence environment not detected. Running in stand-alone mode" -f $prefix)}
            
            #set variable to null
            $Script:tsenv = $null
        }
        Finally{
            #set global Logpath
            if ($Script:tsenv){
                #grab the progress UI
                $Script:TSProgressUi = New-Object -ComObject Microsoft.SMS.TSProgressUI

                # Convert all of the variables currently in the environment to PowerShell variables
                $tsenv.GetVariables() | ForEach-Object { Set-Variable -Name "$_" -Value "$($tsenv.Value($_))" }
                
                # Query the environment to get an existing variable
                # Set a variable for the task sequence log path
                
                #Something like: C:\MININT\SMSOSD\OSDLOGS
                #[string]$LogPath = $tsenv.Value("LogPath")
                #Somthing like C:\WINDOWS\CCM\Logs\SMSTSLog
                [string]$LogPath = $tsenv.Value("_SMSTSLogPath")
                
            }
            Else{
                [string]$LogPath = $env:Temp
            }
        }
    }
    End{
        #If output log path if specified , otherwise output ts environment
        If($ReturnLogPath){
            return $LogPath
        }
        Else{
            return $Script:tsenv
        }
    }
}


Function Format-ElapsedTime($ts) {
    $elapsedTime = ""
    if ( $ts.Minutes -gt 0 ){$elapsedTime = [string]::Format( "{0:00} min. {1:00}.{2:00} sec", $ts.Minutes, $ts.Seconds, $ts.Milliseconds / 10 );}
    else{$elapsedTime = [string]::Format( "{0:00}.{1:00} sec", $ts.Seconds, $ts.Milliseconds / 10 );}
    if ($ts.Hours -eq 0 -and $ts.Minutes -eq 0 -and $ts.Seconds -eq 0){$elapsedTime = [string]::Format("{0:00} ms", $ts.Milliseconds);}
    if ($ts.Milliseconds -eq 0){$elapsedTime = [string]::Format("{0} ms", $ts.TotalMilliseconds);}
    return $elapsedTime
}

Function Format-DatePrefix {
    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
    return ($LogDate + " " + $LogTime)
}

Function Write-LogEntry {
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory=$false,Position=2)]
		[string]$Source = '',
        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4)]
        [int16]$Severity,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $Global:LogFilePath,

        [parameter(Mandatory=$false)]
        [switch]$Outhost
    )
    Begin{
        [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
        [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
        [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
        [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias
        
    }
    Process{
        # Get the file name of the source script
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
            Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.Message) -ForegroundColor Red
        }
    }
    End{
        If($Outhost -or $Global:OutTohost){
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
}

Function Show-ProgressStatus {
    <#
    .SYNOPSIS
        Shows task sequence secondary progress of a specific step
    
    .DESCRIPTION
        Adds a second progress bar to the existing Task Sequence Progress UI.
        This progress bar can be updated to allow for a real-time progress of
        a specific task sequence sub-step.
        The Step and Max Step parameters are calculated when passed. This allows
        you to have a "max steps" of 400, and update the step parameter. 100%
        would be achieved when step is 400 and max step is 400. The percentages
        are calculated behind the scenes by the Com Object.
    
    .PARAMETER Message
        The message to display the progress
    .PARAMETER Step
        Integer indicating current step
    .PARAMETER MaxStep
        Integer indicating 100%. A number other than 100 can be used.
    .INPUTS
         - Message: String
         - Step: Long
         - MaxStep: Long
    .OUTPUTS
        None
    .EXAMPLE
        Set's "Custom Step 1" at 30 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 100 -MaxStep 300
    
    .EXAMPLE
        Set's "Custom Step 1" at 50 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 150 -MaxStep 300
    .EXAMPLE
        Set's "Custom Step 1" at 100 percent complete
        Show-ProgressStatus -Message "Running Custom Step 1" -Step 300 -MaxStep 300
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string] $Message,
        [Parameter(Mandatory=$true)]
        [int]$Step,
        [Parameter(Mandatory=$true)]
        [int]$MaxStep,
        [string]$SubMessage,
        [int]$IncrementSteps,
        [switch]$Outhost
    )

    Begin{

        If($SubMessage){
            $StatusMessage = ("{0} [{1}]" -f $Message,$SubMessage)
        }
        Else{
            $StatusMessage = $Message

        }
    }
    Process
    {
        If($Script:tsenv){
            $Script:TSProgressUi.ShowActionProgress(`
                $Script:tsenv.Value("_SMSTSOrgName"),`
                $Script:tsenv.Value("_SMSTSPackageName"),`
                $Script:tsenv.Value("_SMSTSCustomProgressDialogMessage"),`
                $Script:tsenv.Value("_SMSTSCurrentActionName"),`
                [Convert]::ToUInt32($Script:tsenv.Value("_SMSTSNextInstructionPointer")),`
                [Convert]::ToUInt32($Script:tsenv.Value("_SMSTSInstructionTableSize")),`
                $StatusMessage,`
                $Step,`
                $Maxstep)
        }
        Else{
            Write-Progress -Activity "$Message ($Step of $Maxstep)" -Status $StatusMessage -PercentComplete (($Step / $Maxstep) * 100) -id 1
        }
    }
    End{
        Write-LogEntry $Message -Severity 1 -Outhost:$Outhost
    }
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
[string]$RelativeLogPath = Join-Path -Path $scriptDirectory -ChildPath 'Logs'
#Search this path for local MDT installed
[string]$LocalModulePath = "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"

#Set this to somehting different for testing
$ConfigFile = "mdt_configs.s3i.xml"

#build log name
[string]$FileName = $scriptBaseName +'.log'
#build global log fullpath
$Global:LogFilePath = Join-Path $RelativeLogPath -ChildPath $FileName
#clean old log
if(Test-Path $Global:LogFilePath){remove-item -Path $Global:LogFilePath -ErrorAction SilentlyContinue | Out-Null}

Write-Host "Logging to file: $LogFilePath" -ForegroundColor Cyan



# BUILD PATHS FROM XML
#=======================================================
If(Test-Path "$ConfigPath\$ConfigFile"){
    [string]$MDTXMLFile = (Get-Content "$ConfigPath\$ConfigFile" -ReadCount 0) -replace '&','&amp;'
    [xml]$MDTConfigs = $MDTXMLFile

    #get the list of aoftware
    [string]$3rdSoftwareRootPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.rootPath
    [string]$3rdSoftwareListPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.listPath

    If(-not(Test-Path $3rdSoftwareRootPath) -or -not(Test-Path $3rdSoftwareListPath)){
        Write-LogEntry ("Unable to find software repository path or xml file: {0}" -f $3rdSoftwareRootPath) -Severity 3 -Outhost
        Exit
    }
}
Else{
    Write-LogEntry ("Unable to find configuration settings: {0}" -f "$ConfigPath\$ConfigFile") -Severity 3 -Outhost
    Exit
}

##* ==============================
##* MAIN - DO ACTION
##* ==============================
#get the list of MDT servers to update
$MDTServers = $MDTConfigs.mdtConfigs.server


#Loop through each server
Foreach($Server in $MDTServers) {
    
    #Build Vairabls based on XML
    [string]$MDTHost = $Server.Host
    [string]$MDTShare = $Server.share
    
    
    #build mdt path to pull powershell
    $MDTSharePath = "\\" + $MDTHost + "\" + $MDTShare

    #check config if update application is set to true
    [boolean]$UpdateApplications = [boolean]::Parse($Server.updateApplications)

    If(!$UpdateApplications){
        Write-LogEntry ("Configurations is configured NOT to update software on MDT Server: {0}" -f $MDTHost) -Severity 3 -Outhost
        Break
    }


    #check if path exists
    If(Test-Path $MDTSharePath){
        
        Try{
            #check permissions to share
            (Get-Acl $MDTSharePath).Access | ?{$_.IdentityReference -match $User.SamAccountName} | Select IdentityReference,FileSystemRights | Out-Null
            Write-LogEntry ("User has write permission to share, mapping drive: {0}..." -f $MDTSharePath) -Outhost

            #disconnect drive if mapped
            $MappedMDTDrive = Get-PSDrive | Where-Object{$_.Root -eq $MDTSharePath}
            If($MappedMDTDrive){
                Remove-PSDrive -Name $MappedMDTDrive.Name
            }
            #map MDT drive
            New-PSDrive -Name $MapName -Root $MDTSharePath -PSProvider FileSystem | Out-Null
        }
        Catch{
            #if no write ermissions, try mapping drive using credentials
            $MapName = "MDT_" + $MDTShare -replace "[\W]",""
            #check if cred file exists, otherwise prompt for credentials
            If(Test-path  $Server.remoteAuthFile){
                [System.Management.Automation.PSCredential]$MDTCreds = Get-credential (Import-Clixml $Server.remoteAuthFile)
            }
            Else{
                [System.Management.Automation.PSCredential]$MDTCreds = Get-credential
            }

            #map MDT drive using credentials
            New-PSDrive -Name $MapName -Root $MDTSharePath -PSProvider FileSystem -Credential $MDTCreds | Out-Null
            
            #check permissions again
            Try{
                (Get-Acl (Get-PSDrive -Name $MapName).Root).Access | ?{$_.IdentityReference -match $User.SamAccountName} | Select IdentityReference,FileSystemRights | Out-Null
                Write-LogEntry ("User has write permission to share, mapping drive [{0}] using credentials [{1}]" -f $MDTSharePath,$MDTCreds.UserName) -Outhost
            }
            Catch{
                Write-LogEntry ("Write permission to DeploymentShare [{0}] using credentials [{1}] are denied. Provide new mdt credentials to continue." -f $MDTSharePath,$MDTCreds.UserName) -Severity 3 -Outhost
                Break
            }
        }
    }
    Else{
        Write-LogEntry ("Unable to connect to MDT's DeploymentShare: {0}. Check config path." -f $MDTSharePath) -Severity 3 -Outhost
        Break
    }

    ##* ==============================
    ##* MAIN
    ##* ==============================

    #Grab Variables from MDT's Control folder
    If(Test-Path "$MDTSharePath\Control\Settings.xml"){
        $MDTSettings = [Xml] (Get-Content "$MDTSharePath\Control\Settings.xml")
        [string]$MDT_Physical_Path = $MDTSettings.Settings.PhysicalPath
        [string]$MDT_UNC_Path = $MDTSettings.Settings.UNCPath
        

        $MDTAppGroupsFile = [Xml] (Get-Content "$MDTSharePath\Control\ApplicationGroups.xml")
        [xml]$MDTApps = Get-Content "$MDTSharePath\Control\Applications.xml"

        $NewSoftwareList = Import-Clixml $3rdSoftwareListPath

        <#
        Use for Test only
        $NewSoftware = ($NewSoftwareList | Where{($_.Product -match 'Chrome')} | Select -First 2)[1]
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Notepad\+\+')} | Select -First 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Reader DC')} | Select -first 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Reader DC')} | Select -last 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Flash Plugin')} | Select -last 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Java')} | Select -first 1
        $NewSoftware = $NewSoftwareList | Where{($_.Product -match 'Java')} | Select -last 1
        #>

        $UpdatedAppCount = 0
        $ExistingAppCount = 0
        $MissingAppCount = 0

        foreach($NewSoftware in $NewSoftwareList)
        {

            If($NewSoftware.Arch){
                Write-LogEntry ("Working with [{0} {1} ({2}) - {3} bit] in software list" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Version,$NewSoftware.Arch) -Outhost
            }
            Else{
                Write-LogEntry ("Working with [{0} {1} ({2})] in software list" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Version) -Outhost
            }
            # clear the working app variable
            $MDTAppProducts = $null
            $MDTApp = $null

            #Search root path for file name
            $SourceUNCPath = (Get-ChildItem -Path $3rdSoftwareRootPath -Filter $NewSoftware.File -Recurse).FullName
            If($SourceUNCPath.Count -gt 1){
                #Old school way, just in case file names are the same
                #remove parent path of where the software was downloaded to. Attach rootPath from config
                $SplitSoftwarePath = $NewSoftware.FilePath -split "Software" 
                $SourceUNCPath = $3rdSoftwareRootPath + '\Software' + $SplitSoftwarePath[-1]
            }
            

            # find an MDT app that matches the software list base on Publisher, Product Name and Product Type (not always specified)
            #If one is found count is null, but if 2 is found count is 2
            $MDTAppProducts = $MDTApps.applications.application | Where{($_.Publisher -eq $NewSoftware.Publisher) -and ($_.Name -match [regex]::Escape($NewSoftware.Product))}
            
            #if products found equal to 2 or more, filter on product type to reduce it even further
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter1 = $MDTAppProducts | Where {($_.Name -match [regex]::Escape($NewSoftware.ProductType))}
                If($MDTAppFilter1){$MDTAppProducts = $MDTAppFilter1}
            }
    
            #if products found equal to 2 or more, filter on arch match, if specified (names labeled with x64 or x86) to reduce it. 
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter2 = $MDTAppProducts | Where {($_.Name -match $NewSoftware.Arch) -or ($_.ShortName -match $NewSoftware.Arch)}
                If($MDTAppFilter2){$MDTAppProducts = $MDTAppFilter2}
            }

            #if products found equal to 2 or more, filter on arch no match, if NOT specified (usually labeled with x86) to reduce it.
            If($MDTAppProducts.Count -ge 2){
                $MDTAppFilter3 = $MDTAppProducts | Where {($_.Name -notmatch 'x64') -and ($_.ShortName -notmatch 'x64')}
                If($MDTAppFilter3){$MDTAppProducts = $MDTAppFilter3}
            }    

            #lastly just select the first one (hopefully its the correct one)
            $MDTApp = $MDTAppProducts | Select -First 1

            #If and app is found
            If($MDTApp){
                Write-LogEntry ("Filtered Application in MDT to [{0}]" -f $MDTApp.Name) -Outhost

                #remove share from path to get relative path
                $mappedPath = ($MDTApp.WorkingDirectory).Replace('.',$MDTSharePath)

                #' Get current folders
                #' ===================================
                ## Parse working directory and drill only two folders deep.
                ## Anything else deeper doesn't matter because root folder will be deleted if needed
                $CurrentFolders = Get-ChildItem -Path $mappedPath -Recurse -Depth 1 -Force | ?{ $_.PSIsContainer }

            
                $VersionFolderFound = $false
                $SourceFolderFound = $false
                $ConfigFolderFound = $false
                $CurrentVersion = $null
                $ExtraFolders = @()
                $IgnoreFolders = @()
                $KeepFolders = @()

                ## if mutiple folders exist, loop through them to see if its a version folder name.
                ## anything other than the identified folders will be deleted later on
                $CurrentFolders | Foreach-Object {
            
                    switch($_.Name) { 
                        "Source"            {
                                                $SourceFolderFound = $true
                                                Write-Host "Found [Source] folder; " -NoNewline
                                                $KeepFolders += $_
                                            }
                    
                        "Configs"           {
                                                $ConfigFolderFound = $true
                                                Write-Host "Found [Config] folder; " -NoNewline
                                                $IgnoreFolders += $_
                                            }
                    
                        "Updates"           {
                                                $UpdatesFolderFound = $true
                                                Write-Host "Found [Updates] folder; " -NoNewline
                                                $KeepFolders += $_
                                            }

                        "$($NewSoftware.Version)" {
                                                $NewVersion = $NewSoftware.Version
                                                Write-Host "Found [$NewVersion] folder; " -NoNewline
                                                $VersionFolderFound = $true
                                                $KeepFolders += $_
                                            }

                        default             {   $CurrentVersion = $MDTApp.Version
                                                Write-Host "Found [$CurrentVersion] folder; " -NoNewline
                                                $VersionFolderFound = $true
                                                $ExtraFolders += $_
                                            }
                    }
                    write-host ("Categorizing folder [" + $_.Name + "]")
                } #end folder check loop


                #ensure folders are different and deliminated for matching
                $ExtraFolders = ($ExtraFolders | Select -Unique) -join "|"

                #ensure folders are different and deliminated for matching
                $IgnoreFolders = ($IgnoreFolders | Select -Unique) -join "|"

                #Always add the new verison as a keeper
                $KeepFolders += $NewSoftware.Version
                $KeepFolders = ($KeepFolders | Select -Unique) -join "|"

                #' Build folder paths 
                #' ===================================
                ## Does current directory has version folder and source folder
                ## mimic same path structure with new version
                If($sourceFolderFound){$subpath = '\Source\'}Else{$subpath = '\'}
                If($versionFolderFound){$leafpath = $NewSoftware.Version}Else{$leafpath = ''}
                $DestinationPath = ($mappedPath + $subpath + $leafpath)

                #' Compare the versions
                #' ===================================
                # and check to see if file exists
        
                # assume if version match that previous similiar software updated the application. 
                # This issue exists when multiple architecture version exists
                If($MDTApp.Version -eq $NewSoftware.Version){
                    Write-LogEntry ("Application [{0}] version [{1}] was already found in MDT, checking if file exists..." -f $MDTApp.Name,$MDTApp.Version) -Outhost
                    If(-not(Test-Path "$DestinationPath\$($NewSoftware.File)")){

                        ##' If the copy fails return to the stop processing the new software
                        Try{
                            If(Test-Path $SourceUNCPath){
                                Copy-Item $SourceUNCPath -Destination $DestinationPath -Force -PassThru | Out-null
                                Write-LogEntry ("Copied File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -Outhost
                            }
                        }
                        Catch{
                            Write-LogEntry ("Failed to copy File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -Severity 3 -Outhost
                            Return
                        }
                    }
                    Else{
                        Write-LogEntry ("Application [{0}] version [{1}] was already found in MDT" -f $MDTApp.Name,$MDTApp.Version) -Severity 0 -Outhost
                        $ExistingAppCount ++
                    }
                }

                # Update the version
                Else{
            
                    #' Copy new application files
                    #' ================================
                    ##' Do this before deleting the old files just in case. 
                    ##' If the copy fails return to the stop processing the new software
                    New-Item $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                    Try{
                        If(Test-Path $SourceUNCPath){
                            Copy-Item $SourceUNCPath -Destination $DestinationPath -Force -PassThru | Out-null
                            Write-LogEntry ("Copied File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -Severity 0 -Outhost
                        }
                    }
                    Catch{
                        Write-LogEntry ("Failed to copy File [{0}] to [{1}]" -f $NewSoftware.File,$DestinationPath) -Severity 3 -Outhost
                        Return
                    }
            

                    #' Update Script Installer
                    #' =========================
                    $Command = ($MDTApp.CommandLine).split(" ")
                    $CommandUpdated = $false

                    switch($Command[0]){
                    #second update the installer scripts
                        'cscript' { 
                                    Write-LogEntry ("Found a cscript [{0}] for the installer" -f $Command[1]) -Outhost
                                    #grab content from script that installs application
                                    $content = Get-Content "$($mappedPath + '\' + $Command[1])" | Out-String
                                    #find text line that has sVersion
                                    $pattern = 'sVersion\s*=\s*(\"[\w.]+\")'
                                    $content -match $pattern | Out-Null
                                    # if found in cscript installer, update it and save it
                                    If($matches){
                                        $NewContentVer = $content.Replace($matches[1],'"' + $NewSoftware.Version + '"')
        
                                        #add updated version to vbscript
                                        $NewContentVer | Set-Content -Path "$($mappedPath + '\' + $Command[1])" 
                                        Write-LogEntry ("Updated [{0}] variable [sVersion] from [{1}] to [{2}]" -f $Command[1],$matches[1].replace('"',''),$NewSoftware.Version) -Outhost
                                        $CommandUpdated = $true
                                    }
                                    Else{
                                        Write-LogEntry ("Unable to find [sVersion] variable in [{0}], there may be an issue during deployment" -f $Command[1]) -Severity 3 -Outhost
                                    }

                                    #Clear matches
                                    $matches = $null       
                                }


                        '*.exe' {
                                    Write-LogEntry ("Found a executable [{0}] for the installer" -f $Command[1]) -Outhost
                                }

                        'Powershell*' {
                                    Write-LogEntry ("Found a powershell script [{0}] for the installer" -f $Command[1]) -Outhost
                                }

                        'msiexec*' {
                                    Write-LogEntry ("Found a msi file [{0}] for the installer" -f $Command[1]) -Outhost
                                }
                    }


                    If($CommandUpdated){

                        #' Remove old application files
                        #' =================================
                
                        ## Delete any extra folders found (not keep folders and ignore anything in the with a fullpath of ignored folders)
                        Get-ChildItem -Path $mappedPath -Recurse -Depth 1 -Force -Directory | ?{ $_.Name -match $ExtraFolders -and $_.name -notmatch $KeepFolders -and $_.Fullname -notmatch $IgnoreFolders} | ForEach-Object{
                            #if mutiple files exist, loop through them to see if its a version name.
                            Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
                            Write-LogEntry ("Deleted Folder: {0}" -f $_.FullName) -Severity 3 -Outhost
                        }

                        #' Update MDT Listing
                        #' =========================
                        $MDTApp.Version = $NewSoftware.Version
                        Write-LogEntry ("Configured to change MDT's Application [{0}] version property to [{1}]" -f $MDTApp.Name,$NewSoftware.Version) -Severity 0 -Outhost
                        $UpdatedAppCount ++

                    }
                }

                #' Save MDT Listing
                #' =========================
                Try{
                    If($UpdatedAppCount -gt 0){$mdtapps.save("$MDTSharePath\Control\Applications.xml")}
                    Write-LogEntry ("Saved changes to MDT's Application configuration file [{0}] for [{1}]" -f "$MDTSharePath\Control\Applications.xml",$MDTApp.Name) -Severity 0 -Outhost
                    #reset back to 0
                    $UpdatedAppCount = 0
                }
                Catch{
                    Write-LogEntry ("Failed write changes to MDT's Application configuration file [{0}] for [{1}]" -f "$MDTSharePath\Control\Applications.xml",$MDTApp.Name) -Severity 3 -Outhost
                    Return
                }
            }
            Else{
                Write-LogEntry ("Application [{0} {1} ({2})] was not found in MDT" -f $NewSoftware.Publisher,$NewSoftware.Product,$NewSoftware.Arch) -Severity 2 -Outhost
                $MissingAppCount ++
            }

    
        } #end software loop 

        #write out results
        Write-LogEntry ("Updated " + $UpdatedAppCount + " Applications in MDT") -Outhost
        Write-LogEntry ("Found " + $MissingAppCount + " missing Applications in MDT") -Outhost
        Write-LogEntry ("Existing " + $ExistingAppCount + " Applications already up-to-date") -Outhost
        
        #disconnect drive if mapped
        If(Get-PSDrive -Name $MapName -ErrorAction SilentlyContinue){
            Remove-PSDrive -Name $MapName
        }
    }
    Else{
        Write-LogEntry ("Failed write get to MDT's Settings from [{0}]" -f "$MDTSharePath\Control\Settings.xml") -Severity 3 -Outhost
    }
}