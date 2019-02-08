<#
.SYNOPSIS
    Update 3rd party update files in MDT Apllication section
.DESCRIPTION
    Parses third party updates  CliXML list generated from Get-3rdPartySoftware.ps1. Then check MDT for applications similiar to the list using filters. 
    Current only supports cscript installer
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


function Test-IsISE {
# try...catch accounts for:
# Set-StrictMode -Version latest
    try {    
        return $psISE -ne $null;
    }
    catch {
        return $false;
    }
}        
        
Function Import-SMSTSENV{
    try{
        $Global:tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
        Write-Output "$ScriptName - tsenv is $tsenv "
        $MDTIntegration = $true
        
        #$tsenv.GetVariables() | % { Write-Output "$ScriptName - $_ = $($tsenv.Value($_))" }
    }
    catch{
        Write-Output "$ScriptName - Unable to load Microsoft.SMS.TSEnvironment"
        Write-Output "$ScriptName - Running in standalonemode"
        $MDTIntegration = $false
    }
    Finally{
        if ($MDTIntegration){
            $Logpath = $tsenv.Value("LogPath")
            $LogFile = $Logpath + "\" + "$LogName.log"

        }
        Else{
            $Logpath = $env:TEMP
            $LogFile = $Logpath + "\" + "$LogName.log"
        }
    }
    
    #Start Transcript Logging
    Start-Transcript -path $LogFile -Force
}



Function logstamp {
    $now=get-Date
    $yr=$now.Year.ToString()
    $mo=$now.Month.ToString()
    $dy=$now.Day.ToString()
    $hr=$now.Hour.ToString()
    $mi=$now.Minute.ToString()
    if ($mo.length -lt 2) {
    $mo="0"+$mo #pad single digit months with leading zero
    }
    if ($dy.length -lt 2) {
    $dy ="0"+$dy #pad single digit day with leading zero
    }
    if ($hr.length -lt 2) {
    $hr ="0"+$hr #pad single digit hour with leading zero
    }
    if ($mi.length -lt 2) {
    $mi ="0"+$mi #pad single digit minute with leading zero
    }

    write-output $yr$mo$dy$hr$mi
}

Function Write-Log{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring -Force
}


##* ==============================
##* VARIABLES
##* ==============================
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
#Since running script within Powershell ISE doesn't have a $scriptpath...hardcode it
If(Test-IsISE){$scriptPath = "C:\GitHub\Get3rdPartySoftware\Update-MDT3rdPartySoftware.ps1"}
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

#Get required folder and File paths
[string]$ConfigPath = Join-Path -Path $scriptRoot -ChildPath 'Configs'
[string]$ModulesPath = Join-Path -Path "$MDTSharePath\Tools" -ChildPath 'Modules'

[string]$MDTXMLFile = (Get-Content "$ConfigPath\mdt_configs.xml" -ReadCount 0) -replace '&','&amp;'
[xml]$MDTConfigs = $MDTXMLFile
[string]$MDTHost = $MDTConfigs.mdtConfigs.server.host
[string]$MDTShare = $MDTConfigs.mdtConfigs.server.share
[string]$MDTPhysicalPath = $MDTConfigs.mdtConfigs.server.PhysicalPath
[string]$3rdSoftwareRootPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.rootPath
[string]$3rdSoftwareListPath = $MDTConfigs.mdtConfigs.softwareListCliXml.software.listPath
[boolean]$RemoteProvider = [boolean]::Parse($MDTConfigs.mdtConfigs.server.remoteMDTProvider)
If($RemoteProvider){
    [System.Management.Automation.PSCredential]$MDTCreds = Import-Clixml ($scriptRoot + "\" + $MDTConfigs.mdtConfigs.server.remoteAuthFile)
}
$MDTSharePath = "\\" + $MDTHost + "\" + $MDTShare

$MDTModulePath = Test-Path "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"
##* ==============================
##* MODULE/EXTENSIONS
##* ==============================
#Try to Import SMSTSEnv from MDT server
#Import-SMSTSENV

#import TaskSequence Module
#Import-Module $ModulesPath\ZTIUtility -ErrorAction SilentlyContinue

If($MDTModulePath){
    Import-Module "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"
    #map to mdt drive (must have MDT module loaded)
    $MDTPSProvider = Get-PSProvider -PSProvider MDTProvider -ErrorAction SilentlyContinue
    If(!$MDTPSProvider){
        Try{
            $MDTDrive = New-PSDrive -Name DS001 -PSProvider mdtprovider -Root $MDTSharePath
            $MDTModuleLoaded = $true 

        }Catch{
            Write-Host "Failed to load MDT module; Please try enabled remoteconfig" -ForegroundColor Red
            $MDTModuleLoaded = $false
        }
    }
} Else{
    Write-Host "No MDT module is installed. Please go to [https://www.microsoft.com/en-us/download/details.aspx?id=54259]"
}


If($RemoteProvider -and !$MDTModuleLoaded){
    Try{
        $Session = New-PSSession -ComputerName $MDTHost -Credential $MDTCreds
        Invoke-Command -Session $Session -Script {
            Import-Module "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"; 
            $Drive = New-PSDrive -Name DS001 -PSProvider mdtprovider -Root $args[0]
            cd $Drive.Root 
        } -Args ($MDTConfigs.mdtConfigs.server.PhysicalPath)
        #Enter-PSSession -Session $Session
    }
    Catch{
        Write-Host "Failed to remote into: $($_.Exception.Message)" -ForegroundColor Red
         Break
    }
}




##* ==============================
##* MAIN
##* ==============================

#Grab Variables from MDT's Control folder
$MDTSettings = [Xml] (Get-Content "$MDTSharePath\Control\Settings.xml")
[string]$MDT_Physical_Path = $MDTSettings.Settings.PhysicalPath
[string]$MDT_UNC_Path = $MDTSettings.Settings.UNCPath


$MDTAppGroupsFile = [Xml] (Get-Content "$MDTSharePath\Control\ApplicationGroups.xml")
[xml]$MDTApps = Get-Content "$MDTSharePath\Control\Applications.xml" -Credential $MDTCreds

$SoftwareList = Import-Clixml $3rdSoftwareListPath

<#Test only
$Software = $SoftwareList | Where{($_.Product -match 'Chrome')} | Select -First 1
$Software = $SoftwareList | Where{($_.Product -match 'Notepad\+\+')} | Select -First 1
$Software = $SoftwareList | Where{($_.Product -match 'Reader DC')} | Select -first 1
$Software = $SoftwareList | Where{($_.Product -match 'Reader DC')} | Select -last 1
$Software = $SoftwareList | Where{($_.Product -match 'Flash Plugin')} | Select -last 1
$Software = $SoftwareList | Where{($_.Product -match 'Java')} | Select -first 1
$Software = $SoftwareList | Where{($_.Product -match 'Java')} | Select -last 1
#>

$UpdatedAppCount = 0
$ExistingAppCount = 0
$MissingAppCount = 0

foreach($Software in $SoftwareList)
{
    If($Software.Arch){
        Write-Host ("Found [{0} {1} ({2}) - {3} bit] in software list" -f $Software.Publisher,$Software.Product,$Software.Version,$Software.Arch) -ForegroundColor Cyan
    }
    Else{
        Write-Host ("Found [{0} {1} ({2})] in software list" -f $Software.Publisher,$Software.Product,$Software.Version) -ForegroundColor Cyan
    }
    # clear the working app variable
    $MDTAppProducts = $null
    $MDTApp = $null

    #remove parent path of where the software was downloaded to. Attach rootPath from config
    $SplitSoftwarePath = $Software.FilePath -split "Software" 
    $NewUNCPath = $3rdSoftwareRootPath + '\Software' + $SplitSoftwarePath[-1]
    # find an MDT app that matches the software list base on Publisher, Product Name and Product Type (not always specified)
    $MDTAppProducts = $MDTApps.applications.application | Where{($_.Publisher -eq $Software.Publisher) -and ($_.Name -match [regex]::Escape($Software.Product))}
    
    #if more than 1 are found, filter on product type to reduce it
    If($MDTAppProducts.Count -ge 2){
        $MDTAppFilter1 = $MDTAppProducts | Where {($_.Name -match [regex]::Escape($Software.ProductType))}
        If($MDTAppFilter1){$MDTAppProducts = $MDTAppFilter1}
    }
    
    #if more than 1 are found, filter on arch match, if specified (names labeled with x64 or x86) to reduce it. 
    If($MDTAppProducts.Count -ge 2){
        $MDTAppFilter2 = $MDTAppProducts | Where {($_.Name -match $Software.Arch) -or ($_.ShortName -match $Software.Arch)}
        If($MDTAppFilter2){$MDTAppProducts = $MDTAppFilter2}
    }

    #if more than 1 are found, filter on arch no match, if NOT specified (usually labeled with x86) to reduce it.
    If($MDTAppProducts.Count -ge 2){
        $MDTAppFilter3 = $MDTAppProducts | Where {($_.Name -notmatch 'x64') -and ($_.ShortName -notmatch 'x64')}
        If($MDTAppFilter3){$MDTAppProducts = $MDTAppFilter3}
    }    

    $MDTApp = $MDTAppProducts | Select -First 1
    Write-Host ("Filtered Application in MDT to [{0}]" -f $MDTApp.Name) -ForegroundColor DarkYellow

    #If and app is found
    If($MDTApp){

        #check the version and update it
        If($MDTApp.Version -ne $Software.Version){
            $mappedPath = ($MDTApp.WorkingDirectory).Replace('.',$MDTSharePath)
        
            #[Version]$Version = $MDTApp.Version

            #' Remove current application folders
            #' ===================================
            ## Parse working directory and drill only two folders deep.
            ## Anything else deeper doesn't matter because root folder will be deleted if needed
            $CurrentFolders = Get-ChildItem -Path $mappedPath -Recurse -Depth 1 -Force | ?{ $_.PSIsContainer }
        
            #if mutiple folders exist, loop through them to see if its a version folder name.
            # delete them
            $versionFolderFound = $false
            $sourceFolderFound = $false
            Foreach ($folder in $CurrentFolders){
                # Be sure not to delete the Source folder
                $LastFolder = Split-Path $folder -Leaf
                If($LastFolder -ne 'Source'){
                    If( ($LastFolder -match $MDTApp.Version) ){
                        Remove-Item $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
                        Write-Host ("Deleted version folder [{0}]" -f $LastFolder) -ForegroundColor Red
                        $versionFolderFound = $true
                    }Else{
                        Write-Host ("No version folder [{0}]" -f $LastFolder) -ForegroundColor DarkGray
                    }
                }Else{
                    Write-Host ("Source folder found [{0}]" -f $LastFolder) -ForegroundColor DarkGray
                    $sourceFolderFound = $true
                }
            }


            #' Remove current application files
            #' =================================
            ## Parse working directory for matching file types now that folders are removed
            $CurrentFiles = Get-ChildItem -Path $mappedPath -Filter "*$($Software.FileType)" -Recurse -Force

            #if mutiple files exist, loop through them to see if its a version name.
            # delete them
            Foreach($File in $CurrentFiles){
                #find if file name has the product type in it
                If($File.Name -match [regex]::Escape($Software.ProductType)){
                    Remove-Item $File.FullName -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
                    Write-Host ("Deleted file: {0}" -f $File.Name) -ForegroundColor Gray
                }
            }


            #' Copy new application files
            #' ================================
            If($sourceFolderFound){$subpath = '\Source\'}Else{$subpath = '\'}
            If($versionFolderFound){$leafpath = $Software.Version}Else{$leafpath = ''}
            $DestinationPath = ($mappedPath + $subpath + $leafpath)

            New-Item $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

            Copy-Item $NewUNCPath -Destination $DestinationPath -Force | Out-null

            <#
            If($versionFolderFound){
                 New-Item ($mappedPath + $subpath + $Software.Version) -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                 If(Test-path ($mappedPath + $subpath + $Software.Version)){$DestinationPath = $mappedPath + "\Source\$($Software.Version)"}Else{$DestinationPath = $mappedPath}
                 Write-host ("New folder [{0}] was created at [{0}]" -f $Software.Version,$DestinationPath)
            }
            Else{
                If(Test-path "$mappedPath\Source"){$DestinationPath = $mappedPath + "\Source"}Else{$DestinationPath = $mappedPath}
                Write-host ("Exiting path [{0}] will be used " -f $DestinationPath)
            }  
            
            Copy-Item $NewUNCPath -Destination $DestinationPath -Force | Out-null
            #
            #>            Write-Host ("Copied File [{0}] to [{1}]" -f $Software.File,$DestinationPath) -ForegroundColor Green


            #' Update Script Installer
            #' =========================
            $Command = ($MDTApp.CommandLine).split(" ")

            #second update the installer scripts
            If($Command[0] -eq 'cscript') {
                Write-Host ("Found a cscript [{0}] for the installer" -f $Command[1]) -ForegroundColor Gray
                #grab content from script that installs application
                $content = Get-Content "$($mappedPath + '\' + $Command[1])" | Out-String
                #find text line that has sVersion
                $pattern = 'sVersion\s*=\s*(\"[\w.]+\")'
                $content -match $pattern | Out-Null
                # if found in cscript installer, update it and save it
                If($matches){
                    $NewContentVer = $content.Replace($matches[1],'"' + $Software.Version + '"')
        
                    #add updated version to vbscript
                    $NewContentVer | Set-Content -Path "$($mappedPath + '\' + $Command[1])" 
                    Write-Host ("Updated [{0}] variable [sVersion] from [{1}] to [{2}]" -f $Command[1],$matches[1].replace('"',''),$Software.Version) -ForegroundColor DarkYellow
                }Else{
                    Write-Host ("Unable to find [sVersion] variable in [{0}], there may be an issue during deployment" -f $Command[1]) -ForegroundColor Red
                }

                #Clear matches
                $matches = $null           
            }


            #' Update MDT Listing
            #' =========================
            $MDTApp.Version = $Software.Version
            Write-Host ("Configured to change MDT's Application [{0}] version property to [{1}]" -f $MDTApp.Name,$Software.Version) -ForegroundColor DarkGreen
            $UpdatedAppCount ++
        }
        Else{
            Write-Host ("Application [{0}] version [{1}] was already found in MDT" -f $MDTApp.Name,$MDTApp.Version) -ForegroundColor Green
        }
    }
    Else{
        Write-Host ("Application [{0} {1} ({2})] was not found in MDT" -f $Software.Publisher,$Software.Product,$Software.Arch) -ForegroundColor Yellow
        $MissingAppCount ++
    }
    
} 


#' Save MDT Listing
#' =========================
If($UpdatedAppCount -gt 0){$mdtapps.save("$MDTSharePath\Control\Applications.xml")}

Write-Host ("Saved MDT Application configurations [{0}]" -f "$MDTSharePath\Control\Applications.xml") -ForegroundColor Green