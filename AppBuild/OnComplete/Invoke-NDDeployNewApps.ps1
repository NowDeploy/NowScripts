<#
.SYNOPSIS
    Automated deployment script for new NowDeploy apps in Configuration Manager
.DESCRIPTION
    Use this script as an OnComplete script in NowDeploy to automatically deploy new Configuration Manager applications loaded in by NowDeploy
.NOTES
    Author: Adam Cook (@codaamok)
    To do:
        - Better notification control
        - Support same DeployPurpose as per superseded app, perhaps inc other properties too
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false,Position=0)]
    [string]
    $JsonFile
)

#region User variables - change me
$JsonBackupFolder = "F:\NowDeploy\json" # Every time this script is called, it copies the passed json file to this directory. Comment out to not backup json files
$CMSiteCode = "ABC" # Your site code
$CMSiteServer = "cm.contoso.com" # Your site server FQDN
$DistributionPoints = @("dp1.contoso.com","dp2.contoso.com","dp2.contoso.com") # An array of distribution points to distribute new content to. Value can be "All" to use all DPs in site

$SupersedingAppDeployToLastAppCollections = $true # Deploy new applications to the same collections as of the applications it superseded
$SupersedingAppDeployToCollectionNames = @("CollectionA","CollectionB") # A string or array of collection names to deploy applications that supersede other applications to
$SupersedingAppDeployPurpose = "Available" # Can be Available or Required
$SupersedingAppOverrideServiceWindow = $false # Can be true or false
$SupersedingAppRebootOutsideServiceWindow = $false # Can be true or false
$SupersedingAppUserExperience  = "DisplayAll" # Can be DisplayAll, DisplaySoftwareCenterOnly or HideAll

$NewAppDeployToCollectionNames = @("CollectionA","CollectionB") # A string or array of collection names to deploy new applications to. Comment out to not deploy new applications
$NewAppDeployPurpose = "Available" # Can be Available or Required
$NewAppOverrideServiceWindow = $false # Can be true or false
$NewAppRebootOutsideServiceWindow = $false # Can be true or false
$NewAppUserExperience  = "DisplayAll" # Can be DisplayAll, DisplaySoftwareCenterOnly or HideAll

$MailArgs = @{
    From       = 'example@domain.com'
    To         = 'example@domain.com'
    SmtpServer = 'smtp@domain.com'
    Port       = 587
    UseSsl     = $true
    Credential = (Import-Clixml -LiteralPath "$home\Documents\Keys\pscredential.xml")
}
#endregion

#region Define functions
Function Write-Log {
    [CmdletBinding()]
	Param (
        [string[]]$Message,
        [string]$File,
        [int32]$Level
    )
    $Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    ForEach ($item in $Message) {
        $String = "[{0}] {1}" -f $Time, $item.PadLeft(($item.Length) + ($Level * 4), " ")
        Write-Output $String
        Add-Content -Path $File -Value $String
    }
}

Function Get-CMSupersedeApplications {
    [CmdletBinding()]
    Param (
        # Parameter help description
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string[]]
        $Name
    )
    Begin {
    }
    Process {
        $result = ForEach ($item in $Name) {
            Get-CMDeploymentType -ApplicationName $item | Get-CMDeploymentTypeSupersedence | Get-CMApplication | Select-Object @(
                @{Label="Name";Expression={$_.LocalizedDisplayName}}
                @{Label="DateCreated";Expression={$_.DateCreated}}
                @{Label="IsDeployed";Expression={$_.IsDeployed}}
            )
        }
    }
    End {
        Write-Output ($result | Sort-Object -Unique -Property DateCreated -Descending)
    }
}
#endregion

#region Script variables - do not change me
$OriginalPath = Get-Location | Select-Object -ExpandProperty Path
$logfile = "{0}.log" -f $MyInvocation.MyCommand.Definition
$PSDefaultParameterValues["Write-Log:File"]=$logfile
#endregion

Write-Output ("Using log file: {0}" -f $logfile)
Write-Log -Message "Starting" -Level 0

#region Validate user variables
if ($null -eq $CMSiteCode) {
    $Message = "Please specify your site code within the user variables section of the script"
    Write-Log -Message $Message -Level 0
    throw $Message
}
if ($null -eq $CMSiteServer) {
    $Message = "Please specify your site server FQDN address within the user variables section of the script"
    Write-Log -Message $Message -Level 0
    throw $Message
}
if ($null -eq $DistributionPoints) {
    $Message = "Please specify distribution point(s) within the user variables section of the script"
    Write-Log -Message $Message -Level 0
    throw $Message
}
#endregion

#region Get json content
if ($null -ne $Jsonfile) {
    try {
        $json = Get-Content -LiteralPath $JsonFile -ErrorAction Stop | ConvertFrom-Json
    }
    catch {
        $Message = "Failed to get content from json file {0} ({1})" -f $JsonFile, $error[0].Exception.Message
        Write-Log -Message $Message -Level 0
        throw $Message
    }
    if ($null -ne $JsonBackupFolder) {
        try {
            Copy-Item $JsonFile -Destination $JsonBackupFolder -Force -ErrorAction Stop
        }
        catch {
            $Message = "Failed to back up json file ({0})" -f $error[0].Exception.Message
            Write-Log -Message $Message -Level 0
        }
    }
} else {
    $Message = "No json file given, quiting"
    Write-Log -Message $Message -Level 0
    throw $Message
}
#endregion

#region Import the ConfigurationManager.psd1 module and connect to PS drive
if ($null -eq (Get-Module ConfigurationManager)) {
    try {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    }
    catch {
        $Message = "Failed to import ConfigMgr module, quiting {0}" -f $error[0].Exception.Message
        Write-Log -Message $Message -Level 0
        throw $Message
    }
}
if($null -eq (Get-PSDrive -Name $CMSiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    try {
        New-PSDrive -Name $CMSiteCode -PSProvider CMSite -Root $CMSiteCode -ErrorAction Stop
    }
    catch {
        $Message = "Failed to change PS drive ({0})" -f $error[0].Exception.Message
        Write-Log -Message $Message -Level 0
        throw $Message
    }
}
Set-Location ("{0}:" -f $CMSiteCode)
#endregion

#region Get all DPs if requested
if ($DistributionPoints -eq "All") {
    try {
        $DistributionPoints = Get-CMDistributionPoint -ErrorAction Stop | ForEach-Object { $_.NetworkOSPath -replace "\\" } 
    }
    catch {
        $Message = "Failed to get all DPs ({0})" -f $error[0].Exception.Message
        Write-Log -Message $Message -Level 0
        throw $Message
    }
}
#endregion

#region Log configuration values
[System.Collections.Generic.List[String]]$Messages = @()
switch ($true) {
    $true { # Always log
        $Messages.Add(("Script is configured to connect to site: {0} ({1})" -f $CMSiteServer, $CMSiteCode))
        $Messages.Add(("Script is configured to distribute content to DP(s): {0}" -f [String]::Join(", ", $DistributionPoints)))
    }
    ([String]::IsNullOrEmpty($JsonBackupFolder) -eq $false) {
        $Messages.Add("Script is configured to back up json files to: {0}" -f $JsonBackupFolder)
    }
    $SupersedingAppDeployToLastAppCollections {
        $Messages.Add("For superseding apps, script is configured to deploy to last app's collections")
    }
    ($null -ne $SupersedingAppDeployToCollectionNames) {
        $Messages.Add(("For superseding apps, script is configured to always deploy to collection(s): {0}" -f ([String]::Join(", ", $SupersedingAppDeployToCollectionNames))))
    }
    ($null -ne $NewAppDeployToCollectionNames) {
        $Messages.Add(("For new apps, script is configured to always deploy to collection(s): {0}" -f ([String]::Join(", ", $SupersedingAppDeployToCollectionNames))))
    }
}
Write-Log -Message $Messages -Level 0
#endregion

#region Superseding applications
$Message = "Begin working on superseding apps"
Write-Log -Message $Message -Level 0

# All applications in ConfigMgr marked as superseded by NowDeploy
$CMSupersededApps = $json | Where-Object { $_.Action -eq "Supersede" -And $_.Status -eq "OK" }

# Only the latest applications which are superseded by NowDeploy
# Note: all applications in this chain that are in a supsersedence chain detail one application in the "Comment" property and that's the most recent / latest application
$NewCMSupersededApps = ($CMSupersededApps | Select-Object -ExpandProperty Comment) -replace "Superseded by " | Select-Object -Unique

if ($null -ne $CMSupersededApps -And $CMSupersededApps.count -gt 0) {
    $Message = "Superseded apps: {0}" -f ([String]::Join(", ", $CMSupersededApps.CMAppName))
    Write-Log -Message $Message -Level 1
}

if ($null -ne $NewCMSupersededApps -And $NewCMSupersededApps.count -gt 0) {
    $Message = "Superseding apps: {0}" -f ([String]::Join(", ", $NewCMSupersededApps))
    Write-Log -Message $Message -Level 1
}

if ($SupersedingAppDeployToLastAppCollections -eq $true -Or $null -ne $SupersedingAppDeployToCollectionNames) {

    ForEach ($App in $NewCMSupersededApps) {

        $Message = "Working on: {0}" -f $App
        Write-Log -Message $Message -Level 1

        $Message = "Distributing content"
        Write-Log -Message $Message -Level 2

        try {
            $startCMContentDistributionSplat = @{
                ApplicationName = $App
                DistributionPointName = $DistributionPoints
                ErrorAction = "Stop"
            }
            Start-CMContentDistribution @startCMContentDistributionSplat > $null
        }
        catch {
            # If can't distribut, then we can't deploy, so skip
            $Message = "Failed to distribute, skipping app ({0})" -f $error[0].Exception.Message
            Write-Log -Message $Message -Level 2
            continue
        }

        $Message = "Success"
        Write-Log -Message $Message -Level 2

        # Get N-1 application in a supersedence chain
        # Select -First 1 because it's ordered by DateCreated in desc order and we need the latest from list
        $SupersededCMApp = Get-CMSupersedeApplications -Name $App | Select-Object -First 1

        $Message = "Latest superseded app: {0}" -f $SupersededCMApp.Name
        Write-Log -Message $Message -Level 2

        [System.Collections.Generic.List[String]]$Collections = @()

        if ($SupersedingAppDeployToLastAppCollections -eq $true) {

            if ($SupersededCMApp.IsDeployed -eq $true) {

                $Message = "Getting all collections app is deployed to"
                Write-Log -Message $Message -Level 2

                try {
                    [System.Collections.Generic.List[String]]$Collections = Get-CMApplicationDeployment -Name $SupersededCMApp.Name -ErrorAction Stop | Select-Object -ExpandProperty CollectionName
                }
                catch {
                    $Message = "Failed ({0})" -f $error[0].Exception.Message
                    Write-Log -Message $Message -Level 2
                }

                $Message = "Success"
                Write-Log -Message $Message -Level 2

            } else {

                $Message = "App isn't deployed" -f $App
                Write-Log -Message $Message -Level 2

            }
        }

        # Append to array all collections configured to always deploy superseding apps to
        ForEach ($col in $SupersedingAppDeployToCollectionNames) {
            [System.Collections.Generic.List[String]]$Collections.Add($col)
        }

        $Message = "Collections to deploy app to: {0}" -f ([String]::Join(", ", $Collections))
        Write-Log -Message $Message -Level 2

        ForEach ($Collection in $Collections) {

            $NewCMApplicationDeploymentSplat = @{
                Name = $App
                CollectionName = $Collection
                Comment = "Application supersedes `"{0}`" by NowDeploy on {1}." -f $SupersededCMApp.Name, (Get-Date).ToString()
                DeployPurpose = $SupersedingAppDeployPurpose
                DeployAction = "Install"
                UserNotification = $SupersedingAppUserExperience
                ErrorAction = "Stop"
            }

            # Deploying as Available with these parameters causes warnings printed to console
            if ($SupersedingAppDeployPurpose -eq "Required") {
                $NewCMApplicationDeploymentSplat.Add("OverrideServiceWindow", $SupersedingAppOverrideServiceWindow)
                $NewCMApplicationDeploymentSplat.Add("RebootOutsideServiceWindow", $SupersedingAppRebootOutsideServiceWindow)
            }

            $Message = "Deploying to: {0}" -f $Collection
            Write-Log -Message $Message -Level 2

            try {
                New-CMApplicationDeployment @NewCMApplicationDeploymentSplat > $null
            }
            catch {
                $Message = "Failed ({0})" -f $error[0].Exception.Message
                Write-Log -Message $Message -Level 2
                continue
            }

            $Message = "Success"
            Write-Log -Message $Message -Level 2

        }

    }

    $Message = "No more superseding apps to process"
    Write-Log -Message $Message -Level 1

} else {

    $Message = "Configured to not deploy superseding apps"
    Write-Log -Message $Message -Level 1

}

$Message = "Done working on superseding apps"
Write-Log -Message $Message -Level 0

if ($CMSupersededApps.count -gt 0) {
    $Subject = "[NowDeploy] Supersede: {0} applications" -f $CMSupersededApps.count
    Send-MailMessage @MailArgs -Subject $Subject -Body ($CMSupersededApps | Out-String)
}
#endregion

#region New applications
$Message = "Begin working on new apps"
Write-Log -Message $Message -Level 0

# All newly created applications in ConfigMgr by NowDeploy
$NewCMApps = $json | Where-Object { $_.Action -eq "Create" -And $_.Status -eq "OK" }

if ($null -ne $NewAppDeployToCollectionNames) {

    ForEach ($App in $NewCMApps) {

        # Filter out apps already just deployed from new superseeding apps
        if ($NewCMSupersededApps -notcontains $App.CMAppName) {

            $Message = "Working on: {0}" -f $App.CMAppName
            Write-Log -Message $Message -Level 1

            $Message = "Distributing content"
            Write-Log -Message $Message -Level 2

            try {
                $startCMContentDistributionSplat = @{
                    ApplicationName = $App.CMAppName
                    DistributionPointName = $DistributionPoints
                    ErrorAction = "Stop"
                }
                Start-CMContentDistribution @startCMContentDistributionSplat > $null
            }
            catch {
                # If can't distribut, then we can't deploy, so skip
                $Message = "Failed to distribute, skipping ({0})" -f $error[0].Exception.Message
                Write-Log -Message $Message -Level 2
                continue
            }
            
            $Message = "Success"
            Write-Log -Message $Message -Level 2

            ForEach ($Collection in $NewAppDeployToCollectionNames) {

                $newCMApplicationDeploymentSplat = @{
                    Name = $App.CMAppName
                    CollectionName = $Collection
                    Comment = "Created by NowDeploy on {0}." -f (Get-Date).ToString()
                    DeployPurpose = $NewAppDeployPurpose
                    DeployAction = "Install"
                    UserNotification = $NewAppUserExperience
                    ErrorAction = "Stop"
                }

                # Deploying as Available with these parameters causes warnings printed to console
                if ($NewAppDeployPurpose -eq "Required") {
                    $NewCMApplicationDeploymentSplat.Add("OverrideServiceWindow", $NewAppOverrideServiceWindow)
                    $NewCMApplicationDeploymentSplat.Add("RebootOutsideServiceWindow", $NewAppRebootOutsideServiceWindow)
                }

                $Message = "Deploying to: {0}" -f $Collection
                Write-Log -Message $Message -Level 2

                try {
                    New-CMApplicationDeployment @newCMApplicationDeploymentSplat > $null
                }
                catch {
                    $Message = "Failed ({0})" -f $error[0].Exception.Message
                    Write-Log -Message $Message -Level 2
                    continue
                }

                $Message = "Success"
                Write-Log -Message $Message -Level 2
            
            }

        } else {

            # Skipping superseded app
            continue

        }

    }

    $Message = "No more new apps to process"
    Write-Log -Message $Message -Level 1

} else {

    $Message = "Configured to not deploy new apps"
    Write-Log -Message $Message -Level 1

}

$Message = "Done working on new apps"
Write-Log -Message $Message -Level 0

if ($NewCMApps.count -gt 0) {
    $Subject = "[NowDeploy] New apps: {0} applications" -f $NewCMApps.count
    Send-MailMessage @MailArgs -Subject $Subject -Body ($NewCMApps | Out-String)
}
#endregion

#region Failed items
$Failures = $json | Where-Object { $_.Status -eq "Fail" }
if ($Failures.count -gt 0) {
    $Subject = "[NowDeploy] Failure: {0} applications" -f $Failures.count
    Send-MailMessage @MailArgs -Subject $Subject -Body ($Failures | Out-String)
}
#endregion

Set-Location $OriginalPath

Write-Log -Message "Finished" -Level 0

Start-Sleep -Seconds 60
