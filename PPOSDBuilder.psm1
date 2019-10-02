# Function to browse for ISO file to mount, import, and extract.
Function Get-ISOPath
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.Title = "Find and Select Windows Installation ISO"
	$OpenFileDialog.initialDirectory = "C:\"
	$OpenFileDialog.filter = "Disc Image File (*.iso)| *.iso"
	$OpenFileDialog.ShowDialog() | Out-Null
	$Path = $OpenFileDialog.FileName
	return $Path
}

# Function to ensure OSDBuilder version is latest
Function CheckOSDBuilder{

    $Module = "OSDBuilder"
    $ModuleInstalled = Get-Module -ListAvailable | Where {$_.Name -like "*$($Module)"}
    If ($ModuleInstalled){
        $IsImported = Get-Module $Module
        If (-not($IsImported)){
            Import-Module $Module
        }
    }
    Else{
        Install-Module -Name $Module -Force
        Import-Module $Module
    }

    # Update OSD Builder Module if needed
    $InstBuilderVer = (Get-Module "$($Module)").Version
    $GalleryBuilderVer = (Find-Module "$($Module)").Version
    If (-not($InstBuilderVer -eq $GalleryBuilderVer)){
        OSDBuilder -UpdateModule
        Write-Host "   Please exit and reopen PowerShell. Then attempt script run again...script exiting.   " -ForegroundColor Red -BackgroundColor White 
        Break
    }
    Else{
        Write-Host "OSDBuilder Version is current, no need to update...proceeding" -ForegroundColor Green
    }
}

# Function to check if Windows 10 ADK is installed, install if not already
Function Get-ADKInstalled
{
	if ([IntPtr]::Size -eq 4)
	{
		$regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
	}
	else
	{
		$regpath = @(
			'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
			'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
		)
	}
	$Items = Get-ItemProperty $regpath | .{process { if ($_.DisplayName -and $_.UninstallString) { $_ } } } `
	| Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString | Sort DisplayName
	Foreach ($Item in $Items)
	{
        $LatestVer = "10.1.18362.1"
        If (($Item.DisplayVersion -ge $LatestVer) -and ($Item.DisplayName -eq "Windows Assessment and Deployment Kit - Windows 10")){
            $ADKInstalled = $true
        }		
        If (($Item.DisplayVersion -ge $LatestVer) -and ($Item.DisplayName -eq "Windows PE x86 x64")){
            $WinPEInstalled = $true
		}
	}
    Return $ADKInstalled,$WinPEInstalled
}

# Function to browse for 'ADKsetup.exe' if the Windows ADK is not installed
Function Get-ADKSetup
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.Title = "Select 'adksteup.exe'"
	$OpenFileDialog.initialDirectory = "C:\"
	$OpenFileDialog.filter = "ADK Setup exe (adksetup.exe)| adksetup.exe"
	$OpenFileDialog.ShowDialog() | Out-Null
	$ADKSetupPath = $OpenFileDialog.FileName
	return $ADKSetupPath
}

# Function to browse for 'adkwinpesetup.exe' if the Windows ADK is not installed
Function Get-WinPESetup
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.Title = "Select 'adkwinpesetup.exe'"
	$OpenFileDialog.initialDirectory = "C:\"
	$OpenFileDialog.filter = "ADK WinPE Setup exe (adkwinpesetup.exe)| adkwinpesetup.exe"
	$OpenFileDialog.ShowDialog() | Out-Null
	$WinPESetupPath = $OpenFileDialog.FileName
	return $WinPESetupPath
}

# Function to browse for 'OSBuilder' working directory
Function Set-OSDBuilderFolder
{
	$Workfolder = $null
	[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$Browse = New-Object System.Windows.Forms.FolderBrowserDialog
	$Browse.SelectedPath = "C:\"
	$Browse.ShowNewFolderButton = $true
	$Browse.Description = "Select OSDBuilder Working Folder"
	$Loop = $true
	While ($Loop)
	{
		If ($Browse.ShowDialog() -eq "OK")
		{
			$Loop = $false
			$Folder = $browse.SelectedPath
		}
		Else
		{
			$Loop = $false
		}
	}
	$Browse.Dispose()
	Return $Folder
}

Function Invoke-PPOSDBuilder {

    [CmdletBinding(DefaultParametersetname="ImgIndex")]
    Param(
        [Parameter(Mandatory=$true)][ValidateSet('Wks','Svr')][string]$OSType,
        [Parameter(Mandatory=$false,ParameterSetName='ImgIndex')][switch]$ImgIndex,
        [Parameter(Mandatory=$false,ParameterSetName='NoImgIndex')][switch]$NoImgIndex,
        [Parameter(Mandatory=$true,ParameterSetName='NoImgIndex')][string]$IndexNum,
        [Parameter(Mandatory=$true)][string]$BuildTaskName,
        [Parameter(Mandatory=$false)][string]$ImageBuildName = "Win10-x64",
        [Parameter(Mandatory=$false)][switch]$MakeISO
    )

    # ----------------------------------------------------------------------
    # Check if the Windows ADK and the WinPE Add-on are installed. Install it if it is not.
    $PreReqStatus = Get-ADKInstalled
    $ADKPresent = $PreReqStatus[0]
    $WinPEPresent = $PreReqStatus[1]

    If (-not ($ADKPresent))
    {
	    Write-Host " "
	    Write-Host "      Windows 10 Assessment and Deployment Kit is NOT installed!!" -BackgroundColor White -ForegroundColor Red
	    Write-Host "          Windows 10 Assessment and Deployment Kit will be installed. Please wait..." -ForegroundColor Yellow
	    $ADKSetup = Get-ADKSetup
	    If (-not ($ADKSetup))
	    {
		    Write-Host "   'adksetup.exe' NOT selected!" -ForegroundColor Red
		    Break
	    }
	    Else
	    {
		    #$ADKSetup = Get-ADKSetup
		    $ArgList = @(
			    "/features",
			    "OptionId.DeploymentTools",
			    "OptionId.ImagingAndConfigurationDesigner",
			    "OptionId.ICDConfigurationDesigner",
			    "OptionId.UserStateMigrationTool",
			    "/norestart",
			    "/ceip off",
			    "/quiet"
		    )
		    Start-Process -FilePath $ADKSetup -ArgumentList $ArgList -Wait
	    }
	
    }

    If (-not ($WinPEPresent))
    {
	    Write-Host " "
	    Write-Host "      Windows PE ADK Add-On is NOT installed!!" -BackgroundColor White -ForegroundColor Red
	    Write-Host "          Windows PE ADK Add-On will be installed. Please wait..." -ForegroundColor Yellow
	    $WinPESetup = Get-WinPESetup
	    If (-not ($WinPESetup))
	    {
		    Write-Host "   'adkwinpesetup.exe' NOT selected!" -ForegroundColor Red
		    Break
	    }
	    Else
	    {
		    #$WinPESetup = Get-WinPESetup
		    $ArgList = @(
			    "/features",
			    "OptionId.WindowsPreinstallationEnvironment",
			    "/norestart",
			    "/ceip off",
			    "/quiet"
		    )
		    Start-Process -FilePath $WinPESetup -ArgumentList $ArgList -Wait
	    }
	
    }

    # Check if Windows ADK is installed, report if it is.
    If ($ADKPresent)
    {
	    Write-Host " "
	    Write-Host "   Windows 10 Assessment and Deployment Kit is installed..." -ForegroundColor Green
    }

    # Check if Windows PE ADK Add-On is installed, report if it is.
    If ($WinPEPresent)
    {
	    Write-Host " "
	    Write-Host "   Windows PE ADK Add-On is installed..." -ForegroundColor Green
    }

    # Check Version of OSDBuilder
    CheckOSDBuilder

    # Set the Builder Path
    $BuilderPath = Set-OSDBuilderFolder
    If (-not ($BuilderPath)) {
	    Write-Host "ERROR: No working folder specified." -ForegroundColor Red
	    Break
    }
    Else{
        Get-OSDBuilder -SetPath "$($BuilderPath)" -CreatePaths
    }

    # Get and Mount the Windows ISO from Microsoft
    $ISOPath = Get-ISOPath
    If (-not ($ISOPath))
    {
	    Write-Host "   ERROR - Windows installation ISO was NOT selected!!" -ForegroundColor Red
	    Break
    }
    Else
    {
	    Mount-DiskImage -ImagePath $ISOPath -Verbose -ErrorAction Stop
    }

    # Import the OS Media, skip gridview if specified
    If ($NoImgIndex){
        # Import the OS Media by index number, skip showing the gridview
        Import-OSMedia -ImageIndex "$($IndexNum)" -SkipGridView -Verbose
    }

    # Import the OS Media, show gridview if specified
    If ($ImgIndex){
        # Show gridview of available indexes
        Import-OSMedia -Verbose
    }

    # Apply Servicing updates to image (can take up to 2 hours)
    Update-OSMedia -Name "$($(Get-OSMedia).Name)" -Download -Execute

    # Create a new Image Build Task Set
    If ($OSType -eq 'Svr'){
        New-OSBuildTask -TaskName "$($BuildTaskName)" -CustomName "$($BuildTaskName)" -EnableNetFX3 -RemoveAppx -RemoveCapability -RemovePackage -DisableFeature -EnableFeature
    }
    Else {
        New-OSBuildTask -TaskName "$($BuildTaskName)" -EnableNetFX3 -RemoveAppx -RemoveCapability -RemovePackage -DisableFeature -EnableFeature
    }

    # Update the OneDrive Setup EXE
    If ($OSType -eq 'Wks'){
        Get-DownOSDBuilder -ContentDownload "OneDriveSetup Enterprise"
    }

    # Create the Image Build
    If ($OSType -eq 'Svr') {
        New-OSBuild -ByTaskName "$($BuildTaskName)" -Download -Execute -Verbose
    }
    Else {
        New-OSBuild -Name "$($ImageBuildName)" -Download -Execute -ByTaskName "$($BuildTaskName)" -Verbose
    }

    # Get and copy the new WIM File
    $OSBuilds = Get-OSBuilds
    Foreach ($Build in $OSBuilds){
        $NowDateTime = Get-Date -Format "yyyyMMdd_HHmm"
        $CreateWIMDir = New-Item -Path $BuilderPath -Name "$($Build.ImageName) $($Build.Arch) ($($Build.ReleaseId)) - $($NowDateTime)" -ItemType "directory"
        $WIMDir = "$($BuilderPath)\$($Build.ImageName) $($Build.Arch) ($($Build.ReleaseId)) - $($NowDateTime)"
        $BuildPath = $Build.FullName
        Get-ChildItem -Path "$($BuildPath)" -Filter "install.wim" -Recurse | Copy-Item -Destination $WIMDir
        $WIMFile = (Get-ChildItem -Path "$($WIMDir)" -Filter "install.wim" -Recurse).FullName
        Write-Host "  WIM File location is: '$($WIMFile)'" -ForegroundColor Green
        # Create the new ISO if specified
        If ($MakeISO){
            New-OSBMediaISO
        }
    }

    # Dismount Windows ISO
    Dismount-DiskImage -ImagePath $ISOPath

    # ----------------------------------------------------------------------

}