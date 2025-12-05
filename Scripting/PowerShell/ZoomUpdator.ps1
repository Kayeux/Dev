function main(){ #This is our main function for organization
	$ErrorActionPreference = "Stop"
	$Version_Powershell_Major = $PSVERSIONTABLE.PSVersion.Major
	function Error_Handling($text){
	$PSItem.ScriptStackTrace
	$PSItem.Exception.Message
	$ProgressPreference = 'SilentlyContinue'
	}
	function Removing_Old_Zoom(){ #This will uninstall any old Zoom version avoiding side by side issues
		$zip_path = "C:\temp_zoom\CleanZoom.zip"
		$exe_path = "C:\temp_zoom\CleanZoom.exe"
		$Zoom_Uninstaller = "https://support.zoom.us/hc/en-us/article_attachments/360084068792/CleanZoom.zip"
		$Zoom_Uninstaller_Path = "C:\temp_zoom\CleanZoom.zip"
		$command = "invoke-webrequest" + " -uri " + "$Zoom_Uninstaller" + " -OutFile " + "$zip_path"
		Try {
			Invoke-Expression -Command $command | Out-Null
			while (!(Test-Path $Zoom_Uninstaller_Path)) { Start-Sleep 1 }
		}#Downloading the uninstaller
		catch{
			$Error_Text = 'Could not download the un-installer for zoom'
			Error_Handling($Error_Text)
			exit
		}#error if it cant download un-installer
		Expand-Archive -Path "C:\temp_zoom\CleanZoom.zip" -DestinationPath "C:\temp_zoom\"
		Start-Process ".\CleanZoom.exe" -ArgumentList "/silent /keep_outlook_plugin /keep_lync_plugin /keep_notes_plugin" -wait
	}
	function Downloading_Zoom(){
		$Download_URL = "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"
		cd "C:\temp_zoom"
		$command = "invoke-webrequest" + " -uri " + "$Download_URL" + " -OutFile " + "C:\temp_zoom\ZoomInstaller.msi"
		Try { #Now we try running that with first chrome, then the default browser, then it will print the link out if nothing else works
			Invoke-Expression "$command"
			while (!(Test-Path "C:\temp_zoom\ZoomInstaller.msi")) { Start-Sleep 1 }
		}
		catch{
			$Error_Text = 'Could not download the installer for zoom'
			Error_Handling($Error_Text)
			exit
		}
		try{
			Start-Process "./ZoomInstaller.msi" -ArgumentList "/qn /passive /quiet" -Wait
		}
		catch{
			echo 'Zoom could not install'
		}
	}
	$checking = Test-Path "C:\temp_zoom\ZoomInstaller.msi" ###change
	$Start_Dir = $PWD
	$Download_Dir = "C:\temp_zoom\"
	
	cd $Download_Dir
	Try{
		$File_Checker = "C:\temp_zoom\ZoomInstaller.msi"
	}#testing to see if there is any old previously downloaded installers then removes them
	Catch{
		$File_Checker = 'NULL'
	}#This is mainly a placeholder in-case there is no old installer
	if ($checking){
		Removing_Old_Zoom
	}# this uninstalls the old version of zoom IF THE USER HAS IT
	Downloading_Zoom #Here we install zoom
	cd $Start_Dir
}

if ((Test-Path "C:\temp_zoom\") -eq $true){
	rm C:\temp_zoom -Recurse
}

$req_content = (((Invoke-WebRequest -UseBasicParsing "https://apps.apple.com/us/app/zoom-workplace/id546505307").Content | findstr version  | Select-String -Pattern "whats-new__latest__version") | findstr version).Split("Version").Split("<")[-2].Trim()

if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zoom*") -ne "$null"){
	$key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zoom*"
	$installLocation = (Get-ItemProperty -Path $key -Name "DisplayVersion").DisplayVersion
	$zoom_version = $installLocation.Split(" ")[0]
}
elseif (((Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zoom*") -eq "$null") -and ((Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zoom*") -ne "$null")) {
	$key = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Zoom*"
	$installLocation = (Get-ItemProperty -Path $key -Name "DisplayVersion").DisplayVersion
	$zoom_version = $installLocation.Split(" ")[0]
}
elseif ((Test-Path ("C:\Program Files\Zoom\bin\Zoom.exe")) -ne "$null" ){
	$zoom_version = ($((Get-Item "C:\Program Files\Zoom\bin\Zoom.exe").VersionInfo.FileVersion.Split(".")[0..2]) -join ".").Trim()
}
else {
	$zoom_version = "$null"
}

echo $zoom_version 
mkdir "C:\temp_zoom" | Out-Null

if (($req_content -gt $zoom_version) -or ($zoom_version -eq $null)){
	Write-Host "Beginning Install"
	main #main
}
else{
	Write-Host "No Install Required"
	Write-Output "Latest Zoom Version: $req_content"
	Write-Output "Curent Zoom Version: $zoom_version"
}
rm C:\temp_zoom -Recurse
exit #exit upon completion