mkdir C:\NinjaInstall
Invoke-WebRequest "https://app.ninjarmm.com/agent/installer/5f62e7be-7672-4d11-a421-6c730b918cc2/downholechemicalsolutionsmainoffice-7.0.2317-windows-installer.msi" -OutFile C:\NinjaInstall\Ninja.msi
Start-Process msiexec.exe -Wait -ArgumentList '/i "C:\NinjaInstall\Ninja.msi" /qn'