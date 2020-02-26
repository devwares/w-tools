
function InstallChrome {
    Param(
        [parameter(Mandatory=$True)][string] $TempDir
    )    

    # Installation Chrome
    $Installer = "chrome_installer.exe"; Invoke-WebRequest "http://dl.google.com/chrome/install/375.126/chrome_installer.exe" -OutFile $TempDir\$Installer; Start-Process -FilePath $TempDir\$Installer -Args "/silent /install" -Verb RunAs -Wait; Remove-Item $TempDir\$Installer
}
