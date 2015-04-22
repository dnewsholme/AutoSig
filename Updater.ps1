<#
.SYNOPSIS
Allows for a script to check for a newer version of itself. If the File has been updated then this will update it and exit. 
This should be the first thing the script does.

.PARAMETER Updateserver_UNC_Path
The folder location of the master copy of the script
Eg. \\server\share\

.PARAMETER Scriptname
The file name of the script
EG. executeorder66.ps1

.EXAMPLE
Start-AutoUpdate -updateserver_UNC_Path "\\537210-pdapi12\packages\scripts\User Creation" -scriptname "UserCreation.ps1"

.NOTES
Daryl Bizsley 2015

#>

#Messagebox Function
Function New-MSGBox ($message){
    #Load .Net Assembly for message box
    [Reflection.Assembly]::LoadFile("C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Windows.Forms.dll") | out-null
    [System.Windows.Forms.MessageBox]::Show("$message")
}

#Declare AutoUpdate Function
Function Start-AutoUpdate([String]$updateserver_UNC_Path, [String]$scriptname) {
    #Declare Variables
    $updateserver_UNC_Path = $updateserver_UNC_Path
    $scriptname = $scriptname
    $currentdir = (Get-Location).Path
    #Check If Hash matches to detect change this requires powershell 3.0
    $master = (Get-fileHash "$updateserver_UNC_Path\$scriptname").Hash
    $current = (Get-fileHash "$currentdir\$scriptname").Hash
    IF ($current -ne $master){
        #Set Update Script path as something powershell will execute if spaces are in the name
        $updatescript = "&('$currentdir\updater.ps1')"
        #Get a copy of the latest script and download it, ready to apply after this script exits.
        Copy-Item "$updateserver_UNC_Path\usercreation.ps1" "$currentdir\$scriptname.update" -confirm:$false
        #Generate the script which will overwrite this one with the latest
        echo "move-item '$currentdir\$scriptname.update' '$currentdir\$scriptname' -Force" | out-file "$currentdir\updater.ps1"
        #New-MSGBox "An Update to this script will now be applied. This Script will now Exit Please Re-Run."
        #Call the update script before killing this one.
        powershell.exe $updatescript
        Exit
    }
}
set-location "$env:USERPROFILE"
Start-AutoUpdate -updateserver_UNC_Path "\\servername\share\" -scriptname "Autosig.ps1"
$args3 = "$env:USERPROFILE\autosig.ps1" 
Start-process "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "& '$args3'"}
