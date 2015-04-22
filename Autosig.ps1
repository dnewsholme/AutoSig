$result = Get-Item "$Env:appdata\Microsoft\Signatures\$env:username.htm" -ErrorAction SilentlyContinue
$currentversionnumber = "V2.8"
function signature {

$banner = "http://www.servername.co.uk/downloads/sig.jpg" 

$strName = $env:username

$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objUser = $objPath.GetDirectoryEntry()


$strName = $objUser.FullName
$strTitle = $objUser.Title
$strCompany = $objUser.Company
$strCred = $objUser.info
$strStreet = $objUser.StreetAddress
$strPhone = $objUser.telephoneNumber
$strCity =  $objUser.l
$strPostCode = $objUser.PostalCode
$strCountry = $objUser.co
$strEmail = $objUser.mail
$strWebsite = $objUser.wWWHomePage
$strddi = $objUser.otherTelephone
$strmobile = $objUser.mobile
if ($strddi -like $null){
$ddi =$null}
else {$ddi = "Direct Dial:"}

if ($strmobile -like $null){
$mobile =$null}
else {$mobile = "Mobile:"}


$UserDataPath = $Env:appdata
#if (test-path "HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General") {
  #get-item -path HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force}
  

#if (test-path "HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General") {
  #get-item -path HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force}
#if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
  #get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force}
$FolderLocation = $UserDataPath + '\\Microsoft\\Signatures'  
mkdir $FolderLocation -force

$stream = [System.IO.StreamWriter] "$FolderLocation\\$env:username.htm"
$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
$stream.WriteLine("<DIV align=left><FONT face=Tahoma><STRONG>$strname</STRONG></FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strtitle</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma><a href='mailto:$strEmail'>$strEmail</a></FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>08453009410 Ext: $strPhone</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$ddi $strddi</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$mobile $strmobile</FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strcompany</FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strstreet  $strCity  $strPostCode</FONT></DIV>")
$stream.WriteLine("<DIV align=left><A href='www.lowellgroup.co.uk'></A></DIV>")
$stream.WriteLine("<DIV align=left><A href='$banner'></A><A 
href='$banner'></A><IMG style='MARGIN: 0px' 
border=0 alt='' src='$banner' height='109' width='697'><A 
href='$banner'></A></DIV>")


$stream.WriteLine("</BODY>")
$stream.WriteLine("</HTML>")
$stream.close()

### Add Signature into outlook settings##
#$sig = "$env:username.htm"
#New-PSDrive reg -PSProvider Registry -Root HKEY_CURRENT_USER\Software\Microsoft\Office
#Set-Location reg: -ErrorAction SilentlyContinue
#$officeversions = Get-Item * | where {$_.name -Like "*.0"}
#foreach ($officeversion in $officeversions){
#Remove-Item -Path REG:$officeversion\Outlook\Setup\
#New-ItemProperty -PropertyType ExpandString -Path REG:\$officeversion\Common\Mailsettings\ -Name NewSignature -Value $sig -ErrorAction SilentlyContinue
#New-ItemProperty -PropertyType ExpandString -Path REG:\$officeversion\Common\Mailsettings\ -Name ReplySignature -Value $sig -ErrorAction SilentlyContinue
#New-ItemProperty -PropertyType ExpandString -Path REG:\$officeversion\Common\General\ -Name Signatures -Value "Signature" -ErrorAction SilentlyContinue}
#forcing set to reply/forward

$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.NewMessageSignature = $env:username
$MSWord.Quit()
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
#$EmailSignature.ReplyMessageSignature = $env:username
$MSWord.Quit()  
$currentversionnumber | Out-File "$env:userprofile\SigVersion.log" 
}
if ($result -eq $null) {
signature
}
Else {$version = Get-Content "$env:userprofile\SigVersion.log" -ErrorAction SilentlyContinue
if ($version -ne $currentversionnumber) { 
signature 
}

else{exit}
}
