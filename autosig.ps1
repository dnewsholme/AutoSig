$result = Get-Item "$Env:appdata\Microsoft\Signatures\$env:username.htm" -ErrorAction SilentlyContinue
$currentversionnumber = "V1.01"
function signature {
#get AD info
$banner = "http:/companywebsite.co.uk/signature.jpg" 

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
$switchboard = 08450000000


$UserDataPath = $Env:appdata
$FolderLocation = $UserDataPath + '\\Microsoft\\Signatures'  
mkdir $FolderLocation -force
#create html file for outlook
$stream = [System.IO.StreamWriter] "$FolderLocation\\$env:username.htm"
$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
$stream.WriteLine("<DIV align=left><FONT face=Tahoma><STRONG>$strname</STRONG></FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strtitle</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma><a href='mailto:$strEmail'>$strEmail</a></FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$switchboard Ext: $strPhone</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$strddi</FONT></DIV>")
$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$strmobile</FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strcompany</FONT></DIV>")
$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strstreet  $strCity  $strPostCode</FONT></DIV>")
$stream.WriteLine("<DIV align=left><A href='$strWebsite'></A></DIV>")
$stream.WriteLine("<DIV align=left><A href='$banner'></A><A 
href='$banner'></A><IMG style='MARGIN: 0px' 
border=0 alt='' src='$banner' height='109' width='400'><A 
href='$banner'></A></DIV>")


$stream.WriteLine("</BODY>")
$stream.WriteLine("</HTML>")
$stream.close()


#Force To show in outlook and set defaults
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
