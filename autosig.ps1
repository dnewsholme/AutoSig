1	+$result = Get-Item "$Env:appdata\Microsoft\Signatures\$env:username.htm" -ErrorAction SilentlyContinue
2	+$currentversionnumber = "V1.01"
3	+function signature {
4	+#get AD info
5	+$banner = "http:/companywebsite.co.uk/signature.jpg" 
6	+
7	+$strName = $env:username
8	+
9	+$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"
10	+
11	+$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
12	+$objSearcher.Filter = $strFilter
13	+
14	+$objPath = $objSearcher.FindOne()
15	+$objUser = $objPath.GetDirectoryEntry()
16	+
17	+
18	+$strName = $objUser.FullName
19	+$strTitle = $objUser.Title
20	+$strCompany = $objUser.Company
21	+$strCred = $objUser.info
22	+$strStreet = $objUser.StreetAddress
23	+$strPhone = $objUser.telephoneNumber
24	+$strCity =  $objUser.l
25	+$strPostCode = $objUser.PostalCode
26	+$strCountry = $objUser.co
27	+$strEmail = $objUser.mail
28	+$strWebsite = $objUser.wWWHomePage
29	+$strddi = $objUser.otherTelephone
30	+$strmobile = $objUser.mobile
31	+$switchboard = 08450000000
32	+
33	+
34	+$UserDataPath = $Env:appdata
35	+$FolderLocation = $UserDataPath + '\\Microsoft\\Signatures'  
36	+mkdir $FolderLocation -force
37	+#create html file for outlook
38	+$stream = [System.IO.StreamWriter] "$FolderLocation\\$env:username.htm"
39	+$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
40	+$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
41	+$stream.WriteLine("<DIV align=left><FONT face=Tahoma><STRONG>$strname</STRONG></FONT></DIV>")
42	+$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strtitle</FONT></DIV>")
43	+$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma><a href='mailto:$strEmail'>$strEmail</a></FONT></DIV>")
44	+$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$switchboard Ext: $strPhone</FONT></DIV>")
45	+$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$strddi</FONT></DIV>")
46	+$stream.Writeline("<DIV align=left><FONT size=2 face=Tahoma>$strmobile</FONT></DIV>")
47	+$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strcompany</FONT></DIV>")
48	+$stream.WriteLine("<DIV align=left><FONT size=2 face=Tahoma>$strstreet  $strCity  $strPostCode</FONT></DIV>")
49	+$stream.WriteLine("<DIV align=left><A href='$strWebsite'></A></DIV>")
50	+$stream.WriteLine("<DIV align=left><A href='$banner'></A><A 
51	+href='$banner'></A><IMG style='MARGIN: 0px' 
52	+border=0 alt='' src='$banner' height='109' width='400'><A 
53	+href='$banner'></A></DIV>")
54	+
55	+
56	+$stream.WriteLine("</BODY>")
57	+$stream.WriteLine("</HTML>")
58	+$stream.close()
59	+
60	+
61	+#Force To show in outlook and set defaults
62	+$MSWord = New-Object -com word.application 
63	+$EmailOptions = $MSWord.EmailOptions 
64	+$EmailSignature = $EmailOptions.EmailSignature 
65	+$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
66	+$EmailSignature.NewMessageSignature = $env:username
67	+$MSWord.Quit()
68	+$MSWord = New-Object -com word.application 
69	+$EmailOptions = $MSWord.EmailOptions 
70	+$EmailSignature = $EmailOptions.EmailSignature 
71	+$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
72	+#$EmailSignature.ReplyMessageSignature = $env:username
73	+$MSWord.Quit()  
74	+$currentversionnumber | Out-File "$env:userprofile\SigVersion.log" 
75	+}
76	+if ($result -eq $null) {
77	+signature
78	+}
79	+Else {$version = Get-Content "$env:userprofile\SigVersion.log" -ErrorAction SilentlyContinue
80	+if ($version -ne $currentversionnumber) { 
81	+signature 
82	+}
83	+
84	+else{exit}
85	+}
86	\ No newline at end of file
