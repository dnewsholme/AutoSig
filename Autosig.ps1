#Function to automatically create a signature from AD attributes.
function Update-Signature {
    #region Configurable PARAMETERS
    #Company Main Contact Number
    $companyphone = "+44 (0)00000000"
    #Set image. Dimensions can be set via the two variables $imgwidth and $imgheight
    $banner = "http://someurl/companyimage.png"
    #Sets the image height and width
    $imgheight = 150
    $imgwidth = 300
    #endregion

    #Get username
    $strName =  $env:username
    #look up AD attributes and set variables for later use.
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
    #check if ddi exists if it does add to signature, otherwise remove block from signature.
    if ($strddi -ne $null){
    $ddi = @"
    <p class=MsoNormal><b><span
    style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;color:#023a98;LINE-HEIGHT:0pt'>T </span></b><span style='font-size:10.0pt;
    font-family:"Segoe UI",sans-serif;
    color:black'>&nbsp;</span><span style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;
    color:#000000;LINE-HEIGHT:0pt'>$strddi</span></p>
"@
    }
    #check if mobile exists and add to signature otherwise don't add.
    if ($strmobile -ne $null){
      $mobile = @"
      <p class=MsoNormal><b><span
      style='font-size:10pt;font-family:"Segoe UI",sans-serif;color:#023a98;LINE-HEIGHT:0pt'>M</span></b><span style='font-size:10.0pt;
      font-family:"Segoe UI",sans-serif;
      color:black'> </span><span style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;
      color:#000000;LINE-HEIGHT:0pt'>$strmobile</span></p>

"@
    }
    #set up location to export signaturehtml to.
    $UserDataPath = $Env:appdata
    $FolderLocation = $UserDataPath + '\\Microsoft\\Signatures'
    mkdir $FolderLocation -force
    #create the signature importing the user's details and ensuring the style is set right.
    $signaturehtml = @"
    <head>
    <style>
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    {mso-style-unhide:no;
      mso-style-qformat:yes;
      mso-style-parent:"";
      margin:0cm;
      margin-bottom:.0001pt;
      mso-pagination:widow-orphan;
      font-size:11.0pt;
      font-family:"Calibri",sans-serif;
      mso-ascii-font-family:Calibri;
      mso-ascii-theme-font:minor-latin;
      mso-fareast-font-family:"Times New Roman";
      mso-fareast-theme-font:minor-fareast;
      mso-hansi-font-family:Calibri;
      mso-hansi-theme-font:minor-latin;
      mso-bidi-font-family:"Times New Roman";
      mso-bidi-theme-font:minor-bidi;}
      p {
        margin-top: 0px;
        line-height: 0px;
      }
      span {
        line-height: 0px;
      }
      </style>
      </head>
      <body lang=EN-GB link=#023a98 vlink=#023a98 style='tab-interval:0.0pt'>
      <p class=MsoNormal><b><span
      style='font-size:18.0pt;font-family:"Segoe UI",sans-serif;
      color:#002d6a;LINE-HEIGHT:1pt'>$strname</span></b></p>
      <p class=MsoNormal><b><span
      style='font-size:13.0pt;font-family:"Segoe UI",sans-serif;color:#000000;LINE-HEIGHT:1pt'>
      $strtitle</span></b></p>
      <p class=MsoNormal><b><span
      style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;color:#023a98;LINE-HEIGHT:0pt'>T </span></b><span style='font-size:10.0pt;
      font-family:"Segoe UI",sans-serif;
      color:black'>&nbsp;</span><span style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;
      color:#000000;LINE-HEIGHT:0pt'>$companyphone EXT: $strPhone</span></p>
      $($ddi)
      $($mobile)
      <p class=MsoNormal><b><span
      style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;color:#023a98;LINE-HEIGHT:0pt'>E</span></b> <a href="mailto:$stremail"><span style='color:#000000'>$stremail</span></p>
      <p class=MsoNormal><b><span
      style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;color:#023a98;LINE-HEIGHT:0pt'></span></b><span style='font-size:10.0pt;
      font-family:"Segoe UI",sans-serif;
      color:black'>&nbsp;</span><span style='font-size:10.0pt;font-family:"Segoe UI",sans-serif;
      color:#000000;LINE-HEIGHT:0pt'>$strCompany $StreetAddress $strCity $strPostCode</span></p>
      <br>
      <img border=0 width=$imgheight height=$imgwidth
      src="$banner"</p>

      </body>

      </html>
"@
    #Output the file to outlook signature location.
    $signaturehtml | Out-File "$FolderLocation\\$env:username.htm"
    #Set the signature in outlook as default.
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
}
#Invoke the signature function.
Update-Signature
