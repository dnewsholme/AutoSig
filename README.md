AutoSig
=======

#Automatic Signature Creation For Outlook

This Script will create a signature based on AD attributes. Using Ldap lookup via .Net methods

The script should be a logon script or event driven when outlook is opened. Each time the script runs the signature will be recreated.
Direct Dial and mobile will only populate if the fields have entries in ActiveDirectory


#Customizing the script.
The HTML code is in the script. Colours can be changed by updating the colour codes `#000000` ect.
The `$banner` variable should be updated to your own image path on a publicly accessible url.
`$imgheight` and `imgwidth` should be set to your custom image dimensions.
