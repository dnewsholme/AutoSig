AutoSig
=======

Automatic Signature Creation For Outlook

This Script will create a signature based on AD attributes.

Updater should be a logon script or event driven when outlook is opened. Autosig.ps1 should reside on a unc path. Each time Updater is run it will check if a new version of the signature is available by checking the file hash of autosig.ps1
If autosig.ps1 has changed then the signature will update.

Alternatively you can run Autosig as a logon script and the signature will update every time the user logs on. However this will slow down logon time a bit.
