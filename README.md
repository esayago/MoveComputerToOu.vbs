# MoveComputerToOu.vbs
VBScript to move a domain computer to target OU
- Run this script outside of a DC by using a domain admin account
- Edit the script by adding your domain, username and password
    sADDomain = "domain"
    sADUser = "user" 'domain admin account
    sADPassword = "password"
- On line 32 enter the OU you want the object to be moved to as the first parameter:
- The second parameter uses the varialbes holding your credentials for authentication
    "LDAP://targetOU", , sADDomain & "\" & sADUser, sADPassword, ADS_SECURE_AUTHENTICATION

 --- Avoid having a password on a text file by converting it to vbe
Drag and drop the file on top of the Encoded.vbs and it will create a vbe file on the same folder.
