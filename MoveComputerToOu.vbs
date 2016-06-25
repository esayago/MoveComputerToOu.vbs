Option Explicit

Dim WshNetwork
Dim ComputerName
Dim objADSysInfo
Dim strComputerName
Dim objComputer
Dim strOUName
Dim StrOU
DIM objNewOU
DIM objMoveComputer
DIM ADS_SECURE_AUTHENTICATION
DIM sADDomain
DIM sADUser
DIM sADPassword
DIM userAuth
DIM objRootDSE

ADS_SECURE_AUTHENTICATION = 1
sADDomain = "domain"
sADUser = "user" 'domain admin account
sADPassword = "password"
Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName

Set objADSysInfo = CreateObject("ADSystemInfo")
strComputerName = objADSysInfo.ComputerName
Set objRootDSE = GetObject("LDAP:")

Set objComputer = GetObject("LDAP://" & strComputerName)
strOUName = objComputer.DistinguishedName
Set objNewOU = objRootDSE.OpenDSObject("LDAP://OU=Laptops,OU=CallCenter,OU=Surgeenergy Sync'd Objects,OU=AMBIT,DC=AMBIT,DC=CORP,DC=LOCAL", sADDomain & "\" & sADUser, sADPassword, ADS_SECURE_AUTHENTICATION)
Set objMoveComputer = objNewOU.MoveHere _
    ("LDAP://" & strComputerName, vbNullString)
