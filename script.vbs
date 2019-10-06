Dim WshNetwork, fso

Set WshNetwork = WScript.CreateObject("WScript.Network")
Set fso = WScript.CreateObject("Scripting.FileSystemObject" ) 

' Si un lecteur réseau possédant la lettre S: existe, on le supprime
If fso.DriveExists("S:" ) Then 
    WshNetwork.RemoveNetworkDrive "S:", true, false
End if

' Mapping du lecteur réseau
Dim strRemoteShare, strUser, strPassword

strUser = InputBox("Nom d'utilisateur") ' Utilisateur
strPassword = InputBox("Mot de passe") ' Mot de passe
strRemoteShare = "\\server" ' Chemin réseau (exemple)

WshNetwork.MapNetworkDrive "S:", strRemoteShare, false, strUser, strPassword
' Fin mapping du lecteur réseau