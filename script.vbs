Dim WshNetwork, fso

Set WshNetwork = WScript.CreateObject("WScript.Network")
Set fso = WScript.CreateObject("Scripting.FileSystemObject" ) 

' Si un lecteur réseau possédant la lettre S: existe, on le supprime
If fso.DriveExists("S:" ) Then 
    WshNetwork.RemoveNetworkDrive "S:", true, false
End if

' Mapping du lecteur réseau
Dim strRemoteShare , strUser, strPassword

strUser = "demo" ' Utilisateur
strPassword = "demo" ' Mot de passe
strRemoteShare = "\\server\users" ' Chemin réseau (exemple)

WshNetwork.MapNetworkDrive "S:", strRemoteShare, false, strUser, strPassword
' Fin mapping du lecteur réseau

' Message de succès
'WScript.Echo user & ", votre lecteur reseau est mappe."

' Message d'erreur
'WScript.Echo "Echec du mapping du lecteur reseau."

WScript.Quit