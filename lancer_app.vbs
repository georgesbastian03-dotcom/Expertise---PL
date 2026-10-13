' Lance l'application au quotidien, sans afficher de fenêtre noire à l'écran
' (contrairement à un double-clic direct sur lancer.bat).
'
' Avant la toute première utilisation, lancez d'abord "lancer.bat" lui-même
' au moins une fois (pour installer les dépendances et voir les éventuels
' messages d'erreur) — ou utilisez simplement le raccourci créé par
' "Creer_raccourci_bureau.vbs", qui gère ce cas automatiquement.
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
dossier = fso.GetParentFolderName(WScript.ScriptFullName)

If fso.FileExists(dossier & "\.deps_ok") Then
    ' Dépendances déjà installées : lancement entièrement invisible.
    shell.Run """" & dossier & "\lancer.bat""", 0, False
Else
    ' Première installation : on laisse la fenêtre visible, pour voir les
    ' éventuels messages d'erreur (Python manquant, pas d'internet...).
    shell.Run """" & dossier & "\lancer.bat""", 1, True
End If
