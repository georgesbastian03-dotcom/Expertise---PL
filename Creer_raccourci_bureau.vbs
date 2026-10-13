' À double-cliquer UNE SEULE FOIS : crée une icône sur le Bureau pour lancer
' le programme directement, comme une application normale (plus besoin
' d'ouvrir ce dossier ni de passer par lancer.bat au quotidien). Vous pourrez
' aussi épingler ce raccourci à la barre des tâches Windows une fois créé.
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
dossier = fso.GetParentFolderName(WScript.ScriptFullName)
bureau = shell.SpecialFolders("Desktop")

Set raccourci = shell.CreateShortcut(bureau & "\Rapports de Visite - Experts Nicolai.lnk")
raccourci.TargetPath = dossier & "\lancer_app.vbs"
raccourci.WorkingDirectory = dossier
raccourci.IconLocation = dossier & "\icone.ico"
raccourci.Description = "Rapports de Visite - Experts Nicolai"
raccourci.Save

MsgBox "Raccourci cree sur le Bureau !" & Chr(10) & Chr(10) & _
    "Vous pouvez desormais lancer le programme directement depuis cette icone " & _
    "(vous pouvez aussi l'epingler a la barre des taches en faisant un clic droit dessus).", _
    64, "Rapports de Visite - Experts Nicolai"
