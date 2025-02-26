' Vérifie si un argument est passé
If WScript.Arguments.Count > 0 Then
    ' Récupère le premier argument
    Dim arg
    arg = WScript.Arguments(0)

    ' Affiche le message avec l'argument
    MsgBox "Message 1: " & arg
Else
    ' Affiche un message d'erreur si aucun argument n'est passé
    MsgBox "Aucun argument fourni."
End If

