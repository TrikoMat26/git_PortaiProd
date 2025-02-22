'-------------------------------
' Script VBScript pour SAP / JSON
'-------------------------------

Dim application, connection, session, valeur
Dim WshShell, SAPTitle, hwnd
Dim sapUsername, sapPassword ' Variables globales pour stocker les identifiants SAP

'=== 1. Lecture des arguments passés en ligne de commande
If WScript.Arguments.Count < 1 Then
    MsgBox "Vous devez spécifier au moins une référence."
    WScript.Quit
Else
    valeur = WScript.Arguments(0)
End If

On Error Resume Next  ' Activer la gestion des erreurs
InitializeSAPConnection()
If Err.Number <> 0 Then
    WScript.Sleep 500
    RelancerSAP()
    InitializeSAPConnection()
End If

' Utilisation de l'API Windows pour forcer la mise au premier plan
Sub PremPlan()
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.AppActivate SAPTitle
    WScript.Sleep 100
End Sub

Sub RelancerSAP()
    On Error Resume Next  ' Activer la gestion des erreurs pour éviter tout crash ici
    Set WshShell = WScript.CreateObject("WScript.Shell")

    ' Vérifier si les identifiants sont déjà définis
    If sapUsername = "" Or sapPassword = "" Then
        ' Demander le nom d'utilisateur
        sapUsername = InputBox("Entrez votre nom d'utilisateur SAP :", "Identifiants SAP")
        ' Demander le mot de passe
        sapPassword = InputBox("Entrez votre mot de passe SAP :", "Identifiants SAP")
    End If

    ' Lancer SAP avec les identifiants
    If sapUsername <> "" And sapPassword <> "" Then
        Dim sapShortcutPath
        sapShortcutPath = """C:\Program Files (x86)\SAP\FrontEnd\SapGui\sapshcut.exe"""
        Dim sapParams
        sapParams = " -system=PEO -client=600 -user=" & sapUsername & " -pw=" & sapPassword & " -language=FR"
        WshShell.Run sapShortcutPath & sapParams
        WScript.Sleep 5000  ' Pause pour laisser SAP se lancer
    Else
        MsgBox "Nom d'utilisateur ou mot de passe non fourni. Impossible de relancer SAP."
    End If
End Sub

Sub InitializeSAPConnection()
    If Not IsObject(application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set application = SapGuiAuto.GetScriptingEngine
    End If

    If Not IsObject(connection) Then
        Set connection = application.Children(0)
    End If

    If Not IsObject(session) Then
        Set session = connection.Children(0)
    End If

    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If

End Sub

' Navigation dans SAP
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzp20"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtP_RACINE").text = valeur
session.findById("wnd[0]/tbar[1]/btn[8]").press

' Création d'un dictionnaire pour stocker les données extraites
Dim dict
Set dict = CreateObject("Scripting.Dictionary")

' Récupérer les clés des noeuds de l'arborescence SAP
Set ZP20 = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").GetAllNodeKeys()

For i = 0 To ZP20.length - 1
    nodekey = ZP20.ElementAt(i)
    Key = String(11 - Len(CStr(nodekey)), " ") & CStr(nodekey)

    ' Récupérer les valeurs associées pour chaque noeud
    valZTOPO = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").getItemText(Key, "ZTOPO")
    valZCOMP = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").getItemText(Key, "ZCOMP")
    valZDES  = session.findById("wnd[0]/usr/cntlTREE_CONTAINER/shellcont/shell").getItemText(Key, "ZDES")

    ' Utiliser la valeur ZTOPO comme clé et stocker un tableau contenant ZCOMP et ZDES
    If Not dict.Exists(valZTOPO) Then
        dict.Add valZTOPO, Array(valZCOMP, valZDES)
    End If
Next

' Quitter la transaction SAP
session.findById("wnd[0]/tbar[0]/btn[15]").press

' Fonction de conversion du dictionnaire en JSON en utilisant Chr(34) pour les guillemets
Function DictionaryToJson(d)
    Dim key, json, arr, i, j, value
    json = "["
    i = 0
    For Each key In d.Keys
        If i > 0 Then json = json & ","
        json = json & "["
        json = json & Chr(34) & key & Chr(34) & ","
        arr = d.Item(key)
        For j = 0 To UBound(arr)
            If j > 0 Then json = json & ","
            value = arr(j)
            json = json & Chr(34) & value & Chr(34)
        Next
        json = json & "]"
        i = i + 1
    Next
    json = json & "]"
    DictionaryToJson = json
End Function

' Conversion du dictionnaire en chaîne JSON
jsonString = DictionaryToJson(dict)

' Créer et écrire dans le fichier data_ref.json
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "data_ref.json"), True)
file.WriteLine jsonString
file.Close

Set dict = Nothing
