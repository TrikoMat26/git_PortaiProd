'-------------------------------
' Script VBScript pour SAP / JSON
'-------------------------------

Dim application, connection, session, valeur
Dim WshShell, SAPTitle, hwnd
Dim sapUsername, sapPassword ' Variables globales pour stocker les identifiants SAP

' Constantes pour le registre
Const HKEY_CURRENT_USER = &H80000001
Const REG_SZ = 1
Const REG_PATH = "Software\PortailProd\SAPCredentials"

'=== 1. Lecture des arguments passés en ligne de commande
If WScript.Arguments.Count < 1 Then
    MsgBox "Vous devez spécifier au moins une référence."
    WScript.Quit
Else
    valeur = WScript.Arguments(0)
End If

' Charger les identifiants enregistrés
GetSavedCredentials()

On Error Resume Next  ' Activer la gestion des erreurs
InitializeSAPConnection()
If Err.Number <> 0 Then
    WScript.Sleep 500
    RelancerSAP()
    InitializeSAPConnection()
End If

' Fonction pour récupérer les identifiants enregistrés
Sub GetSavedCredentials()
    On Error Resume Next
    
    Set WshShell = CreateObject("WScript.Shell")
    sapUsername = WshShell.RegRead("HKCU\" & REG_PATH & "\Username")
    
    ' Pour le mot de passe, on utilise un décodage simple
    Dim encodedPassword
    encodedPassword = WshShell.RegRead("HKCU\" & REG_PATH & "\Password")
    If Err.Number = 0 And encodedPassword <> "" Then
        sapPassword = DecodePassword(encodedPassword)
    End If
    
    On Error Goto 0
End Sub

' Fonction pour enregistrer les identifiants
Sub SaveCredentials(username, password)
    On Error Resume Next
    
    Set WshShell = CreateObject("WScript.Shell")
    
    ' Créer la clé si elle n'existe pas
    WshShell.RegWrite "HKCU\" & REG_PATH & "\", "", "REG_SZ"
    
    ' Enregistrer le nom d'utilisateur
    WshShell.RegWrite "HKCU\" & REG_PATH & "\Username", username, "REG_SZ"
    
    ' Enregistrer le mot de passe avec un encodage simple
    WshShell.RegWrite "HKCU\" & REG_PATH & "\Password", EncodePassword(password), "REG_SZ"
    
    On Error Goto 0
End Sub

' Encodage simple du mot de passe (ce n'est pas un cryptage fort)
Function EncodePassword(password)
    Dim i, encoded
    encoded = ""
    For i = 1 To Len(password)
        encoded = encoded & Asc(Mid(password, i, 1)) & "."
    Next
    EncodePassword = encoded
End Function

' Décodage du mot de passe
Function DecodePassword(encoded)
    Dim arr, i, decoded
    arr = Split(encoded, ".")
    decoded = ""
    For i = 0 To UBound(arr) - 1
        If arr(i) <> "" Then
            decoded = decoded & Chr(CInt(arr(i)))
        End If
    Next
    DecodePassword = decoded
End Function

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
        
        ' Demander si l'utilisateur veut enregistrer les identifiants
        If sapUsername <> "" And sapPassword <> "" Then
            Dim saveCredentialsResponse
            saveCredentialsResponse = MsgBox("Souhaitez-vous enregistrer vos identifiants pour les prochaines connexions?", vbYesNo + vbQuestion, "Enregistrer identifiants")
            
            If saveCredentialsResponse = vbYes Then
                SaveCredentials sapUsername, sapPassword
            End If
        End If
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
