Dim application, connection, session, transactionCode, valeur
Dim WshShell, SAPTitle, hwnd
Dim sapUsername, sapPassword ' Variables globales pour stocker les identifiants SAP

'=== 1. Lecture des arguments passés en ligne de commande
If WScript.Arguments.Count < 1 Then
    MsgBox "Vous devez spécifier au moins un code de transaction (ex: /nCS03)."
    WScript.Quit
Else
    transactionCode = WScript.Arguments(0)
End If

'=== 2. Gestion du second argument (valeur)
If WScript.Arguments.Count >= 2 Then
    valeur = WScript.Arguments(1)
End If

On Error Resume Next  ' Activer la gestion des erreurs
InitializeSAPConnection()
If Err.Number <> 0 Then
    WScript.Sleep 500
    RelancerSAP()
    InitializeSAPConnection()
 End If

'===============================================================================
'        4) SELECTION DE LA TRANSACTION SELON transactionCode
'===============================================================================
Select Case LCase(transactionCode)
    Case "/nzp20"     : ZP20()      ' Besoins par article
    Case "/ncs03"     : CS03()      ' Structure technique (BOM)
    Case "/nmd04"     : MD04()      ' Situation stock/approvisionnement
    Case "/nls24"     : LS24()      ' Stock magasin
    Case "/ncs15"     : CS15()      ' O utilis (where used)
    Case "/nmb51"     : MD51()      ' Liste mouvements stocks
    Case "/ncoois_ref": COOIS_REF() ' Ordres de fabrication par référence
    Case "/ncoois_of" : COOIS_OF()  ' Ordres de fabrication par OF
    Case "/nmcbz"     : MCBZ()      ' Besoins nets
    Case "/nzpan"     : ZPAN()      ' Liste ZPAN
    Case "nom_inter"   : ZP20_Dic()  ' 
    Case Else         : TerminateScript "Code transaction inconnu : " & transactionCode
End Select

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

' ZP20 - Besoins par article
' Paramètre : [valeur] = Code article
Sub ZP20()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzp20"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtP_RACINE").Text = valeur
    Session.FindById("wnd[0]/usr/ctxtP_RACINE").caretPosition = 14
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Session.FindById("wnd[0]/tbar[1]/btn[17]").Press
End Sub

' CS03 - Structure technique (BOM)
' Paramètre : [valeur] = Numéro de nomenclature
Sub CS03()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nCS03"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtRC29N-MATNR").Text = valeur
    Session.FindById("wnd[0]").SendVKey 0
End Sub

' MD04 - Situation stock/approvisionnement
' Paramètre : [valeur] = Code article
Sub MD04()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = valeur
    Session.FindById("wnd[0]").SendVKey 0
End Sub

' LS24 - Stock magasin
' Paramètre : [valeur] = Code article
Sub LS24()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nls24"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtRL01S-LGNUM").Text = "BWM"
    Session.FindById("wnd[0]/usr/ctxtRL01S-MATNR").Text = valeur
    Session.FindById("wnd[0]").SendVKey 0
End Sub

' CS15 - O utilis (where used)
' Paramètre : [valeur] = Code article
Sub CS15()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncs15"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/chkRC29L-UEBKL").Selected = True
    Session.FindById("wnd[0]/usr/ctxtRC29L-MATNR").Text = valeur
    Session.FindById("wnd[0]/tbar[1]/btn[5]").Press
    Session.FindById("wnd[0]/usr/chkRC29L-MEHRS").Selected = True
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End Sub

' MD51 - Liste mouvements stocks
' Paramètre : [valeur] = Code article
Sub MD51()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb51"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = valeur
    Session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "ZBDD"
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End Sub

' COOIS_REF - Ordres de fabrication par référence
' Paramètre : [valeur] = Référence
Sub COOIS_REF()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOA000"
    Session.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-PROFID_HIER").Text = "ZM0000000001"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = valeur
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End Sub

' COOIS_OF - Ordres de fabrication par OF
' Paramètre : [valeur] = Numéro d'OF
Sub COOIS_OF()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").Key = "PPIOA000"
    Session.FindById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-PROFID_HIER").Text = "ZM0000000001"
    Session.FindById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_AUFNR-LOW").Text = valeur
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Session.FindById("wnd[0]/usr/lbl[5,4]").SetFocus
    Session.FindById("wnd[0]").SendVKey 2
    Session.FindById("wnd[0]/usr/lbl[9,8]").SetFocus
    Session.FindById("wnd[0]").SendVKey 2

End Sub

' MCBZ - Besoins nets
' Paramètre : [valeur] = Code article
Sub MCBZ()
    On Error Resume Next
    session.FindById("wnd[0]").Maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMCBZ"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/usr/ctxtSL_WERKS-LOW").Text = "DI61"
    session.FindById("wnd[0]/usr/btn%_SL_MATNR_%_APP_%-VALU_PUSH").Press
    Dim popupWindow : Set popupWindow = WaitForObject("wnd[1]", 2000)
    If Not popupWindow Is Nothing Then
        popupWindow.FindById("tbar[0]/btn[24]").Press
        popupWindow.FindById("tbar[0]/btn[8]").Press
    End If
    session.FindById("wnd[0]/usr/ctxtSL_SPMON-LOW").Text = "01.2012"
    session.FindById("wnd[0]/usr/ctxtSL_SPMON-HIGH").Text = "12.2099"
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End Sub
'-----------------------------------------------------------------------------------

Sub ZPAN()
    Session.FindById("wnd[0]").maximize
    SAPTitle = session.findById("wnd[0]").Text
    PremPlan()
    Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "ZPAN"
    Session.FindById("wnd[0]").SendVKey 0
    Session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Session.FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").Select
    Session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
    Session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    Session.FindById("wnd[1]/tbar[0]/btn[0]").Press
End Sub

'---------------------------------------------------------------------------------
' ZP20_Dictionnaire - Besoins par article
' Paramètre : [valeur] = Code article
Sub ZP20_Dic()
' Création du FileSystemObject (non utilisé ici, mais présent si besoin)
Set fso = CreateObject("Scripting.FileSystemObject")

' Création d'un dictionnaire pour stocker les données extraites
Dim dict
Set dict = CreateObject("Scripting.Dictionary")


' Navigation dans SAP
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzp20"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtP_RACINE").text = valeur
session.findById("wnd[0]/tbar[1]/btn[8]").press



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

' Conversion du dictionnaire en chaîne JSON
jsonString = DictionaryToJson(dict)

' Créer et écrire dans le fichier data_ref.json
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "data_ref.json"), True)
file.WriteLine jsonString
file.Close

Set dict = Nothing

End Sub

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