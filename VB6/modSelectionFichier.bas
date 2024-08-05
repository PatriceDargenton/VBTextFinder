Attribute VB_Name = "modSelectionFichier"
Option Explicit

' Module pour choisir un fichier dans une bo�te de dialogue, m�thode bas�e sur les API
'  (le contr�le MSComDlg.CommonDialog �tant limit� � l'environnement VB6)

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function bChoisirUnFichierAPI(ByRef sFichier$, sFiltre$, sTitre$, _
    sInitDir$, lHandelWnd&) As Boolean
    
    Dim OpenFile As OPENFILENAME
    Dim lRet&
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = lHandelWnd
    OpenFile.lpstrFilter = sFiltre
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    ' Ne pas r�initialiser le r�pertoire par d�faut si on ne le demande pas
    If sInitDir <> "" Then OpenFile.lpstrInitialDir = sInitDir
    OpenFile.lpstrTitle = sTitre
    OpenFile.flags = &H1000 ' FileMustExist (OFN_FILEMUSTEXIST)
    lRet = GetOpenFileName(OpenFile)
    If lRet = 0 Then
        sFichier = ""
    Else
        sFichier = Trim$(OpenFile.lpstrFile)
        ' Enlever les caract�res null � la fin
        Dim iPos%
        iPos = InStr(sFichier, vbNullChar)
        If iPos Then sFichier = Left$(sFichier, iPos - 1)
        bChoisirUnFichierAPI = True
    End If
    
End Function
