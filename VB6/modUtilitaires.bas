Attribute VB_Name = "modUtilitaires"
Option Explicit

' Module de fonctions utilitaires

Public Function bFichierExiste(sCheminFichier$) As Boolean
    
    ' Retourner l'existence ou non d'un fichier avec un chemin complet
    
    If sCheminFichier = "" Then Exit Function
    On Error Resume Next
    bFichierExiste = (Len(Dir$(sCheminFichier)) > 0)
    If Err <> 0 Then bFichierExiste = False
    
End Function

Public Function sExtraireChemin$(ByVal sFichier$, _
    Optional sNomFichier$ = "", Optional sExtension$ = "", _
    Optional sNomFichierSansExt$ = "")
    
    ' Retourner le chemin du fichier passé en argument
    ' Non compris le caractère \
    ' Retourner aussi le nom du fichier sans le chemin ainsi que son extension
    
    Dim sChemin$, iTaille%, i%, sCar$
    
    On Error GoTo Err_ExtraireChemin
    
    iTaille = Len(sFichier)
    For i = iTaille To 1 Step -1
        sCar = Mid$(sFichier, i, 1)
        If sCar = "\" Or sCar = ":" Then
            sChemin = Left$(sFichier, i - 1)
            sNomFichier = Mid$(sFichier, i + 1)
            Exit For
        End If
    Next i
    
    ' Rechercher l'extension
    iTaille = Len(sNomFichier)
    For i = iTaille To 1 Step -1
        sCar = Mid$(sNomFichier, i, 1)
        If sCar = "." Then
            sExtension = Mid$(sNomFichier, i + 1)
            Exit For
        End If
    Next i
    
    If sExtension <> "" Then _
        sNomFichierSansExt = Mid$(sNomFichier, 1, _
            Len(sNomFichier) - Len(sExtension) - 1)
    
    sExtraireChemin = sChemin
    Exit Function

Err_ExtraireChemin:
    AfficherMsgErreur Err, "sExtraireChemin"
    Exit Function
      
End Function

Public Function bCreerObjet(ByRef oObjetQcq As Object, ByVal sClasse$, _
    Optional ByRef bObjetDejaOuvert As Boolean) As Boolean

    ' Instancier un contrôle ActiveX en liaison tardive (à l'exécution)

    On Error Resume Next
    ' Vérifier si le serveur automation est déjà activé
    Set oObjetQcq = GetObject(, sClasse)
    If Err <> 0 Then
        bObjetDejaOuvert = False
        Err.Clear
        Set oObjetQcq = CreateObject(sClasse)
    Else
        bObjetDejaOuvert = True
    End If
    If Err <> 0 Then
        AfficherMsgErreur Err, "bCreerObjet", _
            "L'objet de classe [" & sClasse & "] ne peut pas être créé", vbCritical
        Err.Clear: Set oObjetQcq = Nothing: GoTo Fin
    End If
    bCreerObjet = True
    
Fin:
    On Error GoTo 0
    
End Function

Public Sub AfficherMsgErreur(Erreur As Object, Optional sTitreFct$ = "", _
    Optional sInfo$ = "", Optional sDetailMsgErr$ = "")
    
    Const vbDefault% = 0
    'Sablier bDesactiver:=True ' Dépend du contexte applicatif VBA ou VB6
    Dim sMsg$
    If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
    If sInfo <> "" Then sMsg = sMsg & vbCrLf & sInfo
    If Erreur.Number Then
        sMsg = sMsg & vbCrLf & "Err n°" & Str$(Erreur.Number) & " :"
        sMsg = sMsg & vbCrLf & Erreur.Description
    End If
    If sDetailMsgErr <> "" Then sMsg = sMsg & vbCrLf & sDetailMsgErr
    MsgBox sMsg, vbCritical, sTitreMsg
    
End Sub

Public Function sLireCheminApplication$()
    
    ' Lire le chemin de l'application VBTextFinder
    
    #If bVersionWord Then    ' Word
        sLireCheminApplication = Application.ActiveDocument.Path
    #ElseIf bVersionVBA Then ' Sinon Excel
        sLireCheminApplication = Application.ActiveWorkbook.Path
    #Else                    ' Sinon VB6
        sLireCheminApplication = App.Path
    #End If
    
End Function

Public Function sLireCheminFichierApplication$()
    
    ' Lire le chemin avec le fichier de l'application VBTextFinder
    
    #If bVersionWord Then    ' Word
        sLireCheminFichierApplication = Application.ActiveDocument.FullName
    #ElseIf bVersionVBA Then ' Sinon Excel
        sLireCheminFichierApplication = Application.ActiveWorkbook.FullName
    #Else                    ' Sinon VB6
        sLireCheminFichierApplication = App.Path & "\VBTextFinder.exe"
    #End If
    
End Function

Public Function asArgLigneCmd(sFichiers$) As String()

    ' Retourner les arguments de la ligne de commande
    ' Voir aussi : Commande du menu contextuel pour récupérer
    '  les chemins d'une sélection de fichiers dans l'explorateur
    ' www.vbfrance.com/code.aspx?ID=36426

    Dim iNbArg%, asArgs$()
    Dim sGm$, sFichier$, sSepar$, bNomLong As Boolean
    sGm = Chr$(34) ' Guillemets
    
    'MsgBox "Arguments : " & Command, vbInformation, sTitreMsg
    
    ' Parser les noms cours : facile
    'asArgs = Split(Command, " ")
    
    ' Parser les noms longs (fonctionne quelque soit le nombre de fichiers)
    ' Chaque nom long de fichier est entre guillemets : "
    '  une fois le nom traité, les guillemets sont enlevé
    ' S'il y a un non court parmi eu, il n'est pas entre guillemets
    
    Dim sCmd$, iLen%, iFin%, iDeb%, iDeb2%, bFin As Boolean
    sCmd = sFichiers 'Command$
    iLen = Len(sCmd)
    iDeb = 1
    Do
        
        bNomLong = False: sSepar = " "
        ' Si le premier caractère est un guillement, c'est un nom long
        If Mid$(sCmd, iDeb, 1) = sGm Then bNomLong = True: sSepar = sGm
        
        iDeb2 = iDeb
        ' Supprimer les guillemets dans le tableau de fichiers
        If bNomLong Then iDeb2 = iDeb2 + 1
        
        ' 10/9/2005 : iDeb2+1 au lieu de +2 (cf. AccessBackup)
        iFin = InStr(iDeb2 + 1, sCmd, sSepar)
        
        ' Si le séparateur n'est pas trouvé, c'est la fin de la ligne de commande
        If iFin = 0 Then bFin = True: iFin = iLen + 1
        
        sFichier = Trim$(Mid$(sCmd, iDeb2, iFin - iDeb2))
        If sFichier <> "" Then
            ReDim Preserve asArgs(iNbArg)
            asArgs(iNbArg) = sFichier
            iNbArg = iNbArg + 1
        End If
        
        If bFin Or iFin = iLen Then Exit Do
        
        iDeb = iFin + 1
        If bNomLong Then iDeb = iFin + 2
        
    Loop
    
    asArgLigneCmd = asArgs

End Function

Public Function bTableauVide(aString$()) As Boolean

    ' Renvoyer True si le tableau est vide

    On Error Resume Next
    Dim iArgMin%
    iArgMin = LBound(aString())
    If Err > 0 Then bTableauVide = True
    On Error GoTo 0
    Err.Clear

End Function

Public Function bSupprimerFichier(ByVal sCheminFichier$) As Boolean
    
    ' Vérifier si le fichier existe
    If False = bFichierExiste(sCheminFichier) Then _
        bSupprimerFichier = True: Exit Function
    
    On Error Resume Next ' Impossible de supprimer à cause d'un verrou
    Kill sCheminFichier
    Err.Clear: On Error GoTo 0
    If bFichierExiste(sCheminFichier) Then
        MsgBox "Impossible de supprimer le fichier :" & vbLf & _
            sCheminFichier & vbLf & _
            "Cause possible : le fichier est ouvert avec un logiciel", _
            vbCritical, sTitreMsg
        Exit Function
    End If
    
    bSupprimerFichier = True
    
End Function

