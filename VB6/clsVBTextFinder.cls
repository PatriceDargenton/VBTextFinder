VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBTextFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsVBTextFinder : Classe principale du moteur de recherche VBTextFinder

' Pour pouvoir localiser la ligne ayant provoqué une erreur, mettre bTrapErr = False
Private Const bTrapErr As Boolean = True
'Private Const bTrapErr As Boolean = False

Private Const bDebug As Boolean = False

Private Const sMsgOperationTerminee$ = "Opération terminée."
Private Const iMaxLongChaine% = 32767 ' Longueur max. d'une chaîne dans un contrôle

Private Const rVersionFichierVBTxtFndDat! = 1!
' Fichiers de sauvegarde de l'index du glossaire multidocument
Private Const sFichierVBTxtFndIdxDef$ = "VBTextFinder.idx" ' Sauvegarde en cours
'Private Const sFichierVBTxtFndBak$ = "VBTextFinder.bak" ' Sauvegarde précédente
'Private Const sFichierVBTxtFndTmp$ = "VBTextFinder.tmp" ' Sauvegarde de sécurité
Private Const sMsgGestionIndexDat$ = _
    "Gestion du fichier d'index VBTextFinder.idx"
Private Const sExtVBTF$ = ".idx" '".dat"

Private m_sCheminVBTxtFndTmp$, m_sCheminVBTxtFndBak$, m_sCheminVBTxtFndIdx$
Public m_sCheminDossierCourant$

' Fabrication du document index
Private Const sFichierVBTxtFndTxtAlphab$ = "VBTextFinderAlphab.txt"
Private Const sFichierVBTxtFndDocAlphab$ = "VBTextFinderAlphab.doc"
Private Const sFichierVBTxtFndTxtFreq$ = "VBTextFinderFreq.txt"
Private Const sFichierVBTxtFndDocFreq$ = "VBTextFinderFreq.doc"

' Types de tri du document index
Private Const sTriAlpha$ = "Alphabétique"
Private Const sTriFreq$ = "Fréquentiel"
Private Const sTriDef$ = sTriAlpha

' Types d'affichage des résultats de recherche
Private Const sAfficherPhrase$ = "Phrase"
Private Const sAfficherParag$ = "Paragraphe"
Private Const sAfficherParagPM1$ = "Paragraphe +-1"
Private Const sAfficherParagPM2$ = "Paragraphe +-2"
Private Const sAfficherParagPM3$ = "Paragraphe +-3"
Private Const sAfficherDef$ = sAfficherPhrase

Private Const sCodeDocDef$ = "Doc n°"

Private Const sClasseObjetWord$ = "Word.Application"

Private Const iErrCollectionCleAbsente% = 5

' Indexation directe d'un fichier texte passé
'  en argument de la ligne de commande
Public m_bModeDirect As Boolean
Public m_sCheminFichierTxtDirect$

' Booléen pour pouvoir interrompre une longue opération
Private m_bInterrompre As Boolean

' Référence sur le contrôle d'affichage de l'avancement des opérations
Private m_ctrlLblAvancement As Control
' Référence sur l'instance du formulaire pour pouvoir
'  changer le curseur associé en sablier
Private m_oFrmMousePointer As Object

Private m_sListeSeparateursMot$
'Private m_sCheminApp$ ' Chemin de l'application
Private m_sCheminFichierIndex$ ' Chemin complet avec extension du fichier index
Private m_bFichierIndexDef As Boolean

' Booléen pour savoir si l'index est modifié,
'  auquel cas il doit être sauvé lors de la fermeture de l'application
Private m_bIndexModifie As Boolean

' Collection de mots indexés par eux-même : sMot
Private m_colMots As New Collection
Private m_lNbMotsG& ' Nombre de mots indexés en tout

' Collection de phrases indexées par leur numéro (un tableau suffirait dans ce cas)
Private m_colPhrases As New Collection
Private m_lNbPhrases& ' (= m_colPhrases.Count)
Private m_lNbParagG&  ' Nombre de paragraphes indexés en tout (sans les lignes vides)

' Collection des documents indexés
Private m_colDocs As New Collection

Public Sub Initialiser(ctrlLblAvancement As Control, _
    ctrlLstAfficher As Control, ctrlLstTri As Control, _
    oFrmMousePointer As Object)

    ' Initialisation des contrôles de l'interface
    ctrlLstAfficher.AddItem sAfficherPhrase
    ctrlLstAfficher.AddItem sAfficherParag
    ctrlLstAfficher.AddItem sAfficherParagPM1
    ctrlLstAfficher.AddItem sAfficherParagPM2
    ctrlLstAfficher.AddItem sAfficherParagPM3
    ctrlLstAfficher = sAfficherDef

    ctrlLstTri.AddItem sTriAlpha
    ctrlLstTri.AddItem sTriFreq
    ctrlLstTri = sTriDef

    Set m_ctrlLblAvancement = ctrlLblAvancement
    Set m_oFrmMousePointer = oFrmMousePointer

    m_ctrlLblAvancement.Caption = ""

    m_sCheminFichierIndex = sLireCheminApplication() & "\" & sFichierVBTxtFndIdxDef
    m_bFichierIndexDef = True
    
    Dim sArguments$, sArgument$
    ' Extraire l'option passée en argument de la ligne de commande
    sArguments = Command$
    If sArguments <> "" Then
        Dim asArgs$()
        asArgs = asArgLigneCmd(sArguments)
        If False = bTableauVide(asArgs) Then
            sArgument = asArgs(0)
            If bFichierExiste(sArgument) Then
                Dim sExt$
                sExtraireChemin sArgument, , sExt
                If sExt = sExtVBTF Then
                    m_sCheminFichierIndex = sArgument: m_bFichierIndexDef = False
                ElseIf sExt = "txt" Or sExt = "doc" Or _
                    sExt = "html" Or sExt = "htm" Then
                    m_bModeDirect = True
                    m_sCheminFichierTxtDirect = sArgument
                    m_sCheminFichierIndex = _
                        sExtraireChemin(m_sCheminFichierTxtDirect) & "\" & _
                        sFichierVBTxtFndIdxDef
                End If
            Else
                MsgBox "Impossible de trouver le fichier :" & vbLf & _
                    sArgument, vbCritical, _
                    "Passage d'un fichier au démarrage de VBTextFinder"
            End If
        End If
    End If
    
    Dim sNomFichierSansExt$, sNomFichier$, sExtension$
    m_sCheminDossierCourant = sExtraireChemin(m_sCheminFichierIndex, sNomFichier, _
        sExtension, sNomFichierSansExt)
    m_sCheminVBTxtFndTmp = m_sCheminDossierCourant & "\" & sNomFichierSansExt & ".tmp"
    m_sCheminVBTxtFndBak = m_sCheminDossierCourant & "\" & sNomFichierSansExt & ".bak"
    m_sCheminVBTxtFndIdx = m_sCheminDossierCourant & "\" & sNomFichierSansExt & sExtVBTF
    'm_sCheminDossierFichierIndex = m_sCheminDossierCourant

    m_lNbPhrases = 0
    m_lNbParagG = 0
    m_lNbMotsG = 0
    'Set m_colMots = New Collection ' Recommencer une collection de mots à indexer
    
    m_sListeSeparateursMot = sListeSeparateurMot & Chr$(iCodeASCIITabulation) & _
        Chr$(iCodeASCIIEspaceInsecable) & Chr$(iCodeASCIIGuillement) & _
        Chr$(iCodeASCIIGuillementOuvrant) & Chr$(iCodeASCIIGuillementFermant)

End Sub

Public Function bQuitter() As Boolean
    
    Dim lNbDocs&
    lNbDocs = m_colDocs.Count
    If lNbDocs > 0 And m_bIndexModifie And False = m_bModeDirect Then
        Dim iReponse%
        iReponse = MsgBox("Voulez-vous sauvegarder l'index de VBTextFinder :" & vbLf & _
            m_sCheminVBTxtFndIdx & " ?" & vbLf & _
            "(nombre de documents : " & lNbDocs & ")", _
            vbYesNoCancel Or vbQuestion, sMsgGestionIndexDat)
        If iReponse = vbCancel Then Exit Function
        If iReponse = vbNo Then GoTo Fin
        ' Valider la copie temporaire en copie de sauvegarde pour de bon
        '  (ou sinon sauver simplement l'index si l'option bSauvegardeSecurite = False)
        If Not bValiderSauvegardeTmp() Then Exit Function
    End If
    
Fin:
    ' Je me suis toujours posé la question si c'était nécessaire de faire
    '  Set objet = Nothing et la réponse est oui ! sinon VB6 ne parvient pas toujours
    '  à désallouer la RAM (cela arrive même systématiquement pour des gros documents)
    Sablier
    AfficherMessage "Désallocation des mots en mémoire vive..."
    Set m_colMots = Nothing
    AfficherMessage "Désallocation des phrases en mémoire vive..."
    Set m_colPhrases = Nothing
    Set m_colDocs = Nothing
    Sablier bDesactiver:=True
    bQuitter = True

End Function

Public Sub Interrompre()

    ' VB est un langage événementiel multi-thread : deux fonctions peuvent très bien
    '  fonctionner simultanément, on se sert de cela pour pouvoir interrompre une
    '  opération en cours assez longue (il peut même arriver qu'une même fonction
    '  en cours soit ré-appelée : ré-entrance)
    m_bInterrompre = True

End Sub

Public Function iNbDocumentsIndexes%()

    iNbDocumentsIndexes = m_colDocs.Count

End Function

Public Sub ChercherOccurrencesMot(CtrlMot As Control, CtrlResultat As Control, _
    sTypeAffichage$)

    ' Fonction principale du moteur de recherche

    ' Lors de l'initialisation du logiciel, la zone est vide
    Dim sMot$
    sMot = CtrlMot
    If sMot = "" Then Exit Sub
    
    m_bInterrompre = False
    Sablier
    
    Dim oMot As clsMot
    If Not bMotExiste(sMot, oMot) Then
        AfficherMessage "Mot non trouvé : " & sMot
        GoTo Fin
    End If
    
    Dim iTailleLimite%
    iTailleLimite = iMaxLongChaine - 4 ' Laisser de la place pour afficher "..."
    Dim bTailleLimite As Boolean
    
    Dim i&, lNbPhrases&, lNumPhrase&, sResultat$, lMemNumPhraseMin&, lMemNumParagMax&
    Dim oPhrase As clsPhrase, lMemNumPhrase&, lMemNumParagraphe&
    Dim bUnSeulDocument As Boolean
    lNbPhrases = oMot.lNbPhrases
    If m_colDocs.Count = 1 Then bUnSeulDocument = True
    Dim sCodeDoc$, sMemCodeDoc$, sAffCodeDoc$, sMemAffCodeDoc$, lNumParagraphe&
    
    lMemNumPhraseMin = 1
    lMemNumParagMax = 0
    sMemCodeDoc = ""
    
    For i = 1 To lNbPhrases ' Parcours de toutes les phrases contenant ce mot
        
        AfficherMessage "Recherche en cours... " & i & " / " & lNbPhrases
        If m_bInterrompre Then GoTo FinRecherche
        
        lNumPhrase = oMot.lLireNumPhrase(i) ' Numéro de phrase global aux documents
        ' Un mot est présent plusieurs fois dans la même phrase
        If lNumPhrase = lMemNumPhrase Then GoTo PhraseSuivante
        
        Set oPhrase = m_colPhrases(lNumPhrase)
        
        lNumParagraphe = oPhrase.lNumParagraphe ' Numéro de parag. global aux documents
        
        If bUnSeulDocument Then
            ' S'il n'y a qu'un seul document, inutile de le rappeler
            sAffCodeDoc = ""
        Else
            Dim oDoc As clsDoc
            Set oDoc = m_colDocs.Item(oPhrase.sCodeDoc)
            sCodeDoc = oDoc.sCodeDoc
            If oDoc.sCodeDoc <> oDoc.sCle Then
                ' Si le code du document a été édité, on l'affiche
                sAffCodeDoc = oDoc.sCodeDoc & " "
            Else
                ' Sinon on affiche le chemin
                sAffCodeDoc = vbCrLf & oDoc.sChemin & " : " & vbCrLf
                ' Si on indexe plusieurs fois le même document
                If sCodeDoc <> sMemCodeDoc And sMemAffCodeDoc = sAffCodeDoc Then
                    sAffCodeDoc = vbCrLf & oDoc.sCodeDoc & " " & oDoc.sChemin & " : " & vbCrLf
                End If
                ' S'il n'a pas changé depuis le précédent, ne plus l'afficher
                If sCodeDoc = sMemCodeDoc Then sAffCodeDoc = ""
                sMemCodeDoc = sCodeDoc
                sMemAffCodeDoc = sAffCodeDoc
            End If
        End If
        
        If sTypeAffichage = sAfficherPhrase Then
            ' § = symbole pour paragraphe
            sResultat = sResultat & sAffCodeDoc & "§ n°" & lNumParagraphe & " Ph. n°" & _
                oPhrase.lNumPhraseL & " : " & oPhrase.sPhrase & vbCrLf
            lMemNumPhrase = lNumPhrase
        Else ' Parag
            Dim j&, lNumPhrase1ParagMotTrouve&
            If lNumParagraphe = lMemNumParagraphe Then GoTo PhraseSuivante
            If lNumParagraphe = lMemNumParagMax Then GoTo PhraseSuivante
            ' Trouver toutes les phrases du paragraphe
            lNumPhrase1ParagMotTrouve = lNumPhrase ' Initialisation par défaut
            ' Commencer par rechercher les n° de paragraphes des phrases précédentes
            For j = lNumPhrase - 1 To 1 Step -1
                Set oPhrase = m_colPhrases(j)
                If oPhrase.lNumParagraphe = lNumParagraphe Then
                    ' La phrase précédente appartient au même paragraphe
                    '  on commence donc plus tôt l'affichage du paragraphe
                    lNumPhrase1ParagMotTrouve = j
                Else
                    ' La phrase précédente appartient au paragraphe précédent
                    '  l'affichage du paragraphe commence donc à la phrase suivante
                    lNumPhrase1ParagMotTrouve = j + 1: Exit For
                End If
            Next j
            
            ' Puis noter toutes les phrases du paragraphe +- l'écart demandé
            Dim iNbParagAv%, iNbParagAp%
            If sTypeAffichage = sAfficherParagPM3 Then     ' +-3 §
                iNbParagAv = 3: iNbParagAp = 3
            ElseIf sTypeAffichage = sAfficherParagPM2 Then ' +-2 §
                iNbParagAv = 2: iNbParagAp = 2
            ElseIf sTypeAffichage = sAfficherParagPM1 Then ' +-1 §
                iNbParagAv = 1: iNbParagAp = 1
            Else
                iNbParagAv = 0: iNbParagAp = 0 ' Afficher seulement le § en cours
            End If
            
            Dim lNumPhraseDebRech&, lNumPhraseFinRech&, lMemNumPhrasePreced&, lMin&
            
            ' Rechercher les n° de § précédants
            lMin = 1
            lMemNumPhrasePreced = lNumPhrase1ParagMotTrouve
            ' Cas où plusieurs § successifs contiennent le mot : un seul affichage
            If lMemNumPhrasePreced < lMemNumPhraseMin Then
                lMemNumPhrasePreced = lMemNumPhraseMin
                lMin = lMemNumPhraseMin
            End If
            lNumPhraseDebRech = lMemNumPhrasePreced
            For j = lNumPhrase1ParagMotTrouve To lMin Step -1
                Set oPhrase = m_colPhrases(j)
                ' Noter le n° global de la phrase en cours
                If j = lNumPhrase1ParagMotTrouve Then lMemNumPhrase = oPhrase.lNumPhraseG
                ' Ne pas afficher plusieurs fois le même paragraphe
                If j < lMemNumPhraseMin Then _
                    lNumPhraseDebRech = lMemNumPhrasePreced: Exit For
                If oPhrase.lNumParagraphe < lNumParagraphe - iNbParagAv Then _
                    lNumPhraseDebRech = lMemNumPhrasePreced: Exit For
                lMemNumPhrasePreced = j: lNumPhraseDebRech = j
            Next j
            
            ' Rechercher les n° de § suivants
            lMemNumPhrasePreced = lNumPhrase1ParagMotTrouve
            lNumPhraseFinRech = lMemNumPhrasePreced
            For j = lNumPhrase1ParagMotTrouve To m_lNbPhrases
                Set oPhrase = m_colPhrases(j)
                If oPhrase.lNumParagraphe > lNumParagraphe Then _
                    lMemNumPhraseMin = j
                If oPhrase.lNumParagraphe > lNumParagraphe + iNbParagAp Then _
                    lNumPhraseFinRech = lMemNumPhrasePreced: Exit For
                ' Ne pas afficher 2x le dernier §
                If j = m_lNbPhrases Then lMemNumParagMax = oPhrase.lNumParagraphe
                lMemNumPhrasePreced = j: lNumPhraseFinRech = j
            Next j
            
            ' Présenter les résultats
            Dim lMemNumParagraphe0&, sResultat0$
            lMemNumParagraphe0 = 0
            sResultat0 = ""
            For j = lNumPhraseDebRech To lNumPhraseFinRech
                Set oPhrase = m_colPhrases(j)
                If oPhrase.lNumParagraphe > lMemNumParagraphe0 And _
                    lMemNumParagraphe0 > 0 Then
                    sResultat = sResultat & sAffCodeDoc & "§ n°" & _
                        lMemNumParagraphe0 & " : " & sResultat0 & vbCrLf
                    sResultat0 = ""
                End If
                sResultat0 = sResultat0 & oPhrase.sPhrase
                lMemNumParagraphe0 = oPhrase.lNumParagraphe
            Next j
            
            sResultat = sResultat & sAffCodeDoc & "§ n°" & oPhrase.lNumParagraphe & " : " & _
                sResultat0 & vbCrLf
                
            lMemNumParagraphe = lNumParagraphe
            
        End If

PhraseSuivante:
        'If Len(sResultat) > iMaxLongChaine Then Exit For
        If Len(sResultat) > iTailleLimite Then bTailleLimite = True: Exit For
        
        ' Afficher les résultats partiels au cours de la recherche
        If i Mod 10 = 0 Then CtrlResultat = sResultat

    Next i
    
    ' Afficher le résultat final si ce n'est pas déjà fait
    If CtrlResultat <> sResultat Then CtrlResultat = sResultat
    
    ' Ajouter le mot à la combobox des mots déjà recherchés,
    '  si ce n'est pas déjà fait
    Dim bDejaRecherche As Boolean
    For j = 0 To CtrlMot.ListCount - 1
        If LCase(CtrlMot.List(j)) = LCase(sMot) Then _
            bDejaRecherche = True: Exit For
    Next j
    If Not bDejaRecherche Then CtrlMot.AddItem sMot
    
FinRecherche:
    If m_bInterrompre Or bTailleLimite Then
        CtrlResultat = CtrlResultat & "..."
        AfficherMessage "Nombre d'occurrences affichées : " & i & " / " & _
            lNbPhrases & " trouvées"
    Else
        AfficherMessage "Nombre d'occurrences trouvées : " & lNbPhrases
    End If

Fin:
    Sablier bDesactiver:=True

End Sub

Public Function sCodeDocDefaut$()
    
    ' Générer un code document par défaut
    sCodeDocDefaut = sCodeDocDef & m_colDocs.Count + 1
    
End Function

Public Function bIndexerDocument(ByVal sCheminFichier$, _
    ByRef sCodeDoc$) As Boolean

    ' Ajouter un document et l'indexer

    m_bInterrompre = False
    
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    
    If sCodeDoc = "" Then sCodeDoc = sCodeDocDefaut()
    
    If Not bAjouterDocument(sCodeDoc, sCodeDoc, sCheminFichier) Then Exit Function
    m_bIndexModifie = True
    
    Sablier
    
    Dim iNumFichier%, sLigne$, sLignes$, i%, sMot$, iMemPosDebMot%
    Dim sPhrase$, iMemPosDebPhrase%, lNumParagraphe&, lNumPhraseL&
    Dim bNouvParag As Boolean
    lNumParagraphe = 0
    lNumPhraseL = 0
    iNumFichier = FreeFile
    Open sCheminFichier For Binary Access Read As #iNumFichier
    
    Dim lPosFinFichier&, lPosFichier&
    lPosFinFichier = LOF(iNumFichier)
    
    While Not EOF(iNumFichier)
    
        ' Afficher la progression de la lecture
        lPosFichier = Seek(iNumFichier)
        AfficherMessage "Indexation en cours... " & _
            Int(100! * lPosFichier / lPosFinFichier) & "%"
        If m_bInterrompre Then GoTo Fin
    
        On Error Resume Next
        Line Input #iNumFichier, sLigne
        Dim iErr%
        iErr = Err
        If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
        If iErr = 62 Then GoTo Suite ' Fin de fichier
        
        bNouvParag = True
        'm_lNbParagG = m_lNbParagG + 1
        
        Dim iLen%, sCar$
        iLen = Len(sLigne)
        iMemPosDebMot = 1: iMemPosDebPhrase = 1
        For i = 1 To iLen ' Parcours des lignes = paragraphes
            sCar = Mid$(sLigne, i, 1)
            If bSeparateurPhrases(sCar) Then
                ' Ajouter un mot à la fin d'une phrase et ajouter la phrase
                If Not bAjouter(sLigne, sCodeDoc, _
                    lNumParagraphe, m_lNbParagG, lNumPhraseL, _
                    i, iMemPosDebMot, iMemPosDebPhrase, bNouvParag, _
                    bAjouterMot0:=True, bAjouterPhrase0:=True) Then GoTo Fin
            ElseIf bSeparateurMots(sCar) Then
                ' Ajouter un mot en cours de phrase
                If Not bAjouter(sLigne, sCodeDoc, _
                    lNumParagraphe, m_lNbParagG, lNumPhraseL, _
                    i, iMemPosDebMot, iMemPosDebPhrase, bNouvParag, _
                    bAjouterMot0:=True, bAjouterPhrase0:=False) Then GoTo Fin
            End If
        Next i
            
        ' Ajouter un mot à la fin d'une phrase et
        '  ajouter la phrase à la fin d'un paragraphe
        If Not bAjouter(sLigne, sCodeDoc, _
            lNumParagraphe, m_lNbParagG, lNumPhraseL, _
            i, iMemPosDebMot, iMemPosDebPhrase, bNouvParag, _
            bAjouterMot0:=True, bAjouterPhrase0:=True) Then GoTo Fin
    
    Wend
    
Suite:
    ' Ajouter une phrase à la fin d'un fichier
    If Not bAjouter(sLigne, sCodeDoc, _
        lNumParagraphe, m_lNbParagG, lNumPhraseL, _
        i, iMemPosDebMot, iMemPosDebPhrase, bNouvParag, _
        bAjouterMot0:=False, bAjouterPhrase0:=True) Then GoTo Fin
    
    bIndexerDocument = True
    
    ' Faire une sauvegarde de l'index dans le fichier VBTextFinder.tmp
    'If bSauvegardeSecurite Then bSauvegarderIndex sFichierVBTxtFndTmp
    If bSauvegardeSecurite And False = m_bModeDirect Then bSauvegarderIndex m_sCheminVBTxtFndTmp
    
Fin:
    Close #iNumFichier
    AfficherMessage sMsgOperationTerminee
    Sablier bDesactiver:=True
    Exit Function
    
Erreur:
    AfficherMsgErreur Err, "AjouterEtIndexerDocument"
    GoTo Fin
    
End Function

Public Function bCodeDocExiste(sCodeDoc$) As Boolean
    
    ' Vérifier si un code document est déjà utilisé pour un des documents indexés
    
    ' On peut laisser un code document vide : un code numéroté sera généré par défaut
    If sCodeDoc = "" Then Exit Function
    On Error Resume Next
    Dim oDoc As clsDoc
    Set oDoc = m_colDocs.Item(sCodeDoc) ' Collection de documents indexés par sCodeDoc
    ' Référence inexistante dans la collection
    If Err = iErrCollectionCleAbsente Then GoTo Fin
    ' Code document existant
    bCodeDocExiste = True
Fin:
    On Error GoTo 0

End Function

Public Function bMotExiste(ByVal sMot$, ByRef oMot As clsMot, _
    Optional ByVal bSablier As Boolean = True) As Boolean
    
    ' Vérifier si un mot est indexé
    
    If sMot = "" Then Exit Function
    If bSablier Then Sablier
    On Error Resume Next
    Set oMot = m_colMots.Item(sMot) ' Collection de mots indexés par sMot
    If bSablier Then Sablier bDesactiver:=True
    Dim iErr%
    iErr = Err
    'If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    On Error GoTo 0
    'If iErr And iErr <> iErrCollectionCleAbsente Then GoTo Erreur
    ' Clé absente dans la collection
    If iErr = iErrCollectionCleAbsente Then Exit Function
    ' Mot existant
    bMotExiste = True
    
'Fin:
    'Exit Function
    
'Erreur:
    'AfficherMsgErreur Err, "bMotExiste"
    'Resume Fin

End Function

Private Function bAjouterDocument(ByVal sCleDoc$, sCodeDoc$, sCheminFichier$) As Boolean
    
    ' Ajouter un document à la liste des documents indexés
    
    ' Stocker les chemins en relatif le cas échéant
    Dim sFichier$, sChemin$, sCheminAIndexer$
    sChemin = sExtraireChemin(sCheminFichier, sFichier)
    If LCase(sChemin) = LCase(m_sCheminDossierCourant) Then
        sCheminAIndexer = sFichier
    Else
        sCheminAIndexer = sCheminFichier
    End If
    
    Const sMsgModeMultiDoc$ = "Indexation des documents"
    Dim oDoc0 As clsDoc
    For Each oDoc0 In m_colDocs
        If LCase(oDoc0.sChemin) = LCase(sCheminAIndexer) Then
            MsgBox "Le document :" & vbLf & oDoc0.sChemin & vbLf & _
                "a déjà été indexé", vbCritical, sMsgModeMultiDoc
            Exit Function
        End If
    Next oDoc0
    
    On Error Resume Next
    Dim oDoc As New clsDoc
    oDoc.sChemin = sCheminAIndexer
    oDoc.sCle = sCleDoc
    oDoc.sCodeDoc = sCodeDoc
    m_colDocs.Add oDoc, oDoc.sCle
    If Err = 457 Then
        ' Dans l'interface, on a déjà vérifié que la clé est unique
        MsgBox "La clé '" & sCodeDoc & "' a déjà été utilisée", _
            vbCritical, sMsgModeMultiDoc
        GoTo Fin
    ElseIf Err <> 0 Then
        AfficherMsgErreur Err, "bAjouterDocument", vbCritical, _
            sMsgModeMultiDoc
        GoTo Fin
    End If
    bAjouterDocument = True
    
Fin:
    On Error GoTo 0

End Function

Public Function bAjouter(ByRef sLigne$, ByVal sCodeDoc$, _
    ByRef lNumParagraphe&, ByRef lNbParagG&, _
    ByRef lNumPhraseL&, ByVal iPos%, _
    ByRef iMemPosDebMot%, ByRef iMemPosDebPhrase%, ByRef bNouvParag As Boolean, _
    bAjouterMot0 As Boolean, bAjouterPhrase0 As Boolean) As Boolean

    ' Ajouter un mot et/ou une phrase à indexer et
    '  positionner les variables d'indexation

    If bAjouterMot0 Then
        Dim sMot$
        sMot = Trim$(Mid$(sLigne, iMemPosDebMot, iPos - iMemPosDebMot))
        iMemPosDebMot = iPos + 1
        If sMot <> "" Then
            If bNouvParag Then
                bNouvParag = False
                lNumParagraphe = lNumParagraphe + 1
                lNbParagG = lNbParagG + 1
            End If
            If Not bAjouterMot(sMot, m_lNbPhrases + 1) Then Exit Function
            m_lNbMotsG = m_lNbMotsG + 1
        End If
    End If
                
    If bAjouterPhrase0 Then
        Dim sPhrase$
        sPhrase = Mid$(sLigne, iMemPosDebPhrase, iPos - iMemPosDebPhrase + 1)
        iMemPosDebPhrase = iPos + 1
        If sPhrase <> "" Then
            If bNouvParag Then
                bNouvParag = False
                lNumParagraphe = lNumParagraphe + 1
                lNbParagG = lNbParagG + 1
            End If
            If Not bAjouterPhrase(sPhrase, lNumParagraphe, sCodeDoc, lNumPhraseL) Then _
                Exit Function
        End If
    End If
    
    bAjouter = True

End Function

Private Function bAjouterPhrase(ByVal sPhrase$, lNumParagraphe&, sCodeDoc$, _
    ByRef lNbPhrasesL&)

    ' Ajouter une phrase à la liste des phrases indexées

    m_lNbPhrases = m_lNbPhrases + 1 ' Nombre de phrases globales
    lNbPhrasesL = lNbPhrasesL + 1   ' Nombre de phrases du document
    Dim oPhrase As clsPhrase
    
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    Set oPhrase = New clsPhrase
    oPhrase.lNumPhraseG = m_lNbPhrases
    oPhrase.lNumPhraseL = lNbPhrasesL
    oPhrase.sClePhrase = m_lNbPhrases
    oPhrase.lNumParagraphe = lNumParagraphe
    oPhrase.sCodeDoc = sCodeDoc
    oPhrase.sPhrase = sPhrase
    m_colPhrases.Add oPhrase, CStr(m_lNbPhrases)

    bAjouterPhrase = True
    
Fin:
    Exit Function

Erreur:
    ' Exemple d'erreur possible : mémoire insuffisante !
    AfficherMsgErreur Err, "bAjouterPhrase"
    Resume Fin

End Function

Private Function bAjouterMot(sMot$, lNumPhrases&) As Boolean
    
    ' Indexer les mots pour ne conserver que les mots distincts
    
    Dim sCle$, iErr%, mot As clsMot
    sCle = sMot
    
    On Error Resume Next
    Set mot = m_colMots.Item(sCle) ' Collection de mots indexés par sMot
    iErr = Err
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    
    If iErr And iErr <> iErrCollectionCleAbsente Then GoTo Erreur
    
    ' Clé absente dans la collection, on ajoute le mot
    If iErr = iErrCollectionCleAbsente Then
    
        Err.Clear
        Set mot = New clsMot
        With mot
            .sMot = sMot
            .lNbOccurences = 1
            .AjouterNumPhrase lNumPhrases '+ 1
        End With
        m_colMots.Add mot, sCle ' Ajout dans la collection
        
    Else ' Mot déjà existant : incrémenter le nombre d'occurrences
        
        With mot
            .lNbOccurences = .lNbOccurences + 1
            .AjouterNumPhrase lNumPhrases
        End With
        
    End If
    
    bAjouterMot = True
    
Fin:
    Exit Function
    
Erreur:
    ' Exemple d'erreur possible : mémoire insuffisante !
    AfficherMsgErreur Err, "bAjouterMot"
    Resume Fin

End Function

Private Function bSignesPonctuation(sMot$) As Boolean

    ' Indiquer si le mot ne contient que des signes de ponctuation

    Dim i%, iLen%
    iLen = Len(sMot)
    For i = 1 To iLen
        If Not bSignePonctuation(Mid$(sMot, i, 1)) Then Exit Function
    Next i
    bSignesPonctuation = True ' Ce mot ne contient que des signes de ponctuation

End Function

Private Function bSignePonctuation(sCar$) As Boolean
    
    ' Indiquer si le caractère est un signe de ponctuation
    
    Dim iCode%
    iCode = Asc(sCar)
    Select Case iCode
    Case Asc("A") To Asc("Z") ' Majuscule
    Case Asc("a") To Asc("z") ' Minuscule
    Case Else                 ' Ponctuation et chiffre
        bSignePonctuation = True
    End Select
     
End Function

Private Function bContientSeparateurPhrases(sMot$) As Boolean

    ' Indiquer si le mot contient un séparateur de phrases

    Dim i%, iLen%
    iLen = Len(sMot)
    For i = 1 To iLen
        If bSeparateurPhrases(Mid$(sMot, i, 1)) Then _
            bContientSeparateurPhrases = True: Exit Function
    Next i
    ' Ce mot n'en contient pas

End Function

Private Function bSeparateurPhrases(sCar$) As Boolean
    
    ' Indiquer si le caractère est un séparateur de phrases
    
    If InStr(sListeSeparateurPhrase, sCar) > 0 Then bSeparateurPhrases = True
     
End Function

Public Function bSeparateurMots(sCar$) As Boolean
    
    ' Indiquer si le caractère est un séparateur de mots
    If InStr(m_sListeSeparateursMot, sCar) > 0 Then bSeparateurMots = True
    'If bDebug Then Debug.Print "[" & sCar & "] : " & Asc(sCar)
     
End Function

Private Sub AfficherMessage(sMsg$)
    
    m_ctrlLblAvancement.Caption = sMsg
    ' Laisser du temps pour le traitement des messages : affichage du message et
    '  traitement du clic éventuel sur le bouton Interrompre
    DoEvents
    
End Sub

Private Function bValiderSauvegardeTmp() As Boolean
    
    ' Conserver la sauvegarde précédente (si elle existe) :
    '  renommer le fichier VBTextFinder.idx en VBTextFinder.bak
    ' Valider la sauvergarde temporaire (si elle n'existe pas, la créer) :
    '  renommer le fichier VBTextFinder.tmp en VBTextFinder.idx
    
    'Dim sCheminVBTxtFndTmp$, sCheminVBTxtFndBak$, sCheminVBTxtFndDat$
    'Dim sChemin$, sNomFichierSansExt$
    'sChemin = sExtraireChemin(m_sCheminFichierIndex, sNomFichier, _
        sExtension, sNomFichierSansExt)
    'sCheminVBTxtFndTmp = m_sCheminApp & "\" & sFichierVBTxtFndTmp
    'sCheminVBTxtFndBak = m_sCheminApp & "\" & sFichierVBTxtFndBak
    'sCheminVBTxtFndDat = m_sCheminApp & "\" & sFichierVBTxtFndDat
    
    ' Si le fichier .tmp n'existe pas, on sauvegarde l'index
    If Not bFichierExiste(m_sCheminVBTxtFndTmp) Then _
        If Not bSauvegarderIndex(m_sCheminVBTxtFndTmp) Then GoTo Fin
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    ' Renommer le fichier VBTextFinder.idx en VBTextFinder.bak
    If bFichierExiste(m_sCheminVBTxtFndIdx) Then
        ' Si le fichier .bak existe, le supprimer avant
        If bFichierExiste(m_sCheminVBTxtFndBak) Then Kill m_sCheminVBTxtFndBak
        Name m_sCheminVBTxtFndIdx As m_sCheminVBTxtFndBak
    End If
    ' Renommer le fichier VBTextFinder.tmp en VBTextFinder.idx
    Name m_sCheminVBTxtFndTmp As m_sCheminVBTxtFndIdx
    bValiderSauvegardeTmp = True
    
Fin:
    Exit Function

Erreur:
    AfficherMsgErreur Err, "bValiderSauvegardeTmp"
    Resume Fin

End Function

Private Function bSauvegarderIndex(sCheminFichierVBTxtFnd$) As Boolean
    
    ' Sauvegarder l'index dans le fichier VBTextFinder.idx
    
    Sablier
    
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    
    AfficherMessage "Sauvegarde de l'index en cours..."
    
    'Dim sChemin$
    'sChemin = m_sCheminApp & "\" & sFichierVBTxtFnd
    'If bFichierExiste(sChemin) Then Kill sChemin
    If bFichierExiste(sCheminFichierVBTxtFnd) Then Kill sCheminFichierVBTxtFnd
    
    Dim iNumFichier%
    iNumFichier = FreeFile
    Open sCheminFichierVBTxtFnd For Binary Access Write As #iNumFichier
    Put #iNumFichier, , rVersionFichierVBTxtFndDat
    
    ' Sauvegarder le nombre de documents indexés
    Dim lNbDocs&, lNbMots&
    lNbDocs = m_colDocs.Count
    lNbMots = m_colMots.Count
    Put #iNumFichier, , lNbDocs
    Put #iNumFichier, , lNbMots ' Nbr de mots distincts indexés
    Put #iNumFichier, , m_lNbPhrases
    ' Réserver de la place pour compléter les statistiques générales
    '  dans une version future (afin de conserver la compatibilité du fichier index)
    Put #iNumFichier, , m_lNbParagG ' Nbr de paragraphes indexés en tout
    Put #iNumFichier, , m_lNbMotsG  ' Nbr de mots indexés en tout
    ' Nombre de caractères dont les séparateurs de mot
    Put #iNumFichier, , CLng(0) 'lNbCarDontSeparIndexes
    'lNbCarIndexes : pas sauvé
    
    ' Sauvegarder la liste des documents indexés
    Dim oDoc As clsDoc
    For Each oDoc In m_colDocs
        bEcrireChaine iNumFichier, oDoc.sCle
        bEcrireChaine iNumFichier, oDoc.sCodeDoc
        bEcrireChaine iNumFichier, oDoc.sChemin
        ' Réserver de la place pour compléter les statistiques par document
        Put #iNumFichier, , CLng(0) 'oDoc.lNbMotsIndexes
        Put #iNumFichier, , CLng(0) 'oDoc.lNbPhrasesIndexees
        Put #iNumFichier, , CLng(0) 'oDoc.lNbParagIndexes
        Put #iNumFichier, , CLng(0) 'oDoc.lNbCarIndexes
        ' Nombre de caractères dont les séparateurs de mot
        Put #iNumFichier, , CLng(0) 'oDoc.lNbCarDontSeparIndexes
    Next oDoc
    
    ' Sauvegarder les mots de l'index
    Dim lNumMot&, i&, lNumPhraseG&
    Dim oMot As clsMot
    For Each oMot In m_colMots
    
        lNumMot = lNumMot + 1
        If lNumMot Mod iModuloAvanvement = 0 Or _
           lNumMot = lNbMots Then
            AfficherMessage "Sauvegarde des mots en cours : " & _
                lNumMot & " / " & lNbMots
            If m_bInterrompre Then GoTo Interruption
        End If
    
        bEcrireChaine iNumFichier, oMot.sMot
        Put #iNumFichier, , oMot.lNbOccurences
        Put #iNumFichier, , oMot.lNbPhrases
        ' Nombre de phrases dans lesquelles ce mot figure
        For i = 1 To oMot.lNbPhrases
            lNumPhraseG = oMot.lLireNumPhrase(i)
            Put #iNumFichier, , lNumPhraseG
        Next i
    Next oMot
    
    ' Sauvegarder les phrases de l'index
    Dim oPhrase As clsPhrase, lNumPhrase&
    For Each oPhrase In m_colPhrases
    
        lNumPhrase = lNumPhrase + 1
        If lNumPhrase Mod iModuloAvanvement = 0 Or _
           lNumPhrase = m_lNbPhrases Then
            AfficherMessage "Sauvegarde des phrases en cours : " & _
                lNumPhrase & " / " & m_lNbPhrases
            If m_bInterrompre Then GoTo Interruption
        End If
    
        bEcrireChaine iNumFichier, oPhrase.sClePhrase
        bEcrireChaine iNumFichier, oPhrase.sCodeDoc
        Put #iNumFichier, , oPhrase.lNumParagraphe
        Put #iNumFichier, , oPhrase.lNumPhraseG
        Put #iNumFichier, , oPhrase.lNumPhraseL
        bEcrireChaine iNumFichier, oPhrase.sPhrase
    Next oPhrase
    
    bSauvegarderIndex = True
    
Interruption:
    If m_bInterrompre Then
        Close
        ' Ne pas conserver un fichier partiel
        If bFichierExiste(sCheminFichierVBTxtFnd) Then Kill sCheminFichierVBTxtFnd
    End If
    
Fin:
    Close #iNumFichier
    Sablier bDesactiver:=True
    Exit Function
    
Erreur:
    Close ' Fermer tous les fichiers ouverts
    Sablier bDesactiver:=True
    AfficherMsgErreur Err, "bSauvegarderIndex"
    Resume Fin

End Function

Private Function bEcrireChaine(iNumFichier%, ByRef sChaine$) As Boolean
    
    ' Ecrire une chaîne de longueur variable dans un fichier binaire
    '  on utilise ByRef pour éviter de réallouer la chaîne en RAM
    
    Dim iLongChaine%
    iLongChaine = Len(sChaine)
    Put #iNumFichier, , iLongChaine
    Put #iNumFichier, , sChaine
    bEcrireChaine = True

End Function

Private Function bLireChaine(iNumFichier%, ByRef sChaine$) As Boolean
    
    ' Lire une chaîne de longueur variable dans un fichier binaire
    '  pour cela, il faut d'abord sauver la longueur de la chaîne
    
    Dim iLongChaine%
    ' Lire d'abord la longueur de la chaîne
    Get #iNumFichier, , iLongChaine
    If iLongChaine <= 0 Then Exit Function
    ' C'est surement une erreur si la chaîne est trop longue
    If iLongChaine > iMaxLongChaine Then Exit Function
    sChaine = Space(iLongChaine) ' = String(iLongChaine, " ")
    Get #iNumFichier, , sChaine
    bLireChaine = True

End Function

Public Function bLireIndex() As Boolean
    
    ' Lire l'index VBTextFinder.idx
    
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    
    Dim rVersionFichier!, lNbMots&
    If Not bFichierExiste(m_sCheminVBTxtFndIdx) Then GoTo Fin
    
    If m_bFichierIndexDef Then
        If vbNo = MsgBox("Voulez-vous recharger l'index :" & vbLf & _
            m_sCheminVBTxtFndIdx & " ?", vbYesNo Or vbQuestion, _
            sMsgGestionIndexDat) Then GoTo Fin
    End If
    
    Sablier
    AfficherMessage "Lecture de l'index en cours..."
    
    Dim iNumFichier%
    iNumFichier = FreeFile
    Open m_sCheminVBTxtFndIdx For Binary Access Read As #iNumFichier
    
    Dim sMsgErr$, sMsgErrLecture$
    sMsgErrLecture$ = "Version de fichier incorrecte : " & m_sCheminVBTxtFndIdx
    Get #iNumFichier, , rVersionFichier
    If rVersionFichier <> rVersionFichierVBTxtFndDat Then _
        sMsgErr = sMsgErrLecture: GoTo Erreur
    
    Dim lNbDocs&, lNb&
    Get #iNumFichier, , lNbDocs
    sMsgErrLecture$ = "Aucun document trouvé dans : " & m_sCheminVBTxtFndIdx
    If lNbDocs = 0 Then GoTo Erreur
    Get #iNumFichier, , lNbMots
    sMsgErrLecture$ = "Auncun mot trouvé dans : " & m_sCheminVBTxtFndIdx
    If lNbMots = 0 Then GoTo Erreur
    Get #iNumFichier, , m_lNbPhrases
    sMsgErrLecture$ = "Aucune phrase trouvée dans : " & m_sCheminVBTxtFndIdx
    If m_lNbPhrases = 0 Then GoTo Erreur
    Get #iNumFichier, , m_lNbParagG
    Get #iNumFichier, , m_lNbMotsG
    ' Place réservée pour compléter les statistiques générales dans une
    '  version future (afin de conserver la compatibilité du fichier index)
    ' Nombre de caractères dont les séparateurs de mot
    Get #iNumFichier, , lNb ' lNbCarDontSeparIndexes
    
    Dim i&, sCheminDoc$, sCle$, sCodeDoc$, oDoc As clsDoc
    sMsgErrLecture$ = "Impossible de lire les documents du fichier : " & m_sCheminVBTxtFndIdx
    For i = 0 To lNbDocs - 1
        
        Set oDoc = New clsDoc
        
        If Not bLireChaine(iNumFichier, sCle) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oDoc.sCle = sCle
        
        If Not bLireChaine(iNumFichier, sCodeDoc) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oDoc.sCodeDoc = sCodeDoc
        
        If Not bLireChaine(iNumFichier, sCheminDoc) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oDoc.sChemin = sCheminDoc
        
        ' Place réservée pour compléter les statistiques par document
        Get #iNumFichier, , lNb ' oDoc.lNbMotsIndexes = lNb
        Get #iNumFichier, , lNb ' oDoc.lNbPhrasesIndexees = lNb
        Get #iNumFichier, , lNb ' oDoc.lNbParagIndexes = lNb
        Get #iNumFichier, , lNb ' oDoc.lNbCarIndexes = lNb
        Get #iNumFichier, , lNb ' oDoc.lNbCarDontSeparIndexes = lNb
        
        If Not bAjouterDocument(sCle, sCodeDoc, sCheminDoc) Then GoTo Fin
        
    Next i
    
    Dim lPosFinFichier&, lPosFichier&
    lPosFinFichier = LOF(iNumFichier)
    
    Dim oMot As clsMot, sChaine$, j&, lNbPhrases&
    sMsgErrLecture$ = "Impossible de lire les mots du fichier : " & m_sCheminVBTxtFndIdx
    For i = 0 To lNbMots - 1
        
        ' Afficher la progression de la lecture
        If i Mod iModuloAvanvement = 0 Or i = lNbMots - 1 Then
            lPosFichier = Seek(iNumFichier)
            AfficherMessage "Lecture de l'index (mots) en cours... " & _
                Int(100! * lPosFichier / lPosFinFichier) & "%"
            If m_bInterrompre Then GoTo Fin
        End If
        
        Set oMot = New clsMot
        
        If Not bLireChaine(iNumFichier, sChaine) Then
            sMsgErr = sMsgErrLecture: GoTo Erreur
        End If
        oMot.sMot = sChaine
        
        Get #iNumFichier, , lNb
        oMot.lNbOccurences = lNb
        
        Get #iNumFichier, , lNb
        lNbPhrases = lNb
        ' Nombre de phrases dans lesquelles ce mot figure
        For j = 1 To lNbPhrases
            Get #iNumFichier, , lNb
            oMot.AjouterNumPhrase lNb
        Next j
        
        ' Si on reprend un index de la version en VB7, on n'est pas sûr que
        '  certains mots avec accent ou double caractères (oe dans coeur)
        '  soient déjà indexés
        If False = bMotExiste(oMot.sMot, oMot, bSablier:=False) Then _
            m_colMots.Add oMot, oMot.sMot ' Ajouter le mot dans la collection
    
    Next i
    
    sMsgErrLecture$ = "Impossible de lire les phrases du fichier : " & m_sCheminVBTxtFndIdx
    Dim oPhrase As clsPhrase
    For i = 0 To m_lNbPhrases - 1

        ' Afficher la progression de la lecture
        If i Mod iModuloAvanvement = 0 Or i = m_lNbPhrases - 1 Then
            lPosFichier = Seek(iNumFichier)
            AfficherMessage "Lecture de l'index (phrases) en cours... " & _
                Int(100! * lPosFichier / lPosFinFichier) & "%"
            If m_bInterrompre Then GoTo Fin
        End If

        Set oPhrase = New clsPhrase

        If Not bLireChaine(iNumFichier, sChaine) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oPhrase.sClePhrase = sChaine
        
        If Not bLireChaine(iNumFichier, sChaine) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oPhrase.sCodeDoc = sChaine

        Get #iNumFichier, , lNb
        oPhrase.lNumParagraphe = lNb
        Get #iNumFichier, , lNb
        oPhrase.lNumPhraseG = lNb
        Get #iNumFichier, , lNb
        oPhrase.lNumPhraseL = lNb

        If Not bLireChaine(iNumFichier, sChaine) Then _
            sMsgErr = sMsgErrLecture: GoTo Erreur
        oPhrase.sPhrase = sChaine
        
        m_colPhrases.Add oPhrase, CStr(oPhrase.lNumPhraseG)
        m_lNbPhrases = oPhrase.lNumPhraseG
        
    Next i
    
    bLireIndex = True
    m_bIndexModifie = False
    AfficherMessage sMsgOperationTerminee
    
Fin:
    Close #iNumFichier
    Sablier bDesactiver:=True
    Exit Function

Erreur:
    Close ' Fermer tous les fichiers ouverts
    Sablier bDesactiver:=True
    AfficherMsgErreur Err, "bLireIndex", sMsgErr
    If Err Then Resume Fin
    
Debug_:
    For i = 0 To m_colMots.Count - 1
        Debug.Print m_colMots(i + 1).sMot
    Next i

End Function

Private Function sLireCodeDocPhrase$(ByVal lNumPhrase&)
    
    ' Retourner le code document d'un numéro de phrase
    
    Dim oPhrase As clsPhrase
    Set oPhrase = m_colPhrases(lNumPhrase)
    sLireCodeDocPhrase = oPhrase.sCodeDoc

End Function

Private Function sLireCodeDoc$(ByVal sCleDoc$)
    
    ' Retourner le code document de la clé d'un document
    '  util maintenant pour trouver le code document via une clé de document
    '  associée à une phrase
    
    Dim oDoc As clsDoc
    Set oDoc = m_colDocs.Item(sCleDoc)
    sLireCodeDoc = oDoc.sCodeDoc

End Function
    
Public Sub CreerDocIndex(ByVal sTypeTri$) 'bTriFreq As Boolean)

    ' Fabriquer un index à partir de la collection de mots indexés

    If m_colDocs.Count = 0 Then Exit Sub

    ' Rattraper les erreurs (sauf en mode déboguage)
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0

    Dim bTriFreq As Boolean
    If sTypeTri = sTriFreq Then bTriFreq = True


    Dim oMot As clsMot, sFormatOcc$, i&
    If bTriFreq Then
        AfficherMessage ("Recherche du nombre d'occurrence max. pour la présentation...")
        ' Recherche du nbre d'occ max pour le format de présentation : nbre de 0
        Dim lMaxOcc&, sMaxFreq$
        For Each oMot In m_colMots
            If oMot.lNbOccurences > lMaxOcc Then lMaxOcc = oMot.lNbOccurences
        Next
        sMaxFreq = CStr(lMaxOcc)
        Dim iLenMaxFreq%
        iLenMaxFreq = Len(sMaxFreq)
        sFormatOcc = String(iLenMaxFreq, "0")
    End If


    Dim sTitre$, sExplication$, sListeMax$, iNbDocIndexes%
    iNbDocIndexes = m_colDocs.Count
    sTitre = "Document index de VBTextFinder"
    sListeMax = "liste des codes documents - " & iNbOccurrencesMaxListe & _
        " au max. - pour les mots non fréquents <= " & _
        iNbOccurencesMaxRecherchees & " occurrences"
    If iNbDocIndexes = 1 Then sListeMax = ""
    
    If bTriFreq Then
        sTitre = sTitre & " trié en fréquence"
        sExplication = "Explication : Nombre d'occurrences : Mot"
        If sListeMax <> "" Then sExplication = sExplication & " (" & sListeMax & ")"
    Else
        sTitre = sTitre & " trié par ordre alphabétique"
        sExplication = "Explication : Mot (nombre d'occurrences"
        If sListeMax <> "" Then sExplication = sExplication & " : " & sListeMax
        sExplication = sExplication & ")"
    End If
    sExplication = sExplication & vbLf
        
    ' Si le fichier existe, le supprimer avant
    Dim sCheminTxt$
    sCheminTxt = m_sCheminDossierCourant & "\" & sFichierVBTxtFndTxtAlphab
    If bTriFreq Then sCheminTxt = m_sCheminDossierCourant & "\" & sFichierVBTxtFndTxtFreq
    If False = bSupprimerFichier(sCheminTxt) Then GoTo Fin
        
    m_bInterrompre = False
    Sablier
        
    Dim lNumFichier&
    lNumFichier = FreeFile
    Open sCheminTxt For Output Access Write Lock Read Write As #lNumFichier
        
    Dim sMotGlossaire$, lNumMotIndexe&, lNbMotsIndexes&
    lNbMotsIndexes = m_colMots.Count
    lNumMotIndexe = 0
    For Each oMot In m_colMots
        
        lNumMotIndexe = lNumMotIndexe + 1
        If lNumMotIndexe Mod iModuloAvanvement = 0 Or _
           lNumMotIndexe = lNbMotsIndexes Or _
           lNumMotIndexe = 1 Then
            AfficherMessage "Création du document index en cours : " & _
                lNumMotIndexe & " / " & lNbMotsIndexes
            If m_bInterrompre Then Exit For
        End If
        
        Dim sListeRef$, lFin&
        Dim iNbOccEffectives%
        sListeRef = ""
        
        ' S'il n'y a qu'un seul document indexé, inutile d'indiquer toujours
        '  la même référence à ce document
        If iNbDocIndexes = 1 Then GoTo Suite
        
        lFin = oMot.lNbPhrases
        If lFin > iNbOccurencesMaxRecherchees Then GoTo Suite
        
        iNbOccEffectives = 0
        
        Dim lMemNumPhrase&, lNumPhrase&
        Dim sMemCodeDocPhrase$, sCodeDocPhrase$, sCodeDoc$
        lMemNumPhrase = -1: sMemCodeDocPhrase = ""
        For i = 1 To lFin
            lNumPhrase = oMot.lLireNumPhrase(i)
            If lNumPhrase = lMemNumPhrase Then GoTo RefSuiv
            lMemNumPhrase = lNumPhrase
            
            sCodeDocPhrase = sLireCodeDocPhrase(lNumPhrase)
            If sCodeDocPhrase = sMemCodeDocPhrase Then GoTo RefSuiv
            sMemCodeDocPhrase = sCodeDocPhrase
            
            sCodeDoc = sLireCodeDoc(sCodeDocPhrase)
            
            iNbOccEffectives = iNbOccEffectives + 1
            If iNbOccEffectives > iNbOccurrencesMaxListe Then _
                sListeRef = sListeRef & ", ...": Exit For
                
            If i = 1 Then
                sListeRef = sCodeDoc
            Else
                sListeRef = sListeRef & ", " & sCodeDoc
            End If
RefSuiv:
        Next i
            
Suite:
        If bTriFreq Then
            ' Tri fréquentiel : on met le nombre d'occurence du mot en premier
            If sListeRef <> "" Then sListeRef = " (" & sListeRef & ")"
            sMotGlossaire = Format(oMot.lNbOccurences, sFormatOcc) & " : " & oMot.sMot & sListeRef
        Else
            ' Tri alphabétique : on met le mot en premier
            If sListeRef <> "" Then sListeRef = " : " & sListeRef
            sMotGlossaire = oMot.sMot & " (" & oMot.lNbOccurences & sListeRef & ")"
        End If
        
        Print #lNumFichier, sMotGlossaire
        
    Next oMot
    Close #lNumFichier
        
    
    ' Liaison tardive : déclaration au moment de l'exécution
    ' - avantage : le programme ne plante pas si Word n'est pas installé,
    '   et un message approprié peut donc être affiché ; marche avec toutes les versions
    '   de Word qui gèrent le code VBA utilisé ici (je n'ai pas testé avec Word 97)
    ' - inconvénient : pas d'intellisense pour déboguer, et le code est moins clair
    
    Dim oWrd As Object
    
    Const wdDoNotSaveChanges& = 0
    Const wdSortFieldAlphanumeric& = 0
    Const wdSortOrderAscending& = 0
    Const wdFrench& = 1036
    Const wdSortFieldNumeric& = 1
    Const wdSortOrderDescending& = 1
    Const wdPageFitBestFit& = 2
    Const wdSortSeparateByDefaultTableSeparator& = 2

    AfficherMessage "Ouverture de Microsoft Word..."
    ' Instanciation de l'objet Word : vérifier si Word est déjà ouvert
    '  (note : si VB6 dit que oWrd n'est pas défini, quitter et relancer VB6)
    ' Note : si vous arretez le programme avec VB6 en mode débug, alors il faudra
    '  tuer l'instance de Word dans le gestionnaire des tâches
    Dim bObjetDejaOuvert As Boolean
    If Not bCreerObjet(oWrd, sClasseObjetWord, bObjetDejaOuvert) Then GoTo Fin
        
    oWrd.Visible = False
    ' 1252 : Code page de Windows (Latin) pour récupérer les caractères unicodes,
    '  par exemple coeur avec oe collé
    oWrd.Documents.Open sCheminTxt, Encoding:=1252
    
    ' Trier le nouveau document Word par ordre alphabétique ou bien numérique
    If lNbMotsIndexes = 0 Then GoTo Presentation
    
    With oWrd.Selection
        
        AfficherMessage "Tri du document index..."
        .WholeStory ' Sélection de tout le document

        If bTrapErr Then On Error Resume Next Else On Error GoTo 0
        If bTriFreq Then
            Const sChamp1$ = "Paragraphes" '"Champ 1" ' "Mot 1"
            Const sChamp2$ = "" '"Champ 2" ' "Mot 3"
            .Sort ExcludeHeader:=False, FieldNumber:=sChamp1, _
                SortFieldType:=wdSortFieldAlphanumeric, _
                SortOrder:=wdSortOrderDescending, _
                FieldNumber2:=sChamp2, _
                SortFieldType2:=wdSortFieldAlphanumeric, _
                SortOrder2:=wdSortOrderAscending, _
                Separator:=wdSortSeparateByDefaultTableSeparator, _
                SortColumn:=False, CaseSensitive:=False, LanguageID:=wdFrench
                
        Else
            .Sort ExcludeHeader:=False, FieldNumber:="Paragraphes", _
                SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending, _
                Separator:=wdSortSeparateByDefaultTableSeparator, SortColumn:=False, _
                CaseSensitive:=False, LanguageID:=wdFrench
        
        End If
        If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
        
        .InsertBefore vbLf
        
Presentation:
        ' Présentation du glossaire
        .InsertBefore sExplication
        .InsertBefore "Nombre de mots distincts indexés : " & lNbMotsIndexes & vbLf
        
        ' Afficher la liste des documents indexés
        Dim oDoc As clsDoc
        For Each oDoc In m_colDocs
            .InsertBefore oDoc.sChemin & " (" & oDoc.sCodeDoc & ")" & vbLf
        Next oDoc
        .InsertBefore "Liste des documents indexés :" & vbLf
        
        If m_bInterrompre Then .InsertBefore "(création du document index interrompue)" & vbLf
            
        .InsertBefore sTitre & vbLf
        
PresentationSuite:
        
        ' Mettre en largeur de page maximale pour améliorer la lisibilité du document
        oWrd.ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit
        
        ' Enlever la mise en forme Texte brut
        'oWrd.Visible = True
        .WholeStory
        .ClearFormatting
        ' Il ne suffit pas d'effacer la mise en forme,
        '  il faut aussi demander d'oublier le format Brut au moment de sauver
        
        Const wdStory& = 6
        .HomeKey Unit:=wdStory ' Revenir au début du document
        
    End With
        
    ' Si le document est déjà ouvert, on ne peut le sauver
    '  on vérifie donc s'il est ouvert, et on le ferme le cas échéant
    ' Cela ne fonctionne que s'il s'agit de la même instance de Word
    '  sinon on obtient un message d'erreur comme quoi le fichier est vérrouillé
    On Error Resume Next
    Dim sCheminDoc$
    sCheminDoc = m_sCheminDossierCourant & "\" & _
        sFichierVBTxtFndDocAlphab
    If bTriFreq Then sCheminDoc = m_sCheminDossierCourant & "\" & _
        sFichierVBTxtFndDocFreq
    oWrd.Documents(sCheminDoc).Close
    If False = bSupprimerFichier(sCheminDoc) Then GoTo Erreur2
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    Const wdFormatDocument& = 0
    oWrd.ActiveDocument.SaveAs sCheminDoc, FileFormat:=wdFormatDocument
        
Fin0:
    AfficherMessage "Création du document index terminée."
    oWrd.Visible = True

Fin:
    Close
    On Error GoTo 0
    Sablier bDesactiver:=True
    Exit Sub

Erreur:
    Sablier bDesactiver:=True
    AfficherMsgErreur Err, "CreerDocIndex"
    
Erreur2:
    If False = (oWrd Is Nothing) Then
        If bObjetDejaOuvert Then
            oWrd.Visible = True
        Else
            On Error Resume Next
            oWrd.Quit wdDoNotSaveChanges
            Set oWrd = Nothing
            On Error GoTo 0
        End If
    End If
    GoTo Fin
    
End Sub

Public Function ListerDocumentsIndexes(CtrlResultat As Control)

    ' Afficher la liste des documents indexés
    
    Dim oDoc As clsDoc, sResultat$
    sResultat = "Nombre de mots indexés : " & m_lNbMotsG & vbCrLf
    sResultat = sResultat & "Nombre de mots distincts indexés : " & _
        m_colMots.Count & vbCrLf
    sResultat = sResultat & "Nombre de phrases indexées : " & _
        m_colPhrases.Count & vbCrLf
    sResultat = sResultat & "Nombre de paragraphes indexés : " & _
        m_lNbParagG & vbCrLf
    sResultat = sResultat & vbCrLf
    sResultat = sResultat & "Liste des documents indexés (" & _
        m_colDocs.Count & ") :" & vbCrLf
    For Each oDoc In m_colDocs
        sResultat = sResultat & oDoc.sChemin & " (" & oDoc.sCodeDoc & ")" & vbCrLf
    Next oDoc
    CtrlResultat = sResultat

End Function

Public Function bConvertirDocEnTxt() As Boolean

    ' Convertir un fichier .doc ou .html en .txt

    Dim sExtension$, sNomFichier$, sChemin$, sFichierTxt$
    sChemin = sExtraireChemin(m_sCheminFichierTxtDirect, sNomFichier, sExtension)
    sExtension = LCase(sExtension)
    ' On laisse le fichier inchangé si on ne peut pas le convertir avec Word
    If Not (sExtension = "doc" Or sExtension = "htm" Or sExtension = "html") Then _
        bConvertirDocEnTxt = True: Exit Function
    
    ' Dans la version Word, le document pourrait s'indexer lui-même,
    '  mais il ne faudrait pas utiliser la macro AutoOpen
    '  ni fermer le document actif avec Wrd.ActiveDocument.Close
    If m_sCheminFichierTxtDirect = sLireCheminFichierApplication() Then _
        MsgBox "L'Auto-indexation ne fonctionne pas": Exit Function
    
    Sablier
    sFichierTxt = sChemin & "\" & _
        Left$(sNomFichier, Len(sNomFichier) - Len(sExtension)) & "txt"
    
    Dim oWrd As Object, bWrdDejaOuvert As Boolean
    Const wdCRLF& = 0
    Const wdFormatText& = 2
    
    AfficherMessage "Ouverture de Microsoft Word..."
    
    If Not bCreerObjet(oWrd, sClasseObjetWord, bWrdDejaOuvert) Then GoTo Fin
    If bTrapErr Then On Error GoTo Erreur Else On Error GoTo 0
    If bWrdDejaOuvert Then oWrd.Visible = False
    AfficherMessage "Ouverture du fichier " & m_sCheminFichierTxtDirect & "..."
    oWrd.Documents.Open m_sCheminFichierTxtDirect
    AfficherMessage "Conversion en .txt du fichier " & m_sCheminFichierTxtDirect & "..."
    oWrd.ChangeFileOpenDirectory m_sCheminDossierCourant
    oWrd.ActiveDocument.SaveAs FileName:=sFichierTxt, FileFormat:=wdFormatText, _
        LineEnding:=wdCRLF
    oWrd.ActiveDocument.Close
    ' Fermer Word s'il n'était pas ouvert auparavant
    If bWrdDejaOuvert Then
        oWrd.Visible = True
    Else
        oWrd.Quit
    End If
    m_sCheminFichierTxtDirect = sFichierTxt
    AfficherMessage "Conversion en .txt terminée."
    bConvertirDocEnTxt = True

Fin:
    On Error GoTo 0
    'Sablier bDesactiver:=True ' A faire plus tard lorsque l'application rend la main
    Exit Function

Erreur:
    Sablier bDesactiver:=True
    AfficherMsgErreur Err, "bConvertirDocEnTxt"
    GoTo Fin

End Function

Public Sub Sablier(Optional bDesactiver As Boolean = False)

    #If bVersionVBA Then
        Const lSablierDesactive& = fmMousePointerDefault
        Const lSablier& = fmMousePointerHourGlass
    #Else
        Const lSablierDesactive& = vbDefault
        Const lSablier& = vbHourglass
    #End If

    If bDesactiver Then
        m_oFrmMousePointer.MousePointer = lSablierDesactive
    Else
        m_oFrmMousePointer.MousePointer = lSablier
    End If

End Sub

Public Function bHyperTexte(sMotSel$, ByRef sMotSelFin$) As Boolean

    ' Traitement du mode hypertexte
    
    Dim iDeb%, iFin%, iLongMot%, iLong%, sCar1$, sCar2$
    If sMotSel = "" Then Exit Function
    
    ' Extraction d'un mot bien délimité
    ' Vérifier s'il y a une virgule à la fin
    If bSeparateurMots(Right$(sMotSel, 1)) Then _
        sMotSel = Left$(sMotSel, Len(sMotSel) - 1)
    iLongMot = Len(sMotSel)
    For iDeb = 1 To iLongMot - 1
        ' Vérifier s'il y a : l'
        sCar1 = Mid$(sMotSel, iDeb, 1)
        sCar2 = Mid$(sMotSel, iDeb + 1, 1)
        If Not bSeparateurMots(sCar1) And _
           (Not bSeparateurMots(sCar2) Or _
            iLongMot < 4) Then Exit For
    Next iDeb
    If iDeb = iLongMot Then sMotSelFin = Right$(sMotSel, 1): Exit Function
    ' Vérifier les mots composés : c'est-à-dire : est
    For iFin = iDeb + 1 To iLongMot
        sCar1 = Mid$(sMotSel, iFin, 1)
        If bSeparateurMots(sCar1) Then Exit For
    Next iFin
    sMotSelFin = Mid$(sMotSel, iDeb, iFin - iDeb)
    ' Inutile de lancer une recherche automatique pour des mots de moins de 3 lettres
    If Len(Trim$(sMotSelFin)) < 3 Then Exit Function
    bHyperTexte = True

End Function
