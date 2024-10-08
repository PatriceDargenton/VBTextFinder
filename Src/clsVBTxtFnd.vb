
' Fichier clsVBTxtFnd.vb
' ----------------------

Imports System.Text ' Pour StringBuilder
'Imports System.Text.Encoding ' Pour GetEncoding
Imports System.Collections.Specialized.CollectionsUtil ' Pour CreateCaseInsensitiveHashtable

Friend Class clsVBTextFinder ' Classe principale du moteur de recherche VBTextFinder

#Region "Interface publique"

    Public Delegate Sub GestEvAfficherMessage(sender As Object,
        e As clsMsgEventArgs)
    Public Event EvAfficherMessage As GestEvAfficherMessage

    Public m_sCheminDossierCourant$

    ' Types d'index
    Public Const sIndexAlpha$ = "Alphabétique"
    Public Const sIndexFreq$ = "Fréquentiel"
    Public Const sIndexMotsCles$ = "Mots clés"
    Public Const sIndexCitations$ = "Citations"
    Public Const sIndexSimple$ = "Simple"
    Public Const sIndexSimpleComparer$ = "Simple : comparer"
    Public Const sIndexTout$ = "Tout"
    Public Const sIndexEspacesInsecables$ = "Espaces insécables"
    Public Const sIndexEspacesInsecablesAVerifier$ = "Esp. inséc. à vérif."
    Public Const sIndexMajuscules$ = "Majuscules"
    Public Const sIndexAccents$ = "Accents manquants" ' 06/06/2019
    ' Analyse de la fréquence de successions des lettres dans les mots : Projet en cours
    Public Const sIndexNGrammes$ = "N-Grammes"
    Public Const sPrefixeIndex$ = "Index" ' Ne pas indexer tout fichier commençant par Index
    Public Const sPrefixeIndexSimple$ = "IndexSimple"
    Public Const sPrefixeIndexCitations$ = "IndexCitations"
    Public Const sPrefixeEspacesInsecables$ = "EspacesInsecables"
    Public Const sPrefixeMajuscules$ = "Majuscules"

    ' Types d'affichage des résultats de recherche
    Public Const sAfficherPhrase$ = "Phrase"
    Public Const sAfficherParag$ = "Paragraphe"
    Public Const sAfficherParagPM1$ = "Paragraphe +-1"
    Public Const sAfficherParagPM2$ = "Paragraphe +-2"
    Public Const sAfficherParagPM3$ = "Paragraphe +-3"

    ' Indexation directe d'un fichier texte passé 
    '  en argument de la ligne de commande
    Public m_bModeDirect As Boolean = False
    Public m_sCheminFichierTxtDirect$ = ""

    Public iNbPhrasesG% ' (= m_colPhrases.Count)
    Public iNbMotsG%  ' Nombre de mots indexés en tout
    Public iNbParagG% ' Nombre de paragraphes indexés en tout (sans les lignes vides)

    Public m_sbResultatHtml As StringBuilder
    Public m_sbResultatTxt As StringBuilder

    ' sMotsCourants ne contient pas les accents : 
    '  les mots clés ne fonctionneront plus si on indexe les accents
    Private m_bIndexerAccents As Boolean = False
    Private m_styleCompare% = StringComparison.InvariantCultureIgnoreCase
    Private m_styleCompare2 As System.StringComparison = StringComparison.InvariantCultureIgnoreCase
    Public Property IndexerAccents As Boolean
        Get
            Return m_bIndexerAccents
        End Get
        Set(value As Boolean)
            'Dim bAccent = m_bIndexerAccents
            m_bIndexerAccents = value
            ' Non, en fait cela n'a pas d'impact sur l'indexation, 
            '  car on n'enlève les accents par le code, pas par une option de comparaison :
            '  (à faire seulement si on veut distinguer la casse, pas les accents)
            'If m_bIndexerAccents Then
            '    m_styleCompare = StringComparison.InvariantCulture
            '    m_styleCompare2 = StringComparison.InvariantCulture
            'Else
            '    m_styleCompare = StringComparison.InvariantCultureIgnoreCase
            '    m_styleCompare2 = StringComparison.InvariantCultureIgnoreCase
            'End If
            'If bAccent <> m_bIndexerAccents Then
            '    m_htMots = New Hashtable(m_styleCompare)
            'End If
        End Set
    End Property

    ' Hashtable des mots indexés avec pour clé : sMot sans accent (par défaut)
    ' Si on indexe les accents (bIndexerAccents = True)
    '  le constructeur par défaut de Hastable ne fonctionne pas avec 
    '  des mots tels que "Drôle" (et InvariantCulture ne suffit pas)
    ' Avec ces 2 paramètres, Drôle est bien trouvé, et il est bien distinct de Drole
    'Private m_htMots As New Hashtable( _
    '    CaseInsensitiveHashCodeProvider.Default, _
    '    CaseInsensitiveComparer.Default)
    ' BC40000
    Private m_htMots As New Hashtable(m_styleCompare)
    'Private m_htMots As Hashtable() ' 11/12/2022 Ne compile pas si on n'instancie pas ?

    ' Hashtable du dictionnaire des mots existants
    Private m_htDico As Hashtable

    ' Si Unicode alors conserver les accents et tous les caractères exotiques
    Public m_bOptionTexteUnicode As Boolean
    Private Const iMasqueOptionUnicode% = 1
    Private Const iMasqueOptionAccent% = 2

    Public m_bOccurrencesEnGras As Boolean = False
    Public m_bOccurrencesEnCouleurs As Boolean = True
    Public m_sCouleursHtml$ = sCouleursHtmlDef

    Public m_bIndexerChapitre As Boolean = False
    Public m_sChapitrage$ = sChapitrageDef
    Public m_sChapitrageMdb$ = sChapitrageMdbDef
    Public m_sChapitrageXL$ = sChapitrageXLDef
    'Public m_bAfficherChapitre As Boolean = True ' Utile ?
    ' Afficher aussi les chapitres dans les index. alphab. et fréq.
    Public m_bAfficherChapitreIndex As Boolean = False

    ' Afficher les n° de § et de phrase global sinon local à chaque document
    ' Note : si on affiche les n° local, on ne peut pas restaurer la position
    '  du curseur
    Public m_bNumerotationGlobale As Boolean = True

#End Region

#Region "Déclarations"

    Private Const rVersionFichierVBTxtFndIdx10! = 1.0!
    Private Const rVersionFichierVBTxtFndIdx! = 1.15!

    ' Fichiers de sauvegarde de l'index
    Private Const sExtVBTF$ = ".idx" '".dat"
    Private Const sFichierVBTxtFndIdxDef$ = "VBTextFinder" & sExtVBTF ' Sauvegarde en cours
    Private Const sMsgGestionIndex$ = "Gestion du fichier d'index " & sFichierVBTxtFndIdxDef

    ' Pour supprimer les n° des notes de bas de page
    ' (cela coûte 8% du temps d'indexation supplémentaire, ça va)
    Private Const bSupprimerNumeriquesEnFinDeMot = True ' 16/12/2022 Faire une option

    Private m_sCheminVBTxtFndTmp$, m_sCheminVBTxtFndBak$, m_sCheminVBTxtFndIdx$
    Private m_sCheminFichierIndex$ ' Chemin complet avec extension du fichier index

    Private m_sCheminFichierIni$

    Private m_bFichierIndexDef As Boolean

    ' Fabrication du document index
    Private Const sFichierVBTxtFnd$ = "VBTextFinder"
    Private Const sFichierVBTxtFndAlphab$ = "VBTextFinderAlphab"
    Private Const sFichierVBTxtFndFreq$ = "VBTextFinderFreq"
    Private Const sFichierVBTxtFndMotsCles$ = "VBTextFinderMotsCles"
    Private Const sFichierVBTxtFndTout$ = "VBTextFinderTout"
    Private Const sFichierIni$ = "VBTextFinder"

    Private Const sTriDef$ = sIndexAlpha
    Private Const sAfficherDef$ = sAfficherPhrase
    Private Const sCodeDocDef$ = "Doc n°"

    ' Booléen pour pouvoir interrompre une longue opération
    Private m_bInterrompre As Boolean

    Private m_bSablierDesactive As Boolean

    Private m_msgDelegue As clsMsgDelegue

    Private m_sListeSeparateursMot$, m_sListeSeparateursPhrase$

    ' Booléen pour savoir si l'index est modifié,
    '  auquel cas il doit être sauvé lors de la fermeture de l'application
    Private m_bIndexModifie As Boolean

    ' Collection de phrases indexées par leur numéro (un tableau suffirait dans ce cas)
    ' Par rapport à une Collection VB6 , les indices d'une ArrayList commencent à 0 au lieu de 1
    Private m_colPhrases As New ArrayList 'Collection

    ' Collection des documents indexés
    Private m_colDocs As New Collection ' Hashtable : perd l'ordre : dommage !
    'Private m_colDocs As Collections.Specialized.NameObjectCollectionBase ' Hashtable+ArrayList

    Private m_colDocsIni As New Collection ' Reste des codes doc dans le ini

    Const sTagUnicodeIni$ = "Unicode"

    'Private m_colChapitres As New Collection ' sCleChapitre -> clsChapitre
    Private m_sbChapitres As New StringBuilder

    Private tsDiffTps As New TimeSpan

    ' Conversion à la volée : noter si le fichier txt converti existait avant
    '  pour savoir s'il faut le supprimer en quittant
    ' Liste des fichiers txt à supprimer en quittant
    Private m_alsCheminsFichierTxt As New ArrayList

    Private m_sMemExpression$ = ""
    Private m_alExpressions As ArrayList

    ' Tous ces n° sont globaux sur l'ensemble des documents
    Private m_iNumParagSel%, m_iNumPhraseSel%, m_iNumCarSel%, m_iLongSel%

    Private Const sIndicParag$ = "§ n°"
    Private Const sCarParag$ = "§"
    Private Const sIndicPhrase$ = "Ph. n°"
    Private Const sSepIni$ = "|" ' ":"

    Public m_bAuMoinsUnTxtUnicode As Boolean = False
    ' INFORMATION : Le texte contient des caractères Unicode et l'option n'est pas activée 
    ' (ces caractères seront remplacés par des signes '?')
    Public m_bAvertAuMoinsUnTxtUnicode As Boolean = False
    ' Information : Le texte ne contient pas de caractères Unicode (alors que l'option est activée)
    Public m_bInfoAuMoinsUnTxtNonUnicode As Boolean = False

#End Region

#Region "Initialisation et gestion du formulaire"

    Public Sub Initialiser(msgDelegue As clsMsgDelegue,
        ByRef ctrlLstAfficher As System.Windows.Forms.ListBox,
        ByRef ctrlTypeIndex As System.Windows.Forms.ListBox, iTypeIndexSelect%)

        ' 23/11/2018
        m_bAuMoinsUnTxtUnicode = False
        m_bAvertAuMoinsUnTxtUnicode = False
        m_bInfoAuMoinsUnTxtNonUnicode = False

        Me.m_msgDelegue = msgDelegue

        ' Initialisation des contrôles de l'interface
        ctrlLstAfficher.Items.Add(sAfficherPhrase)
        ctrlLstAfficher.Items.Add(sAfficherParag)
        ctrlLstAfficher.Items.Add(sAfficherParagPM1)
        ctrlLstAfficher.Items.Add(sAfficherParagPM2)
        ctrlLstAfficher.Items.Add(sAfficherParagPM3)
        ctrlLstAfficher.SetSelected(0, True)

        ctrlTypeIndex.Items.Add(sIndexAlpha)
        ctrlTypeIndex.Items.Add(sIndexFreq)
        ctrlTypeIndex.Items.Add(sIndexMotsCles)
        ctrlTypeIndex.Items.Add(sIndexCitations)
        ctrlTypeIndex.Items.Add(sIndexSimple)
        ctrlTypeIndex.Items.Add(sIndexSimpleComparer)
        ctrlTypeIndex.Items.Add(sIndexTout)
        If Not bSupprimerEspInsec Then
            ctrlTypeIndex.Items.Add(sIndexEspacesInsecables)
            ctrlTypeIndex.Items.Add(sIndexEspacesInsecablesAVerifier)
        End If

        ctrlTypeIndex.Items.Add(sIndexMajuscules) ' 26/03/2016
        ctrlTypeIndex.Items.Add(sIndexAccents) ' 06/06/2019
        'If bDebug Then ctrlTypeIndex.Items.Add(sIndexNGrammes)
        'ctrlTypeIndex.SetSelected(0, True)
        ctrlTypeIndex.SetSelected(iTypeIndexSelect, True)

        ' Chemin par défaut s'il n'y a pas d'arg en ligne de cmd
        m_sCheminFichierIndex = Application.StartupPath & "\" & sFichierVBTxtFndIdxDef
        m_bFichierIndexDef = True

        Dim sArgument$
        ' Extraire les options passées en argument de la ligne de commande
        ' Ne fonctionne pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command
        ' Extraire l'option passée en argument de la ligne de commande
        Dim sNomFichierSansExt$ = ""
        If sArg0.Length > 0 Then
            Dim asArgs$() = asArgLigneCmd(sArg0)
            If asArgs.Length > 0 Then
                sArgument = asArgs(0)
                If bDossierExiste(sArgument) Then
                    ' Si le dossier existe, alors cela signifie qu'il s'agit d'un dossier
                    '  dans ce cas indexer tous les documents du dossier
                    m_sCheminDossierCourant = sEnleverSlashFinal(sArgument)
                    sNomFichierSansExt = sNomDossierFinal(sArgument)
                    GoTo Suite
                End If
                If bFichierExiste(sArgument) Then
                    Dim sExt$ = IO.Path.GetExtension(sArgument).ToLower
                    If sExt = sExtVBTF Then
                        m_sCheminFichierIndex = sArgument : m_bFichierIndexDef = False
                    Else 'If sExt = sExtTxt Or sExt = sExtDoc Or sExt.StartsWith(sExtHtm) Then
                        ' 29/05/2015 Accepter tout types de fichier
                        m_bModeDirect = True
                        m_sCheminFichierTxtDirect = sArgument
                        m_sCheminFichierIndex =
                            sExtraireChemin(m_sCheminFichierTxtDirect) & "\" &
                            sFichierVBTxtFndIdxDef
                    End If
                Else
                    MsgBox("Impossible de trouver le fichier :" & vbLf &
                        sArgument, MsgBoxStyle.Critical,
                        "Passage d'un fichier au démarrage de VBTextFinder")
                End If
            End If
        End If

        Dim sNomFichier$ = "", sExtension$ = ""
        m_sCheminDossierCourant = sExtraireChemin(m_sCheminFichierIndex, sNomFichier,
            sExtension, sNomFichierSansExt)
Suite:
        m_sCheminVBTxtFndTmp = m_sCheminDossierCourant & "\" & sNomFichierSansExt & ".tmp"
        m_sCheminVBTxtFndBak = m_sCheminDossierCourant & "\" & sNomFichierSansExt & ".bak"
        m_sCheminVBTxtFndIdx = m_sCheminDossierCourant & "\" & sNomFichierSansExt & sExtVBTF
        m_sCheminFichierIni = m_sCheminDossierCourant & "\" & sFichierIni & ".ini"

        Me.iNbPhrasesG = 0

        Dim sChemin = Application.StartupPath & Config.sCheminSeparateursMot
        'MsgBox("Chemin sep mot : " & sChemin)
        If bFichierExiste(sChemin) Then
            Me.m_sListeSeparateursMot = sLireFichier(sChemin)
            ' Dans le fichier, tous les caractères sont bien conservés, sauf l'espace insécable
            ' 20/09/2009 Maintenant c'est nécessaire d'inclure l'espace insécable, car il est conservé Doc->Txt
            Me.m_sListeSeparateursMot &= Chr(iCodeASCIIEspaceInsecable)
        Else
            Me.m_sListeSeparateursMot = Config.sListeSeparateursMot
            Me.m_sListeSeparateursMot &=
                Chr(iCodeASCIITabulation) &
                Chr(iCodeASCIIGuillemet) &
                Chr(iCodeASCIIGuillemetOuvrant) & Chr(iCodeASCIIGuillemetFermant) &
                Chr(iCodeASCIIEspaceInsecable)
            ' 20/09/2009 Maintenant c'est nécessaire d'inclure l'espace insécable, car il est conservé Doc->Txt
        End If
        ' 15/09/2018
        If Me.m_bOptionTexteUnicode Then
            Me.m_sListeSeparateursMot &= ChrW(iCodeUTF16EspaceFineInsecable)
            Me.m_sListeSeparateursMot &= ChrW(iCodeUTF16EspaceInsecable) ' 13/07/2019
        End If


        sChemin = Application.StartupPath & Config.sCheminSeparateursPhrase
        If bFichierExiste(sChemin) Then
            Me.m_sListeSeparateursPhrase = sLireFichier(sChemin)
        Else
            Me.m_sListeSeparateursPhrase = Config.sListeSeparateursPhrase
        End If

        sChemin = Application.StartupPath & Config.sCheminChapitrage
        If bFichierExiste(sChemin) Then Me.m_sChapitrage = sLireFichier(sChemin)
        sChemin = Application.StartupPath & Config.sCheminChapitrageExcel
        If bFichierExiste(sChemin) Then Me.m_sChapitrageXL = sLireFichier(sChemin)
        sChemin = Application.StartupPath & Config.sCheminChapitrageAccess
        If bFichierExiste(sChemin) Then Me.m_sChapitrageMdb = sLireFichier(sChemin)

        Me.tsDiffTps = New TimeSpan(0)

    End Sub

    Public Function bQuitter() As Boolean

        Dim lNbDocs As Integer
        lNbDocs = Me.m_colDocs.Count()
        Dim iReponse% ' As Short
        If lNbDocs > 0 And m_bIndexModifie And Not m_bModeDirect Then
            iReponse = MsgBox("Voulez-vous sauvegarder l'index de VBTextFinder :" & vbLf &
                m_sCheminVBTxtFndIdx & " ?" & vbLf &
                "(nombre de documents : " & lNbDocs & ")",
                MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, sMsgGestionIndex)
            If iReponse = MsgBoxResult.Cancel Then Return False
            If iReponse = MsgBoxResult.No Then GoTo Fin
            If m_bIndexModifie Then
                ' Si la liste des fichiers ini a été modifiée, resauver l'index
                If Not bSauvegarderIndex(m_sCheminVBTxtFndIdx) Then Return False
            Else
                ' Sinon valider la copie temporaire en copie de sauvegarde pour de bon
                '  (ou sinon sauvegarder simplement l'index
                '   si l'option bSauvegardeSecurite = False)
                If Not bValiderSauvegardeTmp() Then Return False
            End If
        End If

Fin:
        Sablier()
        AfficherMessage("Désallocation des mots en mémoire vive...")
        Me.m_htMots = Nothing
        GC.Collect()
        AfficherMessage("Désallocation des phrases en mémoire vive...")
        Me.m_colPhrases = Nothing
        'Me.m_colChapitres = Nothing
        Me.m_colDocs = Nothing
        GC.Collect()

        If Me.m_alsCheminsFichierTxt.Count > 0 AndAlso
            MsgBoxResult.Yes = MsgBox(
                "Voulez-vous supprimer le(s) fichier(s) texte (.txt) temporaire(s) ?",
                MsgBoxStyle.YesNo Or MsgBoxStyle.Question, sTitreMsg) Then
            Dim sCheminFichier$
            For Each sCheminFichier In Me.m_alsCheminsFichierTxt
                bSupprimerFichier(sCheminFichier)
            Next
        End If

        Sablier(bDesactiver:=True)
        bQuitter = True

    End Function

    Public Sub Interrompre()

        ' VB est un langage événementiel multi-thread : deux fonctions peuvent très bien
        '  fonctionner simultanément, on se sert de cela pour pouvoir interrompre une
        '  opération en cours assez longue (il peut même arriver qu'une même fonction
        '  en cours soit ré-appelée : ré-entrance)
        m_bInterrompre = True

    End Sub

    Private Function bInterruption() As Boolean

        ' Laisser du temps pour le traitement des messages : affichage du message et
        '  traitement du clic éventuel sur le bouton Interrompre
        Application.DoEvents()
        bInterruption = m_bInterrompre

    End Function

    Private Sub AfficherMessage(sMsg$)

        Me.m_msgDelegue.AfficherMsg(sMsg)
        ' Rétablir le curseur courant au cas où l'affichage l'aurait fait perdre
        Sablier(Me.m_bSablierDesactive)

    End Sub

    Public Sub Sablier(Optional bDesactiver As Boolean = False)

        Me.m_bSablierDesactive = bDesactiver
        Me.m_msgDelegue.Sablier(bDesactiver)

    End Sub

#End Region

#Region "Indexation"

    Public Function bConvertirDocEnTxt(ByRef sCheminFichierSelect$,
        bVerifierUnicode As Boolean,
        ByRef bTxtUnicode As Boolean,
        ByRef bAvertUnicode As Boolean,
        ByRef bInfoTxtNonUnicode As Boolean,
        bSablier As Boolean) As Boolean

        ' 23/11/2018
        bAvertUnicode = False
        bInfoTxtNonUnicode = False
        bTxtUnicode = False

        Dim sExtension$ = "", sNomFichier$ = ""
        Dim sChemin$, sCheminFichierTxt$
        sChemin = sExtraireChemin(sCheminFichierSelect, sNomFichier, sExtension)
        sExtension = sExtension.ToLower
        ' On laisse le fichier inchangé si on ne peut pas le convertir avec Word
        If Not (sExtension = sExtDoc OrElse sExtension.StartsWith(sExtHtm)) Then _
            bConvertirDocEnTxt = True : Exit Function

        sCheminFichierTxt = sChemin & "\" &
            Left(sNomFichier, Len(sNomFichier) - Len(sExtension)) & sExtTxt

        ' Si le fichier n'existait pas avant, l'ajouter à la liste des fichiers
        '  à supprimer en quittant
        If Not bFichierExiste(sCheminFichierTxt) Then _
            Me.m_alsCheminsFichierTxt.Add(sCheminFichierTxt)

        ' Convertir un fichier .doc ou .html en .txt
        If bSablier Then Sablier()
        bConvertirDocEnTxt = bConvertirDocEnTxt2(sCheminFichierSelect,
            sCheminFichierTxt, m_sCheminDossierCourant, Me.m_msgDelegue,
            m_bOptionTexteUnicode, bVerifierUnicode, bTxtUnicode, bAvertUnicode, bInfoTxtNonUnicode)
        If bConvertirDocEnTxt Then
            AfficherMessage("Conversion en .txt terminée.")
            sCheminFichierSelect = sCheminFichierTxt
            ' 23/11/2018
            If bTxtUnicode Then m_bAuMoinsUnTxtUnicode = True
            If bAvertUnicode Then m_bAvertAuMoinsUnTxtUnicode = True
            If bInfoTxtNonUnicode Then m_bInfoAuMoinsUnTxtNonUnicode = True
        End If
        If bSablier Then Sablier(bDesactiver:=True)

    End Function

    Public Function bIndexerDocuments(sCheminFichier$, bVerifierUnicode As Boolean) As Boolean

        ' Indexer un ou plusieurs documents

        bIndexerDocuments = False
        Sablier()

        Dim dTpsDeb As DateTime = Now

        If sCheminFichier.IndexOfAny("*?".ToCharArray()) > -1 Then
            m_bModeDirect = False
            Dim sRep$ = IO.Path.GetDirectoryName(sCheminFichier)
            Dim sFiltre$ = IO.Path.GetFileName(sCheminFichier)
            Dim aFichiers$() = IO.Directory.GetFiles(sRep, sFiltre)
            Dim i%, sFichier$, iNbFichers%
            iNbFichers = aFichiers.GetUpperBound(0)
            For i = 0 To iNbFichers
                sFichier = aFichiers(i)
                ' Convertir le fichier en .txt si son extension
                '  est celle d'un document convertible (.doc, .html ou .htm)
                ' Le fichier peut être supprimé entre temps
                '  et ne pas ré-indexer les index
                Dim sNomFichier$ = IO.Path.GetFileName(sFichier)
                If sNomFichier.StartsWith(sPrefixeIndex) Then Continue For
                If bFichierExiste(sFichier) AndAlso
                    Left$(sNomFichier, Len(sFichierVBTxtFnd)).ToLower <>
                    sFichierVBTxtFnd.ToLower Then
                    Dim bFichierTxtInexistant As Boolean = False
                    Dim bAvertUnicode As Boolean = False
                    Dim bInfoTxtNonUnicode As Boolean = False
                    Dim bTxtUnicode As Boolean = False
                    If Not bConvertirDocEnTxt(sFichier,
                        bVerifierUnicode, bTxtUnicode, bAvertUnicode, bInfoTxtNonUnicode,
                        bSablier:=False) Then GoTo Fin
                    'If bTxtUnicode Then
                    '    Debug.WriteLine(sNomFichier)
                    'End If
                    If bTxtUnicode Then m_bAuMoinsUnTxtUnicode = True
                    If bAvertUnicode Then m_bAvertAuMoinsUnTxtUnicode = True
                    If bInfoTxtNonUnicode Then m_bInfoAuMoinsUnTxtNonUnicode = True
                    Dim sNumFichier$ = "Doc n°" & i + 1 & " / " & iNbFichers + 1 & " : "
                    ' Le document peut être déjà indexé
                    bIndexerDocument(sFichier, bTxtUnicode, sNumFichier)
                    If m_bInterrompre Then GoTo Fin
                    If m_msgDelegue.m_bAnnuler Then GoTo Fin ' Pas tjrs très réactif ?
                End If
            Next i
        Else
            ' Ici l'info. n'est pas connue, on le sait lorsqu'on conv. le doc en txt
            Const bTxtUnicode As Boolean = False
            If Not bIndexerDocument(sCheminFichier, bTxtUnicode) Then GoTo Fin
        End If

        Dim dTpsFin As DateTime = Now
        Me.tsDiffTps = dTpsFin.Subtract(dTpsDeb)

        If m_bModeDirect Then GoTo FinOk

        EcrireListeDocumentsIndexesIni(bAfficherIni:=False)
        ' Faire une sauvegarde de l'index dans le fichier VBTxtFnd.tmp 
        '  si l'option est activée
        If bSauvegardeSecurite Then bSauvegarderIndex(m_sCheminVBTxtFndTmp)
        AfficherFichierIni()

FinOk:
        bIndexerDocuments = True
        AfficherMessage(sMsgOperationTerminee)

Fin:
        Sablier(bDesactiver:=True)
        If m_bInterrompre Then AfficherMessage("Indexation interrompue.")

    End Function

    Private Function bIndexerDocument(sCheminFichier$, bTxtUnicode As Boolean,
        Optional sNumFichier$ = "") As Boolean

        ' Indexer un document 

        If Not bFichierExiste(sCheminFichier) Then Return False

        m_bInterrompre = False

        ' Générer un code document par défaut
        Dim sCleDoc$ = sCleDocDefaut()

        ' Ajouter le document dans la collection
        If Not bAjouterDocument(sCleDoc, sCleDoc, sCheminFichier, bTxtUnicode) Then Return False

        ' Voir s'il y a code document dans la liste éditable dans le fichier ini
        LireListeDocumentsIndexesIni()

        Dim oDoc As clsDoc = Nothing
        If m_bIndexerChapitre Then
            If m_colDocs.Contains(sCleDoc) Then
                oDoc = DirectCast(m_colDocs.Item(sCleDoc), clsDoc)
                Dim sUnicode$ = ""
                If oDoc.bTxtUnicode Then sUnicode = ":Unicode" ' 24/05/2019
                m_sbChapitres.AppendLine(vbCrLf & oDoc.sChemin & " (" & oDoc.sCodeDoc & sUnicode & ") :")
            End If
        End If

        m_bIndexModifie = True ' Modification de l'index courant
        m_sMemExpression = "" ' La précédente recherche d'expression doit être refaite

        If Not bIndexerDocumentInterne(sCheminFichier, sNumFichier, sCleDoc,
            m_bIndexerChapitre, oDoc) Then Return False

        bIndexerDocument = True

    End Function

    Private Function bIndexerDocumentInterne(sCheminFichier$, sNumFichier$, sCleDoc$,
        Optional bIndexerChapitre As Boolean = False,
        Optional oDoc As clsDoc = Nothing) As Boolean

        Dim sCodeChapitre$ = ""
        Dim iMaxTypeChapitrage% = 0
        Dim asTypesChapitrages$() = Nothing
        Dim iMaxTypeChapitrageXL% = 0
        Dim asTypesChapitragesXL$() = Nothing
        Dim iMaxTypeChapitrageMdb% = 0
        Dim asTypesChapitragesMdb$() = Nothing
        Dim iNumChapitre% = 0
        Dim bTypeChapExclusif As Boolean = False
        If bIndexerChapitre Then
            ParserChapitrage(
                asTypesChapitrages, iMaxTypeChapitrage,
                asTypesChapitragesXL, iMaxTypeChapitrageXL,
                asTypesChapitragesMdb, iMaxTypeChapitrageMdb)
        End If

        Dim sMot$, sLigne$
        Dim bNouvParag As Boolean
        Dim iNbLignes%

        Dim acSepPhrase() As Char = Me.m_sListeSeparateursPhrase.ToCharArray
        Dim acSepMot() As Char = Me.m_sListeSeparateursMot.ToCharArray
        Dim sCle$, sPhrasePonct$
        Dim oPhrase As clsPhrase
        Dim oMot As clsMot
        Dim bCleExiste As Boolean, bPremPhrase As Boolean
        Dim asPhrases$()
        Dim iPosDebPhrase%, iFinLigne%, iNumPhrase%, j%, iNbPhrases%
        Dim iPosDebPhraseSuiv%

        Dim iNbMotsL%, iNbParagL%, iNbPhrasesL%, iDebRech%
        iNbMotsL = 0 : iNbParagL = 0 : iNbPhrasesL = 0

        AfficherMessage("Lecture du document en cours... " & sDossierParent(sCheminFichier))

        Dim asLignes() = asLireFichier(sCheminFichier)

        Dim iNbLignesTot = asLignes.Length
        For Each sLigne In asLignes

            ' Afficher la progression de la lecture
            iNbLignes += 1
            If (iNbLignes Mod iModuloAvanvement) = 0 Then
                AfficherMessage(sNumFichier & "Indexation en cours... " &
                    Int(100.0! * iNbLignes / iNbLignesTot) & "%")
                If m_bInterrompre Then Exit For
            End If

            bNouvParag = True
            If sLigne.Length = 0 Then GoTo LigneSuivante

            asPhrases = sLigne.Split(acSepPhrase)

            iNbPhrases = asPhrases.GetLength(0)
            iFinLigne = sLigne.Length
            iDebRech = 1
            bPremPhrase = False

            For iNumPhrase = 0 To iNbPhrases - 1

                Dim sPhrase$ = asPhrases(iNumPhrase).TrimStart
                If sPhrase.Length = 0 Then GoTo PhraseSuivante

                ' Cas d'une phrase composée seulement de guillemets
                iPosDebPhrase = InStr(iDebRech, sLigne, sPhrase)
                If iPosDebPhrase = 0 Then GoTo PhraseSuivante

                ' Ne pas compter les paragraphes vides
                If bNouvParag Then
                    bNouvParag = False
                    Me.iNbParagG += 1
                    iNbParagL += 1
                End If

                ' Recherche de la phrase avec sa ponctuation
                iPosDebPhraseSuiv = iFinLigne + 1
                If iNumPhrase < iNbPhrases - 1 Then
                    For j = iNumPhrase + 1 To iNbPhrases - 1
                        Dim sPhraseSuiv$ = asPhrases(j).TrimStart
                        If sPhraseSuiv.Length = 0 Then GoTo PhraseSuivante0
                        ' 19/09/2009 Le début de la phrase suivante doit au moins
                        '  etre supérieur à la longueur de la phrase précédante
                        'Dim iPosDebPhraseSuiv_old = InStr(iDebRech + 1, sLigne, sPhraseSuiv)
                        Dim iLenPreced% = asPhrases(j - 1).Length
                        iPosDebPhraseSuiv = InStr(iDebRech + iLenPreced, sLigne, sPhraseSuiv)
                        If iPosDebPhraseSuiv > iPosDebPhrase Then Exit For
PhraseSuivante0:
                    Next j
                    If iPosDebPhraseSuiv <= iPosDebPhrase Then iPosDebPhraseSuiv = iFinLigne + 1
                End If
                ' Tant que l'on n'a pas la première phrase, commencer au début
                If Not bPremPhrase Then iPosDebPhrase = 1
                sPhrasePonct = sLigne.Substring(iPosDebPhrase - 1, iPosDebPhraseSuiv - iPosDebPhrase)
                ' Supprimer l'espace à gauche, car il est présent en double via le split
                '  (sauf la première phrase)
                If bPremPhrase Then sPhrasePonct = sPhrasePonct.TrimStart
                bPremPhrase = True
                iDebRech = iPosDebPhraseSuiv

                ' Ajouter une phrase à la liste des phrases indexées

                Me.iNbPhrasesG += 1 ' Nombre de phrases globales
                iNbPhrasesL += 1 ' Nombre de phrases du document

                ' 19/06/2010 Analyse du chapitrage
                If bIndexerChapitre AndAlso iNumPhrase = 0 Then
                    GestionChapitrage(sLigne, sPhrasePonct, sCleDoc, oDoc,
                        iMaxTypeChapitrage, asTypesChapitrages,
                        iNumChapitre, sCodeChapitre, bTypeChapExclusif,
                        iMaxTypeChapitrageXL, asTypesChapitragesXL,
                        iMaxTypeChapitrageMdb, asTypesChapitragesMdb)
                End If

                ' iNbPhrasesG : Numéro de la phrase
                oPhrase = New clsPhrase With {
                    .iNumPhraseG = Me.iNbPhrasesG,
                    .iNumPhraseL = iNbPhrasesL,
                    .sClePhrase = Me.iNbPhrasesG.ToString,
                    .iNumParagrapheL = iNbParagL,
                    .iNumParagrapheG = iNbParagG,
                    .sCleDoc = sCleDoc,
                    .sCodeChapitre = sCodeChapitre
                }

                ' 02/08/2010 Remplacer les espaces insécables pour faciliter les recherches
                If bSupprimerEspInsec Then
                    oPhrase.sPhrase = sPhrasePonct.Replace(Chr(iCodeASCIIEspaceInsecable), " "c)
                    ' 15/09/2018
                    If Me.m_bOptionTexteUnicode Then
                        oPhrase.sPhrase = oPhrase.sPhrase.Replace(
                            ChrW(iCodeUTF16EspaceFineInsecable), " "c)
                        oPhrase.sPhrase = oPhrase.sPhrase.Replace(
                            ChrW(iCodeUTF16EspaceInsecable), " "c) ' 13/07/2019
                    End If
                Else
                    oPhrase.sPhrase = sPhrasePonct
                End If
                m_colPhrases.Add(oPhrase) ' ArrayList

                Dim asMots$() = asPhrases(iNumPhrase).Split(acSepMot)
                ' 20/11/2016 Le découpage fonctionne très bien, on ne trouve jamais aucune différence
                'Dim asMots2$() = asPhrases(iNumPhrase).VBSplit(acSepMot)
                'Dim bEgal = Linq.Enumerable.SequenceEqual(asMots, asMots2)
                'If Not bEgal Then
                '    MsgBox("Différence trouvée : [" & asPhrases(iNumPhrase) & "]")
                '    Return False
                'End If

                For Each sMot In asMots

                    If sMot.Length = 0 Then GoTo MotSuivant

                    Me.iNbMotsG += 1
                    iNbMotsL += 1

                    ' Indexer les mots pour ne conserver que les mots distincts

                    If bSupprimerNumeriquesEnFinDeMot Then
                        sMot = sSupprimerNumeriquesEnFinDeMot(sMot)
                    End If

                    ' D'abord vérifier rapidement si le mot est indexé tel quel
                    ' Les mots accentués sont distingués
                    sCle = sMot
                    bCleExiste = Me.m_htMots.ContainsKey(sCle)

                    ' S'il n'est pas indexé tel quel, vérifier s'il est indexé
                    '  sans les accents si c'est l'option choisie
                    '  (si le mot a été trouvé tel quel, c'est qu'il n'avait pas d'accent)
                    If Not m_bIndexerAccents And Not bCleExiste Then
                        ' Les mots accentués ne sont pas distingués
                        sCle = sEnleverAccents(sMot)
                        If sCle <> sMot Then bCleExiste = Me.m_htMots.ContainsKey(sCle)
                    End If

                    If bCleExiste Then
                        ' DirectCast = Casting direct comme CType mais sans conversion
                        oMot = DirectCast(Me.m_htMots.Item(sCle), clsMot)
                        ' Mot déjà existant : incrémenter le nombre d'occurrences
                        With oMot
                            .iNbOccurrences += 1
                            .aiNumPhrase.Add(Me.iNbPhrasesG)
                        End With

                    Else

                        ' Clé absente dans la collection, on ajoute le mot
                        oMot = New clsMot
                        With oMot
                            ' On peut laisser les accents ici, contrairement à la clé
                            .sMot = sMot.ToLower
                            .iNbOccurrences = 1
                            .aiNumPhrase.Add(Me.iNbPhrasesG)
                        End With
                        Me.m_htMots.Add(sCle, oMot) ' Ajout du mot dans la Hastable

                    End If

MotSuivant:
                Next sMot

PhraseSuivante:
            Next iNumPhrase

LigneSuivante:
        Next sLigne

        Return True

    End Function

    Private Sub ParserChapitrage(
        ByRef asTypesChapitrages$(), ByRef iMaxTypeChapitrage%,
        ByRef asTypesChapitragesXL$(), ByRef iMaxTypeChapitrageXL%,
        ByRef asTypesChapitragesMdb$(), ByRef iMaxTypeChapitrageMdb%)

        asTypesChapitrages = m_sChapitrage.Split(";"c)
        Dim iMax% = asTypesChapitrages.GetUpperBound(0) ' Types + Codes
        Dim iNbChap% = (iMax + 1) \ 2 ' Types seuls
        iMaxTypeChapitrage = iNbChap - 1

        asTypesChapitragesXL = m_sChapitrageXL.Split(";"c)
        Dim iMaxXL% = asTypesChapitragesXL.GetUpperBound(0)
        Dim iNbChapXL% = (iMaxXL + 1) \ 2
        iMaxTypeChapitrageXL = iNbChapXL - 1

        asTypesChapitragesMdb = m_sChapitrageMdb.Split(";"c)
        Dim iMaxMdb% = asTypesChapitragesMdb.GetUpperBound(0)
        Dim iNbChapMdb% = (iMaxMdb + 1) \ 2
        iMaxTypeChapitrageMdb = iNbChapMdb - 1

        ' 01/05/2012 Arrondir à pair, s'il manque le couple Chapitre-Code chapitre
        iMax = iNbChap * 2 - 1
        iMaxXL = iNbChapXL * 2 - 1
        iMaxMdb = iNbChapMdb * 2 - 1

        Dim iSupplement% = iNbChapXL + iNbChapMdb
        If iSupplement > 0 Then
            ' Déplacer le chapitrage normal à la fin (l'exclusif est prioritaire)
            Dim iMemMax% = iMax
            Dim iMemNbChap% = iNbChap
            iNbChap += iSupplement
            iMaxTypeChapitrage = iNbChap - 1
            iMax = iNbChap * 2 - 1
            Dim iDec% = (iNbChapXL + iNbChapMdb) * 2
            ReDim Preserve asTypesChapitrages(0 To iMax)
            For i = iMemMax To 0 Step -1
                asTypesChapitrages(iDec + i) = asTypesChapitrages(i)
            Next
            ' Copier le chapitrage Excel au début
            For i = 0 To iMaxXL
                asTypesChapitrages(i) = asTypesChapitragesXL(i)
            Next
            ' Ensuite copier le chapitrage Access à la suite
            For i = iMaxXL + 1 To iMaxXL + 1 + iMaxMdb
                asTypesChapitrages(i) = asTypesChapitragesMdb(i - (iMaxXL + 1))
            Next
        End If

    End Sub

    Private Sub GestionChapitrage(sLigne$, sPhrasePonct$,
        sCleDoc$, oDoc As clsDoc,
        iMaxTypeChapitrage%, asTypesChapitrages$(),
        ByRef iNumChapitre%, ByRef sCodeChapitre$,
        ByRef bTypeChapExclusif As Boolean,
        iMaxTypeChapitrageXL%, asTypesChapitragesXL$(),
        iMaxTypeChapitrageMdb%, asTypesChapitragesMdb$())

        ' Gestion du chapitrage (le fait de noter la position d'un mot dans 
        '  un chapitre précis d'un document indexé)

        ' Si la ligne contient des tabulations ou plusieurs espaces consécutifs :
        '  table des matières : ignorer
        Dim iPosTab% = sLigne.IndexOf(vbTab)
        Dim iPos2Esp% = sLigne.IndexOf("  ")
        Dim bTableDesMatieres As Boolean = False
        If iPosTab <> -1 Or iPos2Esp <> -1 Then bTableDesMatieres = True
        If bTableDesMatieres Then Exit Sub

        Dim cEspaceInsec As Char = Chr(iCodeASCIIEspaceInsecable)

        For i = 0 To iMaxTypeChapitrage

            Dim sTypeChap$ = asTypesChapitrages(i * 2)
            Dim sCodeChap$ = asTypesChapitrages(i * 2 + 1)
            If Not sPhrasePonct.StartsWith(sTypeChap, m_styleCompare2) Then Continue For

            ' Détection d'un type de chapitre à ignorer
            If sCodeChap.StartsWith("-") Then Exit For

            Dim sReste$ = sPhrasePonct.Substring(sTypeChap.Length)

            Dim bTypeChapExclusifMaintenant As Boolean = False
            For j = 0 To iMaxTypeChapitrageXL
                If sTypeChap = asTypesChapitragesXL(j * 2) Then
                    bTypeChapExclusifMaintenant = True
                    Exit For
                End If
            Next
            If Not bTypeChapExclusifMaintenant Then
                For j = 0 To iMaxTypeChapitrageMdb
                    If sTypeChap = asTypesChapitragesMdb(j * 2) Then
                        bTypeChapExclusifMaintenant = True
                        Exit For
                    End If
                Next
            End If
            If bTypeChapExclusifMaintenant Then
                bTypeChapExclusif = True
            Else
                ' Ne pas chercher les autres types de chapitres dans ce cas
                ' (éviter de détecter des chapitres intempestifs dans 
                '  le contenu Access ou Excel)
                If bTypeChapExclusif Then Exit For
            End If

            ' Il faut qualifier un type de chapitre : n°, ...
            If sReste.Length = 0 Then Exit For

            ' Autoriser le numChapitre à coller au chapitre, 
            '  à condition qu'il soit numérique
            ' Commencer par vérifier si le numChapitre est collé
            Dim iPosEspace% = sReste.IndexOf(" ")
            Dim iPosEspaceInsec% = sReste.IndexOf(cEspaceInsec)
            Dim bNumChapColle As Boolean = True
            If (iPosEspace = 0 Or iPosEspaceInsec = 0) Then bNumChapColle = False

            If bNumChapColle Then
                ' Ensuite vérifier si le 1er car. qui suit est numérique
                Dim cCar1 As Char = sReste.Chars(0)
                Dim b1erCarNum As Boolean = bCarNumerique(cCar1)
                ' Si le 1er car. n'est pas numérique : refuser (partie de mot)
                If Not b1erCarNum Then Continue For
            End If

            Dim sNumSection$ = ""
            If iPosEspace = 0 Then
            ElseIf iPosEspaceInsec = 0 Then
            ElseIf iPosEspace > 0 Then '-1 Then
                sNumSection = sReste.Substring(0, iPosEspace)
            ElseIf iPosEspaceInsec > 0 Then '-1 Then
                sNumSection = sReste.Substring(0, iPosEspaceInsec)
            Else
                sNumSection = sReste
            End If

            ' Créer une numérotation automatique si pas d'autre solution
            Dim bNumAuto As Boolean = False

            If sNumSection.Trim.Length = 0 Then
                Dim sReste2$ = sRognerDernierCar(sReste.Trim, ":")
                ' Vérifier si la longueur totale de la numérotation 
                '  n'est pas trop grande (sinon on garde)
                sNumSection = sReste2.Trim
                If sNumSection.Length > iNbCarChapitreMax Then bNumAuto = True
            End If

            If bNumAuto Then
                ' Créer une numérotation automatique
                iNumChapitre += 1
                sNumSection = iNumChapitre.ToString
            Else
                sNumSection = sRognerDernierCar(sNumSection, ".")
                sNumSection = sRognerDernierCar(sNumSection, ":")
            End If

            sCodeChapitre = sCodeChap & sNumSection.Trim

            Dim chap As New clsChapitre
            chap.sCodeChapitre = sCodeChapitre
            chap.sCleDoc = sCleDoc
            chap.sCodeDoc = sCleDoc ' Pas encore éditée pour le moment
            chap.sChapitre = sLigne
            chap.sCle = sCleDoc & ":" & chap.sCodeChapitre
            If oDoc.colChapitres.Contains(chap.sCle) Then
                ' Si la clé existe déjà, alors ignorer le chapitrage
                ' (c'est sans doute une occurrence parasite dans le texte)
                Exit For
            End If
            oDoc.colChapitres.Add(chap, chap.sCle)
            m_sbChapitres.Append(sCodeChapitre)
            m_sbChapitres.AppendLine(" : " & sLigne)
            Exit For

        Next

    End Sub

    Public Function sCleDocDefaut$()

        ' Générer un code document par défaut
        sCleDocDefaut = sCodeDocDef & Me.m_colDocs.Count() + 1

    End Function

    Public Function bCleDocExiste(sCleDoc$) As Boolean

        ' Vérifier si un code document est déjà utilisé pour un des documents indexés

        ' On peut laisser un code document vide : un code numéroté sera généré par défaut
        If sCleDoc = "" Then Return False
        If Me.m_colDocs.Count = 0 Then Return False
        'bCleDocExiste = m_colDocs.ContainsKey(sCleDoc) ' Hastable
        Return m_colDocs.Contains(sCleDoc) ' Collection : Contains a été ajouté depuis VB6 !

    End Function

    Private Function sLireCleDocPhrase$(iNumPhraseG%,
                                        Optional ByRef sCodeChapitre$ = "")

        ' Retourner le code document d'un numéro de phrase global

        Dim oPhrase As clsPhrase
        If iNumPhraseG > m_colPhrases.Count Then ' 10/12/2022
            If bDebug Then Stop
            Return ""
        End If
        oPhrase = DirectCast(m_colPhrases.Item(iNumPhraseG - 1), clsPhrase)
        sLireCleDocPhrase = oPhrase.sCleDoc
        sCodeChapitre = oPhrase.sCodeChapitre

    End Function

    Private Function sLireCodeDoc$(sCleDoc$)

        ' Retourner le code document de la clé d'un document
        '  util pour trouver le code document via une clé de document
        '  associée à une phrase

        Dim oDoc As clsDoc
        oDoc = DirectCast(m_colDocs.Item(sCleDoc), clsDoc)
        sLireCodeDoc = oDoc.sCodeDoc

    End Function

    Public Function iNbDocumentsIndexes%()

        iNbDocumentsIndexes = Me.m_colDocs.Count()

    End Function

    Public Sub ListerDocumentsIndexes(ByRef CtrlResultat As Windows.Forms.TextBox,
        Optional bListerPhrases As Boolean = True,
        Optional bHtml As Boolean = False)

        ' Afficher la liste des documents indexés

        AfficherMessage("Ecriture du rapport d'indexation...")

        Dim sbResultat As New StringBuilder
        ' Utiliser le format de présentation en français, 
        '  en utilisant les préférences de l'utilisateur le cas échéant
        Dim nfi As System.Globalization.NumberFormatInfo =
            New System.Globalization.CultureInfo("fr-FR", useUserOverride:=True).NumberFormat
        nfi.NumberDecimalDigits = 0 ' Afficher des nombres entiers, sans virgule

        If m_bAvertAuMoinsUnTxtUnicode Then ' 23/11/2018
            ' Au moins un document ici :
            sbResultat.AppendLine("INFORMATION : Le texte contient des caractères Unicode et l'option n'est pas activée (ces caractères seront remplacés par des signes '?').")
            sbResultat.AppendLine("")
        End If
        If m_bInfoAuMoinsUnTxtNonUnicode Then
            'sbResultat.AppendLine("Information : Le texte ne contient pas de caractères Unicode.")
            sbResultat.AppendLine("Information : Au moins un document ne contient pas de caractères Unicode.")
            sbResultat.AppendLine("")
        End If

        sbResultat.Append("Nombre de mots indexés : " &
            Me.iNbMotsG.ToString("N", nfi) & vbCrLf)
        sbResultat.Append("Nombre de mots distincts indexés : " &
            Me.m_htMots.Count().ToString("N", nfi) & vbCrLf)
        sbResultat.Append("Nombre de phrases indexées : " &
            m_colPhrases.Count().ToString("N", nfi) & vbCrLf)
        sbResultat.Append("Nombre de paragraphes indexés : " &
            Me.iNbParagG.ToString("N", nfi) & vbCrLf)

        If Me.tsDiffTps.Milliseconds <> 0 Then _
            sbResultat.Append("Temps d'indexation : " & tsDiffTps.ToString & vbCrLf)

        sbResultat.Append(vbCrLf)
        sbResultat.Append("Liste des documents indexés (" & Me.m_colDocs.Count() & ") :" & vbCrLf)
        'Dim de As DictionaryEntry
        'For Each de In m_colDocs
        '    Dim oDoc As clsDoc = DirectCast(de.Value, clsDoc)
        '    sbResultat.Append(oDoc.sChemin & " (" & oDoc.sCle & ")" & vbCrLf
        'Next de
        Dim oDoc As clsDoc
        For Each oDoc In Me.m_colDocs
            Dim sUnicode$ = ""
            If oDoc.bTxtUnicode Then sUnicode = ":Unicode" ' 24/05/2019
            sbResultat.AppendLine(oDoc.sChemin & " (" & oDoc.sCodeDoc & sUnicode & ")")
        Next oDoc

        If m_bIndexerChapitre Then
            sbResultat.AppendLine(vbCrLf & "Liste des chapitres :")
            sbResultat.Append(m_sbChapitres)
            ' Identique à m_sbChapitres :
            'For Each oDoc In Me.m_colDocs
            '    sbResultat.AppendLine(vbCrLf & oDoc.sChemin & " (" & oDoc.sCodeDoc & ") :")
            '    For Each chapitre As clsChapitre In oDoc.colChapitres
            '        sbResultat.AppendLine(chapitre.sCodeChapitre & " : " & chapitre.sChapitre)
            '    Next chapitre
            'Next oDoc
        End If

        If Not bHtml Then CtrlResultat.Text = sbResultat.ToString

        If Not bListerPhrases Then GoTo Suite

        sbResultat.Append(vbCrLf & "Liste des phrases :")

        Dim i%, iMemParag%, sMemCleDoc$
        Dim oPhrase As clsPhrase
        sMemCleDoc = ""
        For i = 1 To Me.iNbPhrasesG ' Parcours de toutes les phrases
            oPhrase = DirectCast(m_colPhrases.Item(i - 1), clsPhrase)
            If oPhrase.sCleDoc <> sMemCleDoc Then
                sbResultat.Append(vbCrLf & vbCrLf & "Document : " &
                    DirectCast(Me.m_colDocs(oPhrase.sCleDoc), clsDoc).sChemin &
                    " (" & sLireCodeDoc(oPhrase.sCleDoc) & ")" & vbCrLf)
                sbResultat.Append(vbCrLf)
            ElseIf oPhrase.iNumParagrapheL <> iMemParag Then
                sbResultat.Append(vbCrLf)
            End If
            sbResultat.Append(oPhrase.sPhrase)
            iMemParag = oPhrase.iNumParagrapheL
            sMemCleDoc = oPhrase.sCleDoc

            If bHtml Then Continue For

            If sbResultat.Length > iMaxLongChaine0 Then Exit For

            If i Mod iModuloAvanvementRapide = 0 Then
                CtrlResultat.Text = sbResultat.ToString
            End If

        Next i

Suite:
        If bHtml Then
            ' On duplique ici car on va modifier le sb pour le html
            m_sbResultatTxt = New StringBuilder
            m_sbResultatTxt.Append(sbResultat)
            m_sbResultatHtml = sbResultat.Replace(vbLf, "<br>")
            Exit Sub
        End If

        ' Afficher le résultat final si ce n'est pas déjà fait
        If CtrlResultat.ToString() <> sbResultat.ToString And
           sbResultat.Length <= iMaxLongChaine0 Then _
            CtrlResultat.Text = sbResultat.ToString

        If sbResultat.Length > iMaxLongChaine0 Then CtrlResultat.Text &= "..."

    End Sub

    Private Function bAjouterDocument(sCleDoc$, sCodeDoc$,
        ByRef sCheminFichier$, bTxtUnicode As Boolean,
        Optional colChapitres As Collection = Nothing) As Boolean

        ' Ajouter un document à la liste des documents indexés

        bAjouterDocument = False
        ' Stocker les chemins en relatif le cas échéant
        Dim sCheminAIndexer$
        Dim sFichier$ = ""
        Dim sChemin$ = sExtraireChemin(sCheminFichier, sFichier)
        If sChemin.ToLower = m_sCheminDossierCourant.ToLower Then
            sCheminAIndexer = sFichier
        Else
            sCheminAIndexer = sCheminFichier
        End If

        Dim oDoc As clsDoc
        ' Vérifier si le document est déjà indexé
        For Each oDoc In Me.m_colDocs
            If oDoc.sChemin.ToLower = sCheminAIndexer.ToLower Then Return False
        Next oDoc

        oDoc = New clsDoc
        oDoc.sChemin = sCheminAIndexer
        oDoc.sCle = sCleDoc
        oDoc.sCodeDoc = sCodeDoc
        oDoc.bTxtUnicode = bTxtUnicode ' 26/01/2019

        If Not IsNothing(colChapitres) Then
            oDoc.colChapitres = colChapitres
        End If

        'm_colDocs.Add(oDoc, oDoc.sCle)
        Const sMsgModeMultiDoc$ = "Indexation des documents"
        'If m_colDocs.ContainsKey(sCodeDoc) Then
        If bCleDocExiste(sCleDoc) Then
            ' Pertinent dans la version VB6, dans la version VB7 on ne peut pas le changer
            '  c'est dans la gestion du fichier ini que l'on vérifie l'unicité de la clé
            '  avec la hastable
            MsgBox("La clé '" & sCleDoc & "' a déjà été utilisée",
                MsgBoxStyle.Critical, sMsgModeMultiDoc)
            GoTo Fin
        End If
        Try
            'm_colDocs.Add(oDoc.sCle, oDoc) ' Hastable
            Me.m_colDocs.Add(oDoc, sCleDoc)  ' Collection
        Catch Err As Exception ' Erreur managée
            AfficherMsgErreur2(Err, "bAjouterDocument",
                "Impossible d'ajouter le document : " & sCodeDoc & " : " & sCheminAIndexer)
        End Try

        Return True

Fin:

    End Function

    Public Function bMotExiste(sMot$, ByRef oMot As clsMot) As Boolean

        ' Vérifier si un mot est indexé, et retourner le mot le cas échéant

        Dim sCle$
        If m_bIndexerAccents Then
            ' Les mots accentués sont distingués
            sCle = sMot
        Else
            ' Les mots accentués ne sont pas distingués
            sCle = sEnleverAccents(sMot)
        End If
        bMotExiste = Me.bCleExiste(sCle, oMot)
        If bMotExiste Then Exit Function
        ' Si on récupère un index de la version VB6, tester aussi avec les accents
        If m_bIndexerAccents Then Exit Function
        If Not bCompatVB6RechercheAussiAvecAccents Then Exit Function
        If String.Compare(sMot, sCle) = 0 Then Exit Function
        bMotExiste = Me.bCleExiste(sMot, oMot)

    End Function

    Private Function bCleExiste(sCle$, ByRef oMot As clsMot) As Boolean

        ' Vérifier si une clé figure déjà dans l'index, et retourner le mot le cas échéant

        oMot = Nothing
        If sCle.Length = 0 Then Return False
        bCleExiste = Me.m_htMots.ContainsKey(sCle)
        If Not bCleExiste Then Return False
        oMot = DirectCast(Me.m_htMots.Item(sCle), clsMot)

    End Function

#End Region

#Region "Gestion des fichiers ini"

    Public Sub LireListeDocumentsIndexesIni()

        ' Lire le fichier ini des documents indexés pour éditer les codes document

        If m_bModeDirect Then Exit Sub

        If Not bFichierExiste(m_sCheminFichierIni) Then Exit Sub

        ' Autre solution : CreateCaseInsensitiveHashtable
        Dim htCodesDoc As New Hashtable(StringComparer.InvariantCultureIgnoreCase)

        Me.m_colDocsIni = New Collection

        Dim asLignes() = sLireFichier(m_sCheminFichierIni).Split(CChar(vbLf))
        For Each sLigne In asLignes

            ' 24/05/2019 Changement du séparateur : | pour tenir compte de C: éventuellement
            Dim asChamps() = sLigne.Split(CChar(sSepIni)) '"|"c)
            Dim iNbChamps% = asChamps.GetLength(0)
            If iNbChamps = 0 Then Continue For
            Dim sCheminDoc$ = "", sCodeDoc$ = "", bTxtUnicode = False
            If iNbChamps > 0 Then sCheminDoc = asChamps(0).Trim
            If iNbChamps > 1 Then sCodeDoc = asChamps(1).Trim
            If iNbChamps > 2 Then bTxtUnicode = (asChamps(2).Trim = "Unicode")

            'Dim iPos = sLigne.LastIndexOf(":")
            'If iPos <= 0 Then Continue For
            'Dim sCheminDoc$ = "", sCodeDoc$ = "", bTxtUnicode = False
            'Dim sGauche = Left(sLigne, iPos)
            'Dim sDernChamp = Mid(sLigne, iPos + 2).Trim
            '' 26/01/2019 Ajout du champ optionnel bUnicode : pas possible avec .LastIndexOf(":")
            'If sDernChamp = sTagUnicodeIni Then
            '    bTxtUnicode = True
            '    Dim iPos2 = sGauche.LastIndexOf(":")
            '    If iPos2 <= 0 Then Continue For
            '    sCheminDoc = Left(sLigne, iPos2)
            '    sCodeDoc = Mid(sGauche, iPos2 + 2).Trim
            'Else
            '    sCheminDoc = sGauche
            '    sCodeDoc = sDernChamp
            'End If

            ' Vérifier si le document existe
            LireDoc(htCodesDoc, sCheminDoc, sCodeDoc, bTxtUnicode)

            ' Vérifier si le code doc existe déjà
            If htCodesDoc.ContainsKey(sCodeDoc) Then
                'MsgBox("Le code document '" & sCodeDoc & "' existe déjà !", _
                '    MsgBoxStyle.Information, "Lecture des codes document")
                Continue For
            End If

            htCodesDoc.Add(sCodeDoc, sCodeDoc)

            ' La collection VB6 préserve l'ordre
            Dim oDoc As New clsDoc
            oDoc.sCle = sCodeDoc
            oDoc.sCodeDoc = sCodeDoc
            oDoc.sChemin = sCheminDoc
            oDoc.bTxtUnicode = bTxtUnicode ' 26/01/2019
            If bTxtUnicode Then
                If Not m_bOptionTexteUnicode Then m_bAvertAuMoinsUnTxtUnicode = True
            Else
                If m_bOptionTexteUnicode Then m_bInfoAuMoinsUnTxtNonUnicode = True
            End If
            Me.m_colDocsIni.Add(oDoc, sCodeDoc)

        Next sLigne

    End Sub

    Private Sub LireDoc(htCodesDoc As Hashtable, sCheminDoc$, sCodeDoc$, bTxtUnicode As Boolean)

        For Each oDoc As clsDoc In Me.m_colDocs

            If oDoc.sChemin <> sCheminDoc Then Continue For

            ' Vérifier si le code doc existe déjà
            If htCodesDoc.ContainsKey(sCodeDoc) Then
                MsgBox("Le code document '" & sCodeDoc & "' existe déjà !",
                    MsgBoxStyle.Information, "Lecture des codes document")
                Continue For
            End If

            htCodesDoc.Add(sCodeDoc, sCodeDoc)

            ' Mettre à jour le code doc
            If sCodeDoc <> oDoc.sCodeDoc Then m_bIndexModifie = True
            oDoc.sCodeDoc = sCodeDoc

            ' 27/01/2019 Si dans le fichier ini on a Unicode, alors reporter l'info.
            If bTxtUnicode AndAlso Not oDoc.bTxtUnicode Then
                m_bIndexModifie = True
                oDoc.bTxtUnicode = bTxtUnicode
            End If

            Exit Sub

        Next oDoc

    End Sub

    Public Sub AfficherFichierIni()

        ' 29/08/2010 Pas de fichier ini si un seul document
        If Not bFichierExiste(m_sCheminFichierIni) Then Exit Sub
        Shell("notepad.exe " & m_sCheminFichierIni, AppWinStyle.NormalFocus)
        ' Les fichiers ini ne sont pas forcément associés au bloc-notes, à éviter :
        'OuvrirAppliAssociee(m_sCheminFichierIni)

    End Sub

    Private Sub EcrireListeDocumentsIndexesIni(
        Optional bAfficherIni As Boolean = False)

        ' Afficher la liste des documents indexés

        If Not bFichierAccessible(m_sCheminFichierIni,
            bPrompt:=True, bInexistOk:=True) Then Exit Sub

        ' 01/06/2019 Faire un .bak : VBTextFinder.ini -> VBTextFinder-ini.bak
        If bFichierExiste(m_sCheminFichierIni) Then
            Dim sDossierBak$ = sDossierParent(m_sCheminFichierIni)
            Dim sFichierIniBak$ = IO.Path.GetFileNameWithoutExtension(m_sCheminFichierIni) & "-ini.bak"
            Dim sCheminBak$ = sDossierBak & "\" & sFichierIniBak
            If Not bFichierAccessible(sCheminBak, bPrompt:=True, bInexistOk:=True) Then Exit Sub
            If Not bCopierFichier(m_sCheminFichierIni, sCheminBak) Then Exit Sub
        End If

        Dim sb As New StringBuilder
        Dim oDoc As clsDoc
        For Each oDoc In Me.m_colDocs
            Dim sLigne$ = oDoc.sChemin & sSepIni & oDoc.sCodeDoc
            If oDoc.bTxtUnicode Then sLigne &= sSepIni & sTagUnicodeIni ' 26/01/2019
            sb.Append(sLigne).Append(vbCrLf)
        Next oDoc
        For Each oDoc In Me.m_colDocsIni
            Dim sLigne$ = oDoc.sChemin & sSepIni & oDoc.sCodeDoc
            If oDoc.bTxtUnicode Then sLigne &= sSepIni & sTagUnicodeIni ' 26/01/2019
            sb.Append(sLigne).Append(vbCrLf)
        Next oDoc
        If Not bEcrireFichier(m_sCheminFichierIni, sb) Then Exit Sub
        If bAfficherIni Then AfficherFichierIni()

    End Sub

#End Region

#Region "Algorithme de recherche"

    Public Sub ChercherOccurrencesMot(
        ByRef CtrlMot As ComboBox, ByRef CtrlResultat As TextBox, iNbZoomParag%,
        bAfficherInfoResultat As Boolean, bAfficherInfoDoc As Boolean,
        bAfficherNumParag As Boolean, bAfficherNumPhrase As Boolean,
        bAfficherNumOccur As Boolean, bAfficherTiret As Boolean,
        bHtml As Boolean)

        ' Fonction principale du moteur de recherche d'un mot dans l'index

        ' Lors de l'initialisation du logiciel, la zone est vide
        Dim sMot$ = CtrlMot.Text
        If sMot = "" Then Exit Sub

        Static bRechercheEnCours As Boolean
        If bRechercheEnCours Then Exit Sub
        bRechercheEnCours = True

        Dim oMot As clsMot = Nothing
        If Not bMotExiste(sMot, oMot) Then
            AfficherMessage("Mot non trouvé : " & sMot)
            GoTo Fin
        End If

        m_bInterrompre = False

        Dim sbResultat As New StringBuilder
        Dim alResultats As New ArrayList(oMot.aiNumPhrase)
        'Dim iNbPhrasesTrouvees% = alResultats.Count

        Dim sExpressions$ = CtrlMot.Text

        Dim alExpressions As New ArrayList()
        If m_bIndexerAccents Then
            alExpressions.Add(sExpressions.ToLower)
        Else
            ' Enlever les accents et passer en minuscule
            alExpressions.Add(sEnleverAccents(sExpressions))
        End If

        Dim iNbOccurrencesTot% = oMot.iNbPhrases

        'CtrlResultat.SuspendLayout()
        Sablier() ' 01/05/2010

        AfficherResultats(sExpressions, alResultats, iNbZoomParag, bAfficherInfoResultat,
            bAfficherInfoDoc, bAfficherNumParag, bAfficherNumPhrase, bAfficherNumOccur,
            iNbOccurrencesTot, bAfficherTiret, sbResultat, CtrlResultat, alExpressions, bHtml)

        If bAfficherInfoResultat Then _
        RestaurerPositionCurseur(CtrlResultat, sMot,
            m_iNumParagSel, m_iNumPhraseSel, m_iNumCarSel, m_iLongSel, iNbZoomParag, sbResultat)

        'CtrlResultat.ResumeLayout()

        AjouterMotDejaTrouve(sExpressions, CtrlMot)

Fin:
        Sablier(bDesactiver:=True) ' 01/05/2010
        bRechercheEnCours = False

    End Sub

    Public Sub ChercherOccurrencesMots(
        ByRef CtrlMot As ComboBox, ByRef CtrlResultat As TextBox, iNbZoomParag%,
        bAfficherInfoResultat As Boolean, bAfficherInfoDoc As Boolean,
        bAfficherNumParag As Boolean, bAfficherNumPhrase As Boolean,
        bAfficherNumOccur As Boolean, bAfficherTiret As Boolean,
        bHtml As Boolean)

        ' Chercher des occurrences de mots ou expressions complexes entre guillemets

        ' Lors de l'initialisation du logiciel, la zone est vide
        Dim sExpressions$ = CtrlMot.Text
        If sExpressions = "" Then Exit Sub

        Static bRechercheEnCours As Boolean
        If bRechercheEnCours Then Exit Sub
        bRechercheEnCours = True

        Dim alResultats As ArrayList
        Dim alExpressions As ArrayList
        Static alMemResultats As New ArrayList
        Static alMemExpressions As New ArrayList
        If sExpressions = m_sMemExpression AndAlso Not bDebug Then
            alResultats = alMemResultats
            alExpressions = alMemExpressions
            GoTo AfficherResultats
        End If
        m_sMemExpression = sExpressions

        Sablier()

        ' Extraire les expressions délimitées par les guillemets
        Dim asExpressions() = asArgLigneCmd(sExpressions, bSupprimerEspaces:=False)
        alExpressions = New ArrayList() 'asExpressions)
        For Each sExpression As String In asExpressions
            If m_bIndexerAccents Then
                alExpressions.Add(sExpression.ToLower)
            Else
                ' Enlever les accents et passer en minuscule
                alExpressions.Add(sEnleverAccents(sExpression))
            End If
        Next
        alMemExpressions = alExpressions

        Dim oPhrase As clsPhrase
        Dim iNumPhrase As Integer

        alResultats = New ArrayList ' Liste des n° de phrase validée

        ' Rechercher les phrases contenant les expressions demandées
        For Each oPhrase In m_colPhrases
            iNumPhrase = iNumPhrase + 1
            If iNumPhrase Mod iModuloAvanvementTresLent = 0 Or iNumPhrase = Me.iNbPhrasesG Then
                AfficherMessage("Recherche des phrases en cours : " &
                    iNumPhrase & " / " & Me.iNbPhrasesG)
                If m_bInterrompre Then Exit For
            End If

            ' Une phrase validée doit contenir chaque expression demandée
            Dim bOk As Boolean = False
            Dim iNbExpressions% = 0
            Dim sPhrase$ = ""
            If m_bIndexerAccents Then
                sPhrase = oPhrase.sPhrase.ToLower
            Else
                sPhrase = sEnleverAccents(oPhrase.sPhrase)
            End If
            For Each sExpression As String In alExpressions
                If sPhrase.IndexOf(sExpression) = -1 Then Exit For
                iNbExpressions += 1
            Next
            If iNbExpressions < alExpressions.Count Then Continue For

            alResultats.Add(iNumPhrase)

        Next oPhrase
        alMemResultats = alResultats

AfficherResultats:

        Dim sbResultat As New StringBuilder
        'Dim iNbPhrasesTrouvees% = alResultats.Count
        Dim iNbOccurrencesTot% = -1 ' Inconnu
        AfficherResultats(sExpressions, alResultats, iNbZoomParag, bAfficherInfoResultat,
            bAfficherInfoDoc, bAfficherNumParag, bAfficherNumPhrase,
            bAfficherNumOccur, iNbOccurrencesTot, bAfficherTiret,
            sbResultat, CtrlResultat, alExpressions, bHtml)

        If bAfficherInfoResultat Then _
        RestaurerPositionCurseur(CtrlResultat, sExpressions,
            m_iNumParagSel, m_iNumPhraseSel, m_iNumCarSel, m_iLongSel, iNbZoomParag, sbResultat)

        AjouterMotDejaTrouve(sExpressions, CtrlMot)

        Sablier(bDesactiver:=True)
        bRechercheEnCours = False

    End Sub

    Public Sub InitNouvelleRecherche()

        ' On lance une nouvelle recherche : ignorer la position précédente
        ' (on mémorise la position précédente uniquement lorsque l'on change l'affichage en cours)
        m_iNumParagSel = -1 : m_iNumPhraseSel = -1 : m_iNumCarSel = -1 : m_iLongSel = -1

    End Sub

    Private Function MarquerOccurrencesHtml(sPhraseAAfficher$, iNbCouleursHtml%) As StringBuilder

        If Not m_bOccurrencesEnCouleurs And Not m_bOccurrencesEnGras Then
            MarquerOccurrencesHtml = New StringBuilder(sPhraseAAfficher)
            Exit Function
        End If

        ' Mettre en gras les occurrences trouvées dans le html : <b> </b>
        Const sBaliseOuvCoulXX$ = "<SPAN class='OcXX'>" ' Oc pour Occurrence
        ' Caractère spécial ‡ : ne marche pas !?, aa non plus !?
        'Const sBaliseOuvCoulXX$ = "<SPAN class='aaXX'>"
        Const sCodeNumOcc$ = "XX"
        'Dim sBaliseOuv1$ = sBaliseOuvXX.Replace("XX", "1")
        Const sBaliseFermCoul$ = "</SPAN>"
        Const sBaliseOuvGras$ = "<b>"
        Const sBaliseFermGras$ = "</b>"
        Dim sBaliseOuvXX$ = ""
        Dim sBaliseFerm$ = ""

        If m_bOccurrencesEnCouleurs Then
            sBaliseOuvXX = sBaliseOuvCoulXX
            sBaliseFerm = sBaliseFermCoul
        End If
        If m_bOccurrencesEnGras Then
            sBaliseOuvXX &= sBaliseOuvGras
            sBaliseFerm = sBaliseFermGras & sBaliseFerm
        End If

        Dim sb As New StringBuilder
        Dim iNumExpression% = 0
        Dim iNbExpressions% = m_alExpressions.Count
        For Each sExpression As String In m_alExpressions

            ' Inconvénient : non prise en compte de la casse :
            'sPhrase = sPhrase.Replace(sExpression, "_" & sExpression & "_")

            Dim sBaliseOuv$ = sBaliseOuvXX
            If m_bOccurrencesEnCouleurs Then
                Dim iNumCouleurHtml% = iNumExpression Mod iNbCouleursHtml
                sBaliseOuv = sBaliseOuvXX.Replace(sCodeNumOcc, (iNumCouleurHtml + 1).ToString)
            End If

            ' Si la phrase contient une occurence qui est dans la balise elle-même
            '  alors on ne peut pas la surligner
            If sBaliseOuv.IndexOf(sExpression, m_styleCompare2) > -1 OrElse
               sBaliseFerm.IndexOf(sExpression, m_styleCompare2) > -1 Then
                iNumExpression += 1
                sb.Append(sPhraseAAfficher)
                Continue For
                'GoTo Suite
            End If

            Dim bTailleDifferente As Boolean = False
            Dim sPhraseAExaminer$
            If m_bIndexerAccents Then
                sPhraseAExaminer = sPhraseAAfficher
            Else
                Dim iMemLong0% = sPhraseAAfficher.Length
                sPhraseAExaminer = sEnleverAccents(sPhraseAAfficher)
                If sPhraseAExaminer.Length <> iMemLong0 Then
                    'Debug.WriteLine("!")
                    bTailleDifferente = True
                End If
            End If

            Dim iMemPosDebOcc% = 0
            Dim iDebRechOcc% = 0
            Dim iLong% = 0
            Dim iMemLong% = 0
            Do
                ' ToDo : lorsque prise en compte de la casse, alors changer l'option
                Dim iPosDebOcc% = sPhraseAExaminer.IndexOf(sExpression, iDebRechOcc, m_styleCompare2)
                If iPosDebOcc = -1 Then Exit Do
                Dim iLongPortionAv% = iPosDebOcc - iMemPosDebOcc - iMemLong

                ' 01/05/2019 Cas où l'occurrence se situe juste à la fin
                Dim iLongPAA% = sPhraseAAfficher.Length
                If iLongPortionAv <= iLongPAA Then

                    ' 05/08/2024
                    If iDebRechOcc + iLongPortionAv > iLongPAA Then
                        iLongPortionAv = iLongPAA - iDebRechOcc
                    End If

                    Dim sPortionAv$ = sPhraseAAfficher.Substring(iDebRechOcc, iLongPortionAv)
                    sb.Append(sPortionAv)
                End If
                'Dim s3$ = sb.ToString

                iLong = sExpression.Length

                ' Bug avec les lettres collées, par ex. cœur
                If bTailleDifferente Then
                    'Dim iLongPAA% = sPhraseAAfficher.Length
                    If iLong + iPosDebOcc > iLongPAA Then
                        iLong = iLongPAA - iPosDebOcc
                        If iLong <= 0 Then Exit Do ' 02/08/2010
                    End If
                End If

                Dim sOccurrence$ = sPhraseAAfficher.Substring(iPosDebOcc, iLong)
                sb.Append(sBaliseOuv & sOccurrence & sBaliseFerm)
                'Dim s2$ = sb.ToString
                iDebRechOcc = iPosDebOcc + iLong
                iMemPosDebOcc = iPosDebOcc
                iMemLong = iLong

            Loop While True
            If iLong < 0 Then iLong = 0 ' 01/05/2019
            sb.Append(sPhraseAAfficher.Substring(iMemPosDebOcc + iLong))

            'Suite:
            iNumExpression += 1
            If iNumExpression < iNbExpressions Then
                sPhraseAAfficher = sb.ToString
                sb = New StringBuilder
            End If

        Next
        MarquerOccurrencesHtml = sb
        'Dim s$ = MarquerOccurrencesHtml.ToString

    End Function

    Private Sub AfficherResultats(sExpressions$, alResultats As ArrayList,
        iNbZoomParag%, bAfficherInfoResultat As Boolean,
        bAfficherInfoDoc As Boolean, bAfficherNumParag As Boolean,
        bAfficherNumPhrase As Boolean,
        bAfficherNumOccur As Boolean, iNbOccurrencesTot%,
        bAfficherTiret As Boolean,
        sbResultat As StringBuilder, ByRef CtrlResultat As TextBox,
        alExpressions As ArrayList, bHtml As Boolean)

        Dim bTxt As Boolean = bHtml

        Dim bTailleLimite As Boolean = False

        Dim iNumPhrase%

        Dim sMemAffInfoDoc$ = ""
        Dim iNbPhrasesTrouvees% = 0

        ' Une seule phrase, sinon un ou plusieurs paragraphes avant et après
        Dim bAfficherUnePhrase As Boolean = False
        If iNbZoomParag = -1 Then bAfficherUnePhrase = True
        Dim sResultat0$ = ""
        Dim iMemNumParag% = -1 ' Saut de ligne sauf la 1ère fois
        Dim iMemNumParagraphe% = 0
        Dim iNumParagMin% = 0

        Dim iNumParagMax1%
        Dim iNbPhrasesMot% = m_colPhrases.Count
        Dim iMemNumPhraseMin% = 1
        Dim iMemNumPhraseG% = 0
        Dim iNumPhrase1ParagMotTrouve% = 0

        Dim bUnSeulDocument As Boolean
        'If Me.m_colDocs.Count() = 1 Then bUnSeulDocument = True
        ' 29/08/2010 Afficher quand même les chapitres s'il y en a
        Dim bUnSeulDocumentAvecChapitres As Boolean
        If Me.m_colDocs.Count() = 1 Then
            Dim oDoc As clsDoc = DirectCast(m_colDocs.Item(1), clsDoc)
            If oDoc.colChapitres.Count <= 1 Then
                bUnSeulDocument = True
            Else
                bUnSeulDocumentAvecChapitres = True
            End If
        End If

        Const iTailleLimiteInteger% = iMaxLongChaine - 4 ' Laisser de la place pour afficher "..."
        Const iTailleLimiteAffichageTextBox% = iMaxLongChaine0 - 4

        ' Afficher les phrases ou paragraphes correspondants
        Dim iNumOccurrence% = 0 ' Décompte de toutes les occurrences trouvées (ctrl web)
        Dim iNumOccurrenceAffichee% = 0 ' Décompte des occurrences affichées dans le ctrl textBox
        Dim iNbOccurrences% = alResultats.Count
        Dim sInfoOccurr$ = ""
        m_alExpressions = alExpressions
        Dim sEnteteHtml$ = "<html><body>"

        Dim iNumOcc% = 0
        If m_bOccurrencesEnCouleurs Then
            Dim asCouleursHtml$() = m_sCouleursHtml.Split(";"c)
            For Each sCouleur As String In asCouleursHtml
                If String.IsNullOrEmpty(sCouleur) Then Continue For
                iNumOcc += 1
                ' Oc pour Occurrence
                sEnteteHtml &= vbCrLf & "<STYLE type='text/css'>SPAN.Oc" & iNumOcc &
                    " { BACKGROUND-COLOR: " & sCouleur & " }</STYLE>"
            Next
        End If
        Dim iNbCouleursHtml% = iNumOcc
        'sEnteteHtml &= vbCrLf & "<STYLE type='text/css'>SPAN.Oc1 { BACKGROUND-COLOR: yellow }</STYLE>"
        'sEnteteHtml &= vbCrLf & "<STYLE type='text/css'>SPAN.Oc2 { BACKGROUND-COLOR: green }</STYLE>"
        'sEnteteHtml &= vbCrLf & "<STYLE type='text/css'>SPAN.Oc3 { BACKGROUND-COLOR: blue }</STYLE>"

        ' Pas obligatoire (mais il faudra préciser l'encodage au moment d'écrire le fichier) :
        'If m_bTexteUnicode Then sEnteteHtml = _
        '    "<html><meta http-equiv='content-type' content='text/html; charset=utf-8' /><body>"
        Const sPiedHtml$ = "</body></html>"
        Dim sbResultatHtml As StringBuilder = Nothing
        Dim sbResultatTxt As StringBuilder = Nothing
        If bHtml Then sbResultatHtml = New StringBuilder(sEnteteHtml & vbCrLf)
        If bTxt Then sbResultatTxt = New StringBuilder
        'Const sDebLigneHtml$ = "<li>"
        'Const sFinLigneHtml$ = "</li>"
        Const sSautLigneHtml$ = "<br>" & vbCrLf
        'Const sDebParagHtml$ = "<p>"
        'Const sFinParagHtml$ = "</p>"
        For Each iNumPhrase In alResultats
            iNumOccurrence += 1
            If sbResultat.Length <= iTailleLimiteAffichageTextBox Then _
                iNumOccurrenceAffichee += 1
            If bAfficherNumOccur Then sInfoOccurr = "(occ.n°" & iNumOccurrence & ") "

            ' 01/05/2010 Inutile
            'If iNumOccurrence Mod iModuloAvanvementLent = 0 Or iNumOccurrence = iNbOccurrences Then
            '    AfficherMessage("Affichage des occurrences en cours : " & _
            '        iNumOccurrence & " / " & iNbOccurrences)
            '    If m_bInterrompre Then Exit For
            'End If

            If iNumPhrase > m_colPhrases.Count Then ' 10/12/2022
                If bDebug Then Stop
                Continue For
            End If

            Dim oPhrase As clsPhrase = DirectCast(m_colPhrases.Item(iNumPhrase - 1), clsPhrase)

            ' 25/04/2010 Bug depuis la version V1.12 du 25/10/2009
            ' Un mot est présent plusieurs fois dans la même phrase
            If iNumPhrase = iMemNumPhraseG Then GoTo PhraseSuivante
            iMemNumPhraseG = iNumPhrase

            'Dim iNumParagraphe% = oPhrase.iNumParagrapheL ' Numéro de parag. local aux documents
            Dim iNumParagraphe% = oPhrase.iNumParagrapheG ' Numéro de parag. global aux documents

            ' Si la phrase suivante est dans le même §, alors elle a déjà été affichée
            ' (si on affiche que les phrases, alors iMemNumParagraphe reste à 0)
            If iNumParagraphe = iMemNumParagraphe Then GoTo PhraseSuivante

            Dim sAffInfoDoc$ = ""
            Dim sAffInfoDocHtml$ = ""
            Dim sInfos$ = ""
            Dim sInfosHtml$ = ""
            Dim sCleDoc$ = oPhrase.sCleDoc
            Dim iDecParagG2L% = 0
            Dim iDecPhraseG2L% = 0
            If bAfficherInfoResultat Then

                If bAfficherInfoDoc Then
                    If bUnSeulDocument Then
                        ' S'il n'y a qu'un seul document, inutile de le rappeler
                        sAffInfoDoc = ""
                        sAffInfoDocHtml = ""
                    Else
                        ' Trouver le document pour connaitre son CodeDoc edité
                        Dim oDoc As clsDoc
                        oDoc = DirectCast(m_colDocs.Item(oPhrase.sCleDoc), clsDoc)
                        Dim sCodeDoc$ = oDoc.sCodeDoc
                        Dim sChapitre$ = ""
                        If oPhrase.sCodeChapitre.Length > 0 Then
                            sCodeDoc &= ":" & oPhrase.sCodeChapitre
                            Dim sCleChap$ = oPhrase.sCleDoc & ":" & oPhrase.sCodeChapitre
                            Dim chapitre As clsChapitre = DirectCast(
                                oDoc.colChapitres(sCleChap), clsChapitre)
                            sChapitre = chapitre.sChapitre & " : "
                        End If
                        ' Note : le nom du document n'a pas encore été détecté
                        '  ce n'est pas forcément évident, car cela peut être la 1ère phrase
                        '  ou bien le nom du fichier (ou pour les documents word, une propriété)
                        '  conclusion : le chemin est le plus simple pour le moment
                        Dim s$ = oDoc.sChemin & " (" & sCodeDoc & ") : " & sChapitre
                        If bUnSeulDocumentAvecChapitres Then s = sChapitre
                        sAffInfoDoc = vbCrLf & s & vbCrLf
                        'sAffInfoDocHtml = vbCrLf & sDebParagHtml & s & sFinParagHtml & vbCrLf
                        sAffInfoDocHtml = sSautLigneHtml & s & sSautLigneHtml
                        If sAffInfoDoc = sMemAffInfoDoc Then
                            ' S'il n'a pas changé depuis le précédent
                            sAffInfoDoc = ""
                            sAffInfoDocHtml = ""
                        Else
                            sMemAffInfoDoc = sAffInfoDoc
                        End If
                    End If
                End If

                If m_bNumerotationGlobale Then
                    ' 11/10/2009 Afficher tjrs le n° de phrase global pour être cohérent
                    ' (le n° de § est global aussi)
                    Dim s$ = ""
                    If bAfficherNumParag Then _
                        s = sIndicParag & iNumParagraphe & " " '" Ph. n°" & oPhrase.iNumPhraseG & " "
                    If bAfficherNumPhrase Then s &= sIndicPhrase & oPhrase.iNumPhraseG & " "
                    sInfos = sAffInfoDoc & s
                    sInfosHtml = sAffInfoDocHtml & s
                Else
                    iDecParagG2L = oPhrase.iNumParagrapheG - oPhrase.iNumParagrapheL
                    iDecPhraseG2L = oPhrase.iNumPhraseG - oPhrase.iNumPhraseL
                    Dim s$ = ""
                    If bAfficherNumParag Then _
                        s = sIndicParag & iNumParagraphe - iDecParagG2L & " "
                    '" Ph. n°" & oPhrase.iNumPhraseL & " "
                    If bAfficherNumPhrase Then s &= sIndicPhrase & oPhrase.iNumPhraseL & " "
                    sInfos = sAffInfoDoc & s
                    sInfosHtml = sAffInfoDocHtml & s
                End If
                If bAfficherNumOccur Then sInfos &= sInfoOccurr : sInfosHtml &= sInfoOccurr

            End If

            If bAfficherUnePhrase Then
                Dim sTiret$ = ""
                If bAfficherTiret Then sTiret = "- " ' 29/05/2015 Optionnel
                Dim s$ = sInfos & sTiret & oPhrase.sPhrase & vbCrLf
                sbResultat.Append(s)
                If bTxt Then sbResultatTxt.Append(s)
                If bHtml Then sbResultatHtml.Append(sInfosHtml & sTiret).Append(
                    MarquerOccurrencesHtml(oPhrase.sPhrase, iNbCouleursHtml)).Append(sSautLigneHtml)
                GoTo PhraseSuivante
            End If


            If iNumParagraphe = iNumParagMax1 Then GoTo PhraseSuivante

            ' Examiner l'occurrence suivante
            Dim iNumParagMax2%
            iNumParagMax2 = -1
            Dim iNumPhraseMot% = iNumPhrase
            Dim iNumOccurrence0% = iNumOccurrence
            While iNumOccurrence0 < iNbOccurrences

                ' (n° de phrase) Global de l'occurrence suivante (+1) : GP1 : GlobPlus1
                Dim iNumPhraseGP1 = CInt(alResultats(iNumOccurrence0))
                Dim iNumParagGP1% = iLireNumParagGPhrase(iNumPhraseGP1)
                If iNumParagGP1 = iNumParagraphe Then
                    ' L'occurrence suivante est dans le même § : voir l'occurrence suivante
                    iNumOccurrence0 += 1
                    Continue While
                End If
                ' Vérifier si l'occ. suiv. est tjrs dans le même doc.
                Dim sCleDocPhraseGP1$ = Me.sLireCleDocPhrase(iNumPhraseGP1)
                If sCleDocPhraseGP1 = sCleDoc Then
                    If iNumParagGP1 > iNumParagraphe + 2 * iNbZoomParag Then
                        'iNumParagMax2 = 0 ' Déjà traité
                    Else
                        iNumParagMax2 = iNumParagraphe + (iNumParagGP1 - iNumParagraphe) \ 2
                    End If
                End If
                Exit While
            End While

            ' Rechercher le n° de phrase du début du parag contenant le mot trouvé
            ' Algorithme : trouver la première phrase appartenant au parag précédent
            Dim iNumPhraseG% = iNumPhrase
            iNumPhrase1ParagMotTrouve = iNumPhraseG ' Initialisation par défaut
            Dim j%
            For j = iNumPhraseG - 1 To 1 Step -1
                If bInterruption() Then GoTo FinRecherche
                ' ToDo : Dans cette boucle, on caste 3 fois la phrase : à optimiser
                If Me.sLireCleDocPhrase(j) <> sCleDoc Then Exit For
                Dim iNumParag_Phr_j% = iLireNumParagGPhrase(j)
                If iNumParag_Phr_j = iNumParagraphe Then
                    ' La phrase précédente appartient au même paragraphe
                    '  elle doit donc être inclue dans le paragraphe courant
                    iNumPhrase1ParagMotTrouve = j
                Else
                    ' La phrase précédente appartient au paragraphe précédent
                    '  l'affichage du paragraphe commence donc à la phrase suivante
                    iNumPhrase1ParagMotTrouve = j + 1 : Exit For
                End If
            Next j

            ' Puis noter toutes les phrases du paragraphe +- l'écart demandé
            Dim iNbParagAv%, iNbParagAp%
            iNbParagAv = iNbZoomParag : iNbParagAp = iNbZoomParag

            ' Rechercher les n° de § précédants
            Dim iMin% = 1
            Dim iMemNumPhrasePreced% = iNumPhrase1ParagMotTrouve
            ' Cas où plusieurs § successifs contiennent le mot : un seul affichage
            If iMemNumPhrasePreced < iMemNumPhraseMin Then
                iMemNumPhrasePreced = iMemNumPhraseMin
                iMin = iMemNumPhraseMin
            End If
            Dim iNumPhraseDebRech% = iMemNumPhrasePreced
            For j = iNumPhrase1ParagMotTrouve To iMin Step -1
                If bInterruption() Then GoTo FinRecherche
                If Me.sLireCleDocPhrase(j) <> sCleDoc Then Exit For
                Dim iNumParag_Phr_j% = iLireNumParagGPhrase(j)
                ' Noter le n° global de la phrase en cours
                'If j = iNumPhrase1ParagMotTrouve Then iMemNumPhraseG = oPhrase.iNumPhraseG
                ' Ne pas afficher plusieurs fois le même paragraphe
                If j < iMemNumPhraseMin Then _
                    iNumPhraseDebRech = iMemNumPhrasePreced : Exit For
                If iNumParag_Phr_j < iNumParagraphe - iNbParagAv Then _
                    iNumPhraseDebRech = iMemNumPhrasePreced : Exit For
                iMemNumPhrasePreced = j : iNumPhraseDebRech = j
            Next j

            ' Rechercher les n° de § suivants
            iMemNumPhrasePreced = iNumPhrase1ParagMotTrouve
            Dim iNumPhraseFinRech% = iMemNumPhrasePreced
            For j = iNumPhrase1ParagMotTrouve To Me.iNbPhrasesG
                If bInterruption() Then GoTo FinRecherche
                Dim iNumParag_Phr_j% = iLireNumParagGPhrase(j)
                If iNumParag_Phr_j > iNumParagraphe Then iMemNumPhraseMin = j
                If iNumParag_Phr_j > iNumParagraphe + iNbParagAp Then _
                    iNumPhraseFinRech = iMemNumPhrasePreced : Exit For
                If iNumParagMax2 > -1 And iNumParag_Phr_j > iNumParagMax2 Then
                    iNumPhraseFinRech = iMemNumPhrasePreced : Exit For
                End If
                ' Ne pas afficher 2x le dernier §
                If j = Me.iNbPhrasesG Then iNumParagMax1 = oPhrase.iNumParagrapheL
                iMemNumPhrasePreced = j : iNumPhraseFinRech = j
            Next j

            ' Afficher les § précédents et suivants demandés
            For j = iNumPhraseDebRech To iNumPhraseFinRech
                If bInterruption() Then GoTo FinRecherche
                oPhrase = DirectCast(m_colPhrases.Item(j - 1), clsPhrase)
                Dim iNumParag_Phr_j% = iLireNumParagGPhrase(j)

                Dim sIndicParagFinal$ = ""
                If iNumParag_Phr_j < iNumParagraphe Then
                    sIndicParagFinal = "< "
                ElseIf iNumParag_Phr_j > iNumParagraphe Then
                    sIndicParagFinal = "> "
                ElseIf bAfficherTiret OrElse iNbZoomParag > 0 Then ' 29/05/2015 Optionnel
                    sIndicParagFinal = "- "
                End If

                If (j = iNumPhraseDebRech Or
                    iNumParag_Phr_j > iMemNumParag) And iMemNumParag > -1 Then
                    sbResultat.Append(vbCrLf) ' Nouv. Parag 
                    If bTxt Then sbResultatTxt.Append(vbCrLf)
                    If bHtml Then sbResultatHtml.Append(sSautLigneHtml)
                End If
                If j = iNumPhraseDebRech Or iNumParag_Phr_j > iMemNumParag Then
                    If bAfficherInfoResultat Then
                        Dim s$ = ""
                        If bAfficherNumParag Then _
                            s = sIndicParag & iNumParag_Phr_j - iDecParagG2L & " "
                        '" Ph. n°" & j - iDecPhraseG2L & " "
                        If bAfficherNumPhrase Then s &= sIndicPhrase & j - iDecPhraseG2L & " "
                        sbResultat.Append(sAffInfoDoc & s)
                        If bTxt Then sbResultatTxt.Append(sAffInfoDoc & s)
                        If bAfficherNumOccur Then sbResultat.Append(sInfoOccurr)
                        If bAfficherNumOccur And bTxt Then sbResultatTxt.Append(sInfoOccurr)
                        If bHtml Then sbResultatHtml.Append(sAffInfoDocHtml & s)
                        If bAfficherNumOccur And bHtml Then sbResultatHtml.Append(sInfoOccurr)
                    End If
                    sbResultat.Append(sIndicParagFinal)
                    If bTxt Then sbResultatTxt.Append(sIndicParagFinal)
                    If bHtml Then sbResultatHtml.Append(sIndicParagFinal)
                End If

                If bDebug Then sResultat0 &= oPhrase.sPhrase
                sbResultat.Append(oPhrase.sPhrase)
                If bTxt Then sbResultatTxt.Append(oPhrase.sPhrase)
                If bHtml Then sbResultatHtml.Append(MarquerOccurrencesHtml(oPhrase.sPhrase, iNbCouleursHtml))
                sAffInfoDoc = "" ' N'afficher qu'une seule fois le chemin
                sAffInfoDocHtml = ""
                If bDebug Then sResultat0 = ""

                iMemNumParag = iNumParag_Phr_j
            Next j

            iMemNumParagraphe = iNumParagraphe

PhraseSuivante:
            'If sbResultat.Length > iTailleLimite Then bTailleLimite = True : Exit For
            If sbResultat.Length > iTailleLimiteInteger Then Exit For

        Next

        If bHtml Then
            sbResultatHtml.Append(sPiedHtml)
            m_sbResultatHtml = sbResultatHtml
        End If
        If bTxt Then m_sbResultatTxt = sbResultatTxt

        AfficherMessage("Affichage des résultats...")

        Dim iLen1% = sbResultat.Length
        If iLen1 > iTailleLimiteAffichageTextBox Then
            bTailleLimite = True
            Dim sbResultat0 As New StringBuilder
            sbResultat0.Length = 0
            'sbResultat0.Append(sbResultat.ToString.Substring(0, iTailleLimite))
            sbResultat0.Append(sbResultat.ToString.Substring(0, iTailleLimiteAffichageTextBox))
            sbResultat0.Append("...")
            'If CtrlResultat.Text <> sbResultat0.ToString Then
            If String.Compare(CtrlResultat.Text, sbResultat0.ToString) <> 0 Then
                ' C'est cette ligne qui prend du temps
                CtrlResultat.Text = sbResultat0.ToString
            End If
        Else
            If m_bInterrompre Then sbResultat.Append("...")
            'If CtrlResultat.Text <> sbResultat.ToString Then
            If String.Compare(CtrlResultat.Text, sbResultat.ToString) <> 0 Then
                ' C'est cette ligne qui prend du temps
                'CtrlResultat.Text = sbResultat.ToString
                ' 01/05/2010 On va le faire en deux temps
                Const iPrevisu% = 5000
                If sbResultat.Length > iPrevisu Then
                    Dim sbResultat0 As New StringBuilder
                    sbResultat0.Append(sbResultat.ToString.Substring(0, iPrevisu))
                    sbResultat0.Append("...")
                    CtrlResultat.Text = sbResultat0.ToString
                    Application.DoEvents()
                    ' Parfois le sablier n'a pas été bien activé 
                    ' (car le ctrl text avait encore le focus ?)
                    Sablier()
                End If
                CtrlResultat.SuspendLayout()
                CtrlResultat.Text = sbResultat.ToString
                CtrlResultat.ResumeLayout()
            End If
        End If

FinRecherche:
        iNbPhrasesTrouvees = iNumOccurrenceAffichee 'iNumOccurrence
        If m_bInterrompre Or bTailleLimite Then
            If iNbOccurrencesTot = -1 Then
                ' Cas des expressions : on ne connait pas le nombre total d'occurences
                AfficherMessage(sExpressions & " : Nombre d'occurrences affichées : " &
                    iNbPhrasesTrouvees)
            Else
                ' Cas des mots : on connait le nbre total via l'index
                AfficherMessage(sExpressions & " : Nombre d'occurrences affichées : " &
                    iNbPhrasesTrouvees & " / " & iNbOccurrencesTot & " trouvées")
            End If
        Else
            ' Si on connait le total, alors il est plus fiable
            If iNbOccurrencesTot <> -1 Then iNbPhrasesTrouvees = iNbOccurrencesTot
            AfficherMessage(sExpressions & " : Nombre d'occurrences trouvées : " &
                iNbPhrasesTrouvees)
        End If

    End Sub

    Public Sub NoterPositionCurseur(CtrlResultat As Windows.Forms.TextBox,
        bAfficherInfoResultat As Boolean, bAfficherNumParag As Boolean,
        bAfficherNumPhrase As Boolean)
        'Debug.WriteLine(Now & " : Memo pos. curseur")

        m_iNumParagSel = -1
        m_iNumPhraseSel = -1
        m_iNumCarSel = -1
        m_iLongSel = -1

        ' On a besoin de la numérotation globale pour que cela marche :
        If Not m_bNumerotationGlobale Then Exit Sub
        If Not bAfficherInfoResultat Then Exit Sub
        ' On a besoin des 2 repères § et Ph. pour que cela marche :
        If Not bAfficherNumParag Then Exit Sub
        If Not bAfficherNumPhrase Then Exit Sub

        NoterPositionCurseur2(CtrlResultat,
            m_iNumParagSel, m_iNumPhraseSel, m_iNumCarSel, m_iLongSel)

    End Sub

    Private Sub NoterPositionCurseur2(CtrlResultat As Windows.Forms.TextBox,
        ByRef iNumParagSel%, ByRef iNumPhraseSel%, ByRef iNumCarSel%, ByRef iLongSel%)

        'iNumCarSel    : n° du car. sel. dans la phrase en cours
        'iNumPhraseSel : n° de la phrase sel. global
        'iNumParagSel  : n° du parag. sel. global
        iNumParagSel = -1
        iNumPhraseSel = -1
        iNumCarSel = -1
        iLongSel = -1

        Const sReperePhrase$ = " " & sIndicPhrase '" Ph. n°"
        ' Recherche du parag. courant : contenant le mot actuellement sélectionné
        Dim sCtrlResultat$ = CtrlResultat.Text.ToString
        Dim iPosDebParagSel% = -1
        Dim iNumCarDebSelCtrl%
        ' Noter la sélection en cours dans le ctrl pour tenter de la restituer
        iNumCarDebSelCtrl = CtrlResultat.SelectionStart
        iLongSel = CtrlResultat.SelectionLength
        If iNumCarDebSelCtrl >= 0 And iNumCarDebSelCtrl < sCtrlResultat.Length Then _
            iPosDebParagSel = sCtrlResultat.LastIndexOf(sCarParag, iNumCarDebSelCtrl)

        If iPosDebParagSel > -1 Then
            Dim iLongMax% = 40
            If sCtrlResultat.Length < iLongMax + iPosDebParagSel Then _
                iLongMax = sCtrlResultat.Length - iPosDebParagSel
            If iLongMax < 0 Then iLongMax = 0
            Dim sLigne$ = sCtrlResultat.Substring(iPosDebParagSel, iLongMax)
            Dim iLen% = sLigne.Length
            Dim iCarNumero% = sLigne.IndexOf("°")
            If iCarNumero = -1 Then GoTo Fin
            Dim iCarEspace% = sLigne.IndexOf(" ", iCarNumero)
            If iCarEspace = -1 Then GoTo Fin
            Dim sNumParagSel$ = ""
            If iCarNumero + 1 >= 0 And iCarEspace - iCarNumero - 1 <= iLen Then
                sNumParagSel = sLigne.Substring(iCarNumero + 1, iCarEspace - iCarNumero - 1)
                'iNumParagSel = CInt(sNumParagSel)
                iNumParagSel = iConv(sNumParagSel)
            End If
            If iNumParagSel = -1 Then GoTo Fin
            Dim iCarPhrase% = sLigne.IndexOf(sReperePhrase, iCarEspace)
            If iCarPhrase = -1 Then GoTo Fin
            Dim iCarEspace2% = 0
            'iDebLigne : nbr de car. en partant de la gauche du ctrl, de la pos. du curseur
            ' même paragraphe : utile lors du passage de Phrase vers Parag.
            Dim iDebLigne% = iCarEspace + 3
            Dim sNumPhrase$ = ""
            If iCarPhrase > -1 Then
                iCarEspace2 = sLigne.IndexOf(" ", iCarPhrase + sReperePhrase.Length)
                If iCarEspace2 = -1 Then GoTo Fin
                iDebLigne = iCarEspace2 + 3
                Dim iPosNum% = iCarPhrase + sReperePhrase.Length
                If iPosNum >= 0 And iCarEspace2 - iPosNum <= iLen Then
                    sNumPhrase = sLigne.Substring(iPosNum, iCarEspace2 - iPosNum)
                    'iNumPhraseSel = CInt(sNumPhrase)
                    iNumPhraseSel = iConv(sNumPhrase)
                End If
            End If
            iNumCarSel = iNumCarDebSelCtrl - iPosDebParagSel - iDebLigne

            If iNumPhraseSel = -1 Then GoTo Fin

            iDebLigne = iCarEspace2 + 3
            Const bDebugPos As Boolean = False
            If bDebugPos Then _
                MsgBox("DebLigne=" & iDebLigne & ", Car=" & iNumCarSel &
                    ", Ph:" & iNumPhraseSel & ", §:" & iNumParagSel)

            ' Si on passe de § à phrase, décaller le curseur au début de la phrase effective
            Dim iLen2% = 0
            Do
                Dim sPhrase$ = sLirePhrase(iNumPhraseSel)
                iLen2 = sLirePhrase(iNumPhraseSel).Length
                If iNumCarSel < iLen2 Then Exit Do
                iNumCarSel -= iLen2
                If iNumPhraseSel >= Me.iNbPhrasesG Then Exit Do
                iNumPhraseSel += 1
                If bDebugPos Then _
                        MsgBox("DebLigne=" & iDebLigne & ", Car=" & iNumCarSel &
                            ", Ph:" & iNumPhraseSel & ", §:" & iNumParagSel)
            Loop While True

        End If

Fin:
        'Debug.WriteLine(Now & " : Pos. curseur : §" & iNumParagSel & ", Ph." & _
        '    iNumPhraseSel & ", Car." & iNumCarSel & ", Long." & iLongSel)

    End Sub

    Private Sub RestaurerPositionCurseur(CtrlResultat As Windows.Forms.TextBox, sMot$,
        iNumParagSel%, iNumPhraseSel%, iNumCarSel%, iLongSel%,
        iNbZoomParag%, sbResultat As StringBuilder)

        Static sMemMot$ = ""
        If Not m_bNumerotationGlobale OrElse sMemMot <> sMot Then

            CtrlResultat.Select()
            CtrlResultat.SelectionStart = 0
            CtrlResultat.SelectionLength = 0
            CtrlResultat.ScrollToCaret()

        ElseIf iNumParagSel > -1 Then

            ' Analyse du parag courant
            Dim iCumulLongPhrPrecedParag% = 0
            'iCumulLongPhrPrecedParag : Cumul des longueurs des phrases précédentes du §
            ' 24/10/2009 And iNumPhraseSel > -1
            If sMemMot = sMot And iNbZoomParag > -1 And iNumPhraseSel > -1 Then
                Dim sCleDoc$ = Me.sLireCleDocPhrase(iNumPhraseSel)
                For j = iNumPhraseSel - 1 To 1 Step -1
                    ' ToDo : Dans cette boucle, on caste 3 fois la phrase : à optimiser
                    If Me.sLireCleDocPhrase(j) <> sCleDoc Then Exit For
                    Dim iNumParag_Phr_j% = -1
                    iNumParag_Phr_j = iLireNumParagGPhrase(j)
                    If iNumParag_Phr_j < iNumParagSel Then Exit For
                    ' La phrase précédente appartient au même paragraphe
                    '  elle doit donc être inclue dans le paragraphe courant
                    If iNumParag_Phr_j = iNumParagSel And j < iNumPhraseSel Then
                        'And iMemZoomParag = -1 'si on passe de Phrase à § 
                        ' Décaller la sélection de la longueur de la phrase précéd.
                        ' Si même §, même mot, si on est en §
                        iCumulLongPhrPrecedParag += sLirePhrase(j).Length
                    End If
                Next j
            End If

            CtrlResultat.SuspendLayout() ' Eviter le scintillement pdt le focus
            CtrlResultat.Select() ' Focus
            ' Déselectionner après Select
            CtrlResultat.SelectionStart = 0
            CtrlResultat.SelectionLength = 0
            CtrlResultat.ResumeLayout()

            ' Sélection du paragraphe suivant pour rendre visible le précédent
            Dim sRepere$ = sIndicParag & (iNumParagSel + 1) & " "
            Dim sResultat$ = sbResultat.ToString
            Dim iPos% = sResultat.IndexOf(sRepere)
            If iPos = -1 Then
                ' 25/10/2009 Le prochain § n'a pas été trouvé (passage en mode phrase seul) :
                '  dans ce cas il faut d'abord rechercher le § en cours
                sRepere = sIndicParag & iNumParagSel & " "
                iPos = sResultat.IndexOf(sRepere)
                If iPos > -1 Then
                    ' Puis le prochain saut de ligne
                    Dim iPos0% = sResultat.IndexOf(vbLf, iPos)
                    If iPos0 > -1 Then iPos = iPos0
                End If
            End If
            If iPos > -1 Then
                CtrlResultat.SelectionStart = iPos
                CtrlResultat.SelectionLength = 0
                CtrlResultat.ScrollToCaret()
            End If
            ' Sélection du paragraphe courant maintenant
            sRepere = sIndicParag & iNumParagSel & " " & sIndicPhrase & iNumPhraseSel & " "
            iPos = sResultat.IndexOf(sRepere)
            If iPos > -1 Then

                CtrlResultat.SelectionStart = iPos
                CtrlResultat.SelectionLength = 0

                CtrlResultat.ScrollToCaret()

                ' Si possible retrouver la position du curseur dans le parag. sel.
                Dim iSelStart% = iPos + sRepere.Length + 2 +
                    iCumulLongPhrPrecedParag + iNumCarSel
                If iSelStart > -1 And iNumCarSel > -1 Then
                    ' C'est possible
                    CtrlResultat.SelectionStart = iSelStart
                    CtrlResultat.SelectionLength = iLongSel
                Else
                    ' Echec : selectionner seulement le repère
                    CtrlResultat.SelectionLength = sRepere.Length - 1
                End If

            Else

                ' Si la phrase exacte n'a pu etre retrouvée et que l'on est en mode phrase
                '  alors on perd le mot sélectionné
                If iNbZoomParag = -1 Then iLongSel = 0 : iNumCarSel = -1

                ' Si les phrases sont regroupées en parag. ne pas tenter de rech. la phr.
                sRepere = sIndicParag & iNumParagSel & " " ' Ph. n°" & iNumPhraseSel & " "
                iPos = sResultat.IndexOf(sRepere)
                If iPos > -1 Then
                    CtrlResultat.SelectionStart = iPos
                    CtrlResultat.SelectionLength = 0
                    CtrlResultat.ScrollToCaret()

                    Dim iSelStart% = iPos + sRepere.Length + 2 +
                        iCumulLongPhrPrecedParag + iNumCarSel

                    ' Retrouver le début de la première phrase
                    Dim sRepere2$ = " " & sIndicPhrase '" Ph. n°"
                    Dim iPos2% = sResultat.IndexOf(sRepere2, iPos)
                    If iPos2 > 1 Then
                        iPos2 = sResultat.IndexOf(" ", iPos2 + sRepere2.Length)
                        iSelStart = iPos2 + 3 + iCumulLongPhrPrecedParag + iNumCarSel
                    End If

                    ' Si possible retrouver la position du curseur dans le parag. sel.
                    If iSelStart > -1 And iNumCarSel > -1 Then
                        ' C'est possible
                        CtrlResultat.SelectionStart = iSelStart
                        CtrlResultat.SelectionLength = iLongSel
                    Else
                        ' Echec : selectionner seulement le repère
                        CtrlResultat.SelectionLength = sRepere.Length - 1
                    End If
                End If

            End If
        End If
        sMemMot = sMot

    End Sub

    Private Sub AjouterMotDejaTrouve(sExpressions$, ByRef CtrlMot As Windows.Forms.ComboBox)

        ' Ajouter le mot à la combobox des mots déjà recherchés,
        '  si ce n'est pas déjà fait
        Dim sExpressionsMin$ = sExpressions.ToLower
        Dim bDejaRecherche As Boolean
        For j = 0 To CtrlMot.Items.Count - 1
            If DirectCast(CtrlMot.Items(j), String).ToLower = sExpressionsMin Then _
                bDejaRecherche = True : Exit For
        Next j
        If Not bDejaRecherche Then CtrlMot.Items.Add(sExpressions)

    End Sub

    'Private Function iLireNumParagLPhrase%(iNumPhraseG%)

    '    ' Lire le n° de paragraphe local en fonction du n° de phrase global

    '    Dim oPhrase As clsPhrase
    '    oPhrase = DirectCast(m_colPhrases.Item(iNumPhraseG - 1), clsPhrase)
    '    iLireNumParagLPhrase = oPhrase.iNumParagrapheL

    'End Function

    Private Function iLireNumParagGPhrase%(iNumPhraseG%)

        ' Lire le n° de paragraphe global en fonction du n° de phrase global

        If iNumPhraseG < 1 OrElse iNumPhraseG > Me.iNbPhrasesG Then
            iLireNumParagGPhrase = -1 : Exit Function
        End If
        Dim oPhrase As clsPhrase
        oPhrase = DirectCast(m_colPhrases.Item(iNumPhraseG - 1), clsPhrase)
        iLireNumParagGPhrase = oPhrase.iNumParagrapheG

    End Function

    Private Function sLirePhrase$(iNumPhraseG%)

        ' Lire la phrase en fonction du n° de phrase global

        If iNumPhraseG < 1 OrElse iNumPhraseG > Me.iNbPhrasesG Then
            sLirePhrase = "" : Exit Function
        End If
        Dim oPhrase As clsPhrase
        oPhrase = DirectCast(m_colPhrases.Item(iNumPhraseG - 1), clsPhrase)
        sLirePhrase = oPhrase.sPhrase

    End Function

    Public Function bHyperTexte(ByRef sMotSel$, ByRef sMotSelFin$) As Boolean

        ' Traitement du mode hypertexte

        Dim iLongMot%, iDeb%, iFin%
        Dim sCar1$, sCar2$
        If sMotSel = "" Then Return False

        ' Extraction d'un mot bien délimité
        ' Vérifier s'il y a une virgule à la fin
        If bSeparateurMots(Right(sMotSel, 1)) Then _
            sMotSel = Left(sMotSel, Len(sMotSel) - 1)
        ' Tester aussi le cas du .
        If bSeparateurPhrases(Right(sMotSel, 1)) Then _
            sMotSel = Left(sMotSel, Len(sMotSel) - 1)
        iLongMot = Len(sMotSel)
        For iDeb = 1 To iLongMot - 1
            ' Vérifier s'il y a : l'
            sCar1 = Mid(sMotSel, iDeb, 1)
            sCar2 = Mid(sMotSel, iDeb + 1, 1)
            If Not bSeparateurMots(sCar1) And (Not bSeparateurMots(sCar2) Or iLongMot < 4) Then Exit For
        Next iDeb
        If iDeb = iLongMot Then sMotSelFin = Right(sMotSel, 1) : Return False
        ' Vérifier les mots composés : c'est-à-dire : est
        For iFin = iDeb + 1 To iLongMot
            sCar1 = Mid(sMotSel, iFin, 1)
            If bSeparateurMots(sCar1) Then Exit For
        Next iFin
        sMotSelFin = Mid(sMotSel, iDeb, iFin - iDeb)
        ' Inutile de lancer une recherche automatique pour des mots de moins de 3 lettres
        If Len(Trim(sMotSelFin)) < 3 Then Return False
        Return True

    End Function

    Private Function bContientSeparateurPhrases(sMot$) As Boolean

        ' Indiquer si le mot contient un séparateur de phrases

        Dim i%, iLen%
        iLen = Len(sMot)
        For i = 1 To iLen
            If bSeparateurPhrases(Mid(sMot, i, 1)) Then Return True
        Next i
        ' Ce mot n'en contient pas
        Return False

    End Function

    Private Function bSeparateurPhrases(sCar$) As Boolean

        ' Indiquer si le caractère est un séparateur de phrases

        If InStr(Me.m_sListeSeparateursPhrase, sCar) > 0 Then Return True
        Return False

    End Function

    Public Function bSeparateurMots(sCar$) As Boolean

        ' Indiquer si le caractère est un séparateur de mots
        If InStr(m_sListeSeparateursMot, sCar) > 0 Then Return True
        Return False

    End Function

#End Region

#Region "Sérialisation de l'index"

    Private Function bValiderSauvegardeTmp() As Boolean

        ' Conserver la sauvegarde précédente (si elle existe) :
        '  renommer le fichier VBTxtFnd.idx en VBTxtFnd.bak
        ' Valider la sauvergarde temporaire (si elle n'existe pas, la créer) :
        '  renommer le fichier VBTxtFnd.tmp en VBTxtFnd.idx

        ' Si le fichier .tmp n'existe pas, on sauvegarde l'index
        If Not bFichierExiste(m_sCheminVBTxtFndTmp) Then _
            If Not bSauvegarderIndex(m_sCheminVBTxtFndTmp) Then Return False
        ' Renommer le fichier VBTxtFnd.idx en VBTxtFnd.bak
        If Not bRenommerFichier(m_sCheminVBTxtFndIdx, m_sCheminVBTxtFndBak) Then Return False
        ' Renommer le fichier VBTxtFnd.tmp en VBTxtFnd.idx
        If Not bRenommerFichier(m_sCheminVBTxtFndTmp, m_sCheminVBTxtFndIdx) Then Return False
        Return True

    End Function

    Private Function bSauvegarderIndex(sCheminFichierIndex$) As Boolean

        ' Sauvegarder l'index dans le fichier VBTextFinder.idx

        If Not bFichierAccessible(sCheminFichierIndex,
            bPrompt:=True, bInexistOk:=True) Then Return False

        bSauvegarderIndex = False
        Sablier()
        AfficherMessage("Sauvegarde de l'index en cours...")
        LireListeDocumentsIndexesIni()

        Dim iEncodage% = iCodePageWindowsLatin1252
        If m_bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8

        Dim fs As IO.FileStream = Nothing
        Try
            fs = New IO.FileStream(sCheminFichierIndex, IO.FileMode.Create, IO.FileAccess.Write)
            Using bw As New IO.BinaryWriter(fs, Encoding.GetEncoding(iEncodage))

                Dim rVersion! = rVersionFichierVBTxtFndIdx10
                If m_bIndexerChapitre Then rVersion = rVersionFichierVBTxtFndIdx
                bw.Write(rVersion)

                ' Sauvegarder le nombre de documents indexés
                Dim iNbDocs%, iNbMots%
                iNbDocs = Me.m_colDocs.Count()
                iNbMots = Me.m_htMots.Count()
                bw.Write(iNbDocs)
                bw.Write(iNbMots) ' Nbr de mots distincts indexés
                bw.Write(Me.iNbPhrasesG)
                ' Réserver de la place pour compléter les statistiques générales
                '  dans une version future (afin de conserver la compatibilité du fichier index)
                bw.Write(Me.iNbParagG) ' Nombre de paragraphes indexés en tout
                bw.Write(Me.iNbMotsG)  ' Nombre de mots indexés en tout

                ' Non : Nombre de caractères y compris les séparateurs de mot : à faire
                'bw.Write(CInt(0))  'iNbCarDontSeparIndexes
                ' 22/05/2010 bUnicode ou pas
                Dim iOptionsEncodage% = 0
                If m_bOptionTexteUnicode Then iOptionsEncodage += iMasqueOptionUnicode
                If m_bIndexerAccents Then iOptionsEncodage += iMasqueOptionAccent
                bw.Write(iOptionsEncodage)

                ' Sauvegarder la liste des documents indexés
                'Dim de As DictionaryEntry
                'For Each de In m_colDocs
                '    Dim oDoc As clsDoc = DirectCast(de.Value, clsDoc)
                'Next de 
                Dim oDoc As clsDoc
                For Each oDoc In Me.m_colDocs
                    bEcrireChaine(bw, oDoc.sCle)
                    bEcrireChaine(bw, oDoc.sCodeDoc)
                    'Debug.WriteLine(oDoc.sCle & ", " & oDoc.sCodeDoc)
                    bEcrireChaine(bw, oDoc.sChemin)
                    ' Réserver de la place pour compléter les statistiques par document
                    Dim iVal% = 0
                    bw.Write(iVal) 'oDoc.iNbMotsIndexes 
                    bw.Write(iVal) 'oDoc.iNbPhrasesIndexees
                    bw.Write(iVal) 'oDoc.iNbParagIndexes
                    bw.Write(iVal) 'oDoc.iNbCarIndexes
                    ' Nombre de caractères dont les séparateurs de mot
                    ' 17/07/2010 Finalement, le dernier Int32 va servir à indiquer le 
                    '  nombre de chapitres trouvés dans le document
                    If m_bIndexerChapitre Then iVal = oDoc.colChapitres.Count
                    bw.Write(iVal) 'iNbChapitresDoc 'oDoc.iNbCarDontSeparIndexes

                    If m_bIndexerChapitre Then
                        'Dim iNbChapitresDoc% = oDoc.colChapitres.Count
                        EcrireChapitre(bw, oDoc)
                        'For Each oChap As clsChapitre In oDoc.colChapitres
                        '    bEcrireChaine(bw, oChap.sCle) ' La clé contient toutes les infos.
                        '    bEcrireChaine(bw, oChap.sChapitre)
                        'Next
                    End If

                Next oDoc

                ' Sauvegarder les mots de l'index
                Dim i%, iNumMot%
                Dim oMot As clsMot

                Dim de As DictionaryEntry
                For Each de In Me.m_htMots
                    oMot = DirectCast(de.Value, clsMot)

                    iNumMot += 1

                    If iNumMot Mod iModuloAvanvementLent = 0 Or iNumMot = iNbMots Then
                        AfficherMessage("Sauvegarde des mots en cours : " &
                    iNumMot & " / " & iNbMots)
                        If m_bInterrompre Then GoTo Interruption
                    End If

                    bEcrireChaine(bw, oMot.sMot)
                    bw.Write(oMot.iNbOccurrences)
                    bw.Write(oMot.iNbPhrases)
                    ' Nombre de phrases dans lesquelles ce mot figure
                    For i = 1 To oMot.iNbPhrases
                        bw.Write(oMot.iLireNumPhrase(i))
                    Next i
                Next de 'oMot

                ' Sauvegarder les phrases de l'index
                Dim oPhrase As clsPhrase
                Dim iNumPhrase As Integer
                For Each oPhrase In m_colPhrases
                    'For Each de In m_colPhrases
                    '    oPhrase = DirectCast(de.Value, clsPhrase)

                    iNumPhrase = iNumPhrase + 1
                    If iNumPhrase Mod iModuloAvanvementLent = 0 Or iNumPhrase = Me.iNbPhrasesG Then
                        AfficherMessage("Sauvegarde des phrases en cours : " &
                    iNumPhrase & " / " & Me.iNbPhrasesG)
                        If m_bInterrompre Then GoTo Interruption
                    End If

                    bEcrireChaine(bw, oPhrase.sClePhrase)
                    bEcrireChaine(bw, oPhrase.sCleDoc)
                    If m_bIndexerChapitre Then _
                        bEcrireChaine(bw, oPhrase.sCodeChapitre)
                    bw.Write(oPhrase.iNumParagrapheL)
                    bw.Write(oPhrase.iNumPhraseG)
                    bw.Write(oPhrase.iNumPhraseL)
                    bEcrireChaine(bw, oPhrase.sPhrase)
                Next oPhrase ' de

            End Using ' bw.Close()
            'End Using ' fs.Close()
            bSauvegarderIndex = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bSauvegarderIndex")
            'Finally
            '    If Not IsNothing(fs) Then fs.Close()
        End Try

Interruption:
        If m_bInterrompre Then
            ' Ne pas conserver un fichier partiel
            If bFichierExiste(sCheminFichierIndex) Then
                bSupprimerFichier(sCheminFichierIndex)
            End If

        End If
        Sablier(bDesactiver:=True)

    End Function

    Private Sub EcrireChapitre(bw As IO.BinaryWriter, oDoc As clsDoc)
        For Each oChap As clsChapitre In oDoc.colChapitres
            bEcrireChaine(bw, oChap.sCle) ' La clé contient toutes les infos.
            bEcrireChaine(bw, oChap.sChapitre)
        Next
    End Sub

    Public Function bLireIndex() As Boolean

        ' Lire l'index VBTxtFinder.idx

        bLireIndex = False
        Dim rVersionFichier As Single
        Dim lNbMots As Integer
        If Not bFichierExiste(m_sCheminVBTxtFndIdx) Then Return False

        If m_bFichierIndexDef Then
            Dim iReponse% = MsgBox("Voulez-vous recharger l'index :" & vbLf &
                m_sCheminVBTxtFndIdx & " ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question,
                sMsgGestionIndex)
            If iReponse = MsgBoxResult.No Then Return False
        End If

        Sablier()
        AfficherMessage("Lecture de l'index en cours...")
        m_bIndexModifie = False
        Dim sMsgErr$ = ""
        Dim bRecommencer As Boolean = False

Recommencer:
        Dim iEncodage% = iCodePageWindowsLatin1252
        If m_bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8

        Dim fs As IO.FileStream = Nothing
        Try

            fs = New IO.FileStream(m_sCheminVBTxtFndIdx, IO.FileMode.Open, IO.FileAccess.Read)
            Using br As New IO.BinaryReader(fs, Encoding.GetEncoding(iEncodage))

                Dim sMsgErrLecture$ = "Version de fichier incorrecte : " & vbLf & m_sCheminVBTxtFndIdx
                rVersionFichier = br.ReadSingle()
                If rVersionFichier <> rVersionFichierVBTxtFndIdx And
           rVersionFichier <> rVersionFichierVBTxtFndIdx10 Then
                    sMsgErr = sMsgErrLecture
                    sMsgErr &= vbLf & "Version = " & rVersionFichier & " <> " &
                rVersionFichierVBTxtFndIdx & " attendu."
                    GoTo Erreur
                End If
                m_bIndexerChapitre = True
                Dim bVersion1p0 As Boolean = False
                If rVersionFichier = rVersionFichierVBTxtFndIdx10 Then
                    bVersion1p0 = True
                    m_bIndexerChapitre = False
                End If

                Dim iNbDocs%, iNb%
                iNbDocs = br.ReadInt32()
                sMsgErrLecture = "Aucun document trouvé dans : " & m_sCheminVBTxtFndIdx
                If iNbDocs = 0 Then GoTo Erreur
                lNbMots = br.ReadInt32()
                sMsgErrLecture = "Auncun mot trouvé dans : " & m_sCheminVBTxtFndIdx
                If lNbMots = 0 Then GoTo Erreur
                Me.iNbPhrasesG = br.ReadInt32()
                sMsgErrLecture = "Aucune phrase trouvée dans : " & m_sCheminVBTxtFndIdx
                If Me.iNbPhrasesG = 0 Then GoTo Erreur
                Me.iNbParagG = br.ReadInt32()
                Me.iNbMotsG = br.ReadInt32()

                ' Non : Place réservée pour compléter les statistiques générales dans une
                '  version future (afin de conserver la compatibilité du fichier index)
                ' Nombre de caractères dont les séparateurs de mot

                ' 22/05/2010 Place utilisée pour l'encodage
                bRecommencer = False
                'Dim iUnicode% = br.ReadInt32()
                Dim iOptionsEncodage% = br.ReadInt32()
                Dim bUnicode As Boolean = (iOptionsEncodage And iMasqueOptionUnicode) > 0
                Dim bAccents As Boolean = (iOptionsEncodage And iMasqueOptionAccent) > 0

                'MsgBox("Encodage unicode = " & iUnicode)

                If bUnicode And Not m_bOptionTexteUnicode Then
                    m_bOptionTexteUnicode = True
                    bRecommencer = True
                ElseIf Not bUnicode And m_bOptionTexteUnicode Then
                    m_bOptionTexteUnicode = False
                    bRecommencer = True
                End If
                If bAccents And Not m_bIndexerAccents Then
                    m_bIndexerAccents = True
                    bRecommencer = True
                ElseIf Not bAccents And m_bIndexerAccents Then
                    m_bIndexerAccents = False
                    bRecommencer = True
                End If
                If bRecommencer Then GoTo FinUsing

                If m_bIndexerChapitre Then m_sbChapitres = New StringBuilder

                Dim sCheminDoc$ = "", sCleDoc$ = "", sCodeDoc$ = ""
                sMsgErrLecture = "Impossible de lire les documents du fichier : " &
                    m_sCheminVBTxtFndIdx & vbLf &
                    "Cause possible : l'encodage ne correspond pas à celui attendu."

                For i = 0 To iNbDocs - 1

                    If Not bLireChaine(br, sCleDoc) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                    If Not bLireChaine(br, sCodeDoc) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                    If Not bLireChaine(br, sCheminDoc) Then sMsgErr = sMsgErrLecture : GoTo Erreur

                    ' 24/05/2019 Voir si on peut trouver l'info. sur Unicode dans le fichier ini
                    ' (il faudrait sauver l'info. dans l'index)
                    Dim bDocUnicode As Boolean = False
                    If Me.m_colDocsIni.Contains(sCodeDoc) Then
                        bDocUnicode = DirectCast(Me.m_colDocsIni(sCodeDoc), clsDoc).bTxtUnicode
                    End If

                    ' Place réservée pour compléter les statistiques par document
                    iNb = br.ReadInt32() ' oDoc.iNbMotsIndexes = iNb
                    iNb = br.ReadInt32() ' oDoc.iNbPhrasesIndexees = iNb
                    iNb = br.ReadInt32() ' oDoc.iNbParagIndexes = iNb
                    iNb = br.ReadInt32() ' oDoc.iNbCarIndexes = iNb
                    ' 17/07/2010 Finalement, le dernier Int32 va servir à indiquer le 
                    '  nombre de chapitres trouvés dans le document
                    Dim iNbChapitresDuDoc% = br.ReadInt32() ' Non : oDoc.iNbCarDontSeparIndexes = iNb

                    Dim colChapitres As Collection = Nothing
                    If m_bIndexerChapitre Then
                        colChapitres = New Collection
                        ' Il faut sauver les chapitres dans la boucle des documents
                        '  car si aucun chapitre, il faut pouvoir l'indiqer :
                        m_sbChapitres.AppendLine(vbCrLf & sCheminDoc & " (" & sCodeDoc & ") :")

                        sMsgErrLecture = "Impossible de lire les chapitres du fichier : " &
                            m_sCheminVBTxtFndIdx
                        For j0 As Integer = 1 To iNbChapitresDuDoc

                            Dim sChapitre$ = ""
                            Dim sCle$ = ""
                            If Not bLireChaine(br, sCle) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                            If Not bLireChaine(br, sChapitre) Then sMsgErr = sMsgErrLecture : GoTo Erreur

                            Dim chap As New clsChapitre
                            chap.sCle = sCle ' sCleDoc & ":" & chap.sCodeChapitre
                            Dim asChamps$() = sCle.Split(":"c)
                            Dim sCleDoc0$ = asChamps(0)
                            Dim sCodeChap0$ = asChamps(1)
                            chap.sCodeChapitre = sCodeChap0
                            chap.sCleDoc = sCleDoc0
                            chap.sCodeDoc = sCodeDoc 'sLireCodeDoc(sCleDoc)
                            chap.sChapitre = sChapitre
                            colChapitres.Add(chap, chap.sCle)

                            m_sbChapitres.AppendLine(sCodeChap0 & " : " & sChapitre)
                        Next
                    End If

                    ' Pour le moment l'info. n'est pas dans l'index (mais dans le fichier ini oui)
                    'Const bTxtUnicode As Boolean = False ' 26/01/2019
                    ' On peut cependant lire l'info. dans le fichier ini (en attendant de le sauver dans l'index)
                    Dim bTxtUnicode = bDocUnicode ' 24/05/2019
                    If Not bAjouterDocument(sCleDoc, sCodeDoc, sCheminDoc, bTxtUnicode, colChapitres) Then
                        sMsgErr = sMsgErrLecture & vbLf & "Impossible d'ajouter le document"
                        GoTo Erreur
                    End If

                Next i

                Dim lPosFinFichier, lPosFichier As Long
                lPosFinFichier = fs.Length

                Dim oMot As clsMot
                Dim sMot$ = ""
                Dim j, iNbPhrases As Integer
                sMsgErrLecture = "Impossible de lire les mots du fichier : " &
                    m_sCheminVBTxtFndIdx
                For i = 0 To lNbMots - 1

                    ' Afficher la progression de la lecture
                    '  Pour les 100 premiers mots, le nombre de phrases peut être elevé
                    If i Mod iModuloAvanvementLent = 0 Or i = lNbMots - 1 Or i < 100 Then
                        lPosFichier = fs.Position
                        Dim rPC! = 100
                        If lPosFinFichier <> 0 Then rPC = CInt(100.0! * lPosFichier / lPosFinFichier) ' 05/05/2018
                        AfficherMessage("Lecture de l'index (mots) en cours... " & rPC & "%")
                        If m_bInterrompre Then GoTo Fin
                    End If

                    ' Lecture du mot
                    If Not bLireChaine(br, sMot) Then sMsgErr = sMsgErrLecture : GoTo Erreur

                    ' Si on récupère un index de la version VB6, faire attention :
                    ' D'abord vérifier rapidement si le mot est indexé tel quel
                    ' (si le mot n'a pas d'accent, cette vérification est rapide)

                    Dim sCle$ = sMot.ToLower
                    Dim sCleAvecAccent$ = sCle
                    Dim sCleSansAccent$ = ""
                    Dim bCleExiste As Boolean = Me.m_htMots.ContainsKey(sCle)

                    ' Si l'index provient de VB6, on ne peut pas savoir
                    '  si le mot est avec ou sans accent
                    ' Optimisation : oublier la compatibilité VB6 pour accélérer le
                    '  chargement de l'index
                    If Not bCleExiste Or Not m_bIndexerAccents Then
                        sCleSansAccent = sEnleverAccents(sMot)
                    End If

                    ' Vérifier si le mot est indexé sans les accents 
                    If Not bCleExiste And Not m_bIndexerAccents Then
                        If String.Compare(sCleSansAccent, sMot) <> 0 Then
                            bCleExiste = Me.m_htMots.ContainsKey(sCleSansAccent)
                            If bCleExiste Then sCle = sCleSansAccent
                        End If
                    End If

                    ' Si la clé sans accent existe déjà, cela signifie qu'il s'agit d'un index 
                    '  en provenance de VB6 (ou DotNet avec les accents et plus maintenant) 
                    '  dans lequel un mot accentué à déjà été ajouté avec une clé sans accent
                    ' Dans ce cas, on fusionne les informations sur les mots,
                    '  pour que la recherche continue à trouver tous les résultats
                    If bCleExiste Then
                        m_bIndexModifie = True
                        oMot = DirectCast(Me.m_htMots.Item(sCle), clsMot)
                        oMot.iNbOccurrences += br.ReadInt32() ' Ajouts des occurences des 2 mots
                        iNbPhrases = br.ReadInt32()
                        ' Nombre de phrases dans lesquelles ce mot figure
                        For j = 1 To iNbPhrases
                            ' Ajout des n° de phrase des 2 mots
                            oMot.AjouterNumPhrase3(br.ReadInt32())
                        Next j
                        Dim sCleMotIndexe$ = oMot.sMot.ToLower

                        ' 03/07/2020 Ssi sCleSansAccent est non vide
                        If Not String.IsNullOrEmpty(sCleSansAccent) AndAlso
                            oMot.sMot.ToLower <> sCleSansAccent Then
                            ' Pour VB6, il suffit de noter le mot lui-même sans les accents
                            '  la clé est déjà sans les accents
                            oMot.sMot = sCleSansAccent
                            If m_bIndexerAccents Then
                                ' Enlever le mot et le réindexer sans les accents
                                Me.m_htMots.Remove(sCle)
                                Me.m_htMots.Add(sCleSansAccent, oMot)
                            End If
                        End If
                    Else
                        oMot = New clsMot
                        oMot.iNbOccurrences = br.ReadInt32()

                        iNbPhrases = br.ReadInt32()

                        ' Nombre de phrases dans lesquelles ce mot figure
                        oMot.RedimPhrases(iNbPhrases)
                        For j = 1 To iNbPhrases
                            oMot.AjouterNumPhrase3(br.ReadInt32())
                        Next j

                        ' Ajouter le mot dans le hastable
                        oMot.sMot = sMot
                        If m_bIndexerAccents Then
                            Me.m_htMots.Add(sCleAvecAccent, oMot)
                        Else
                            Me.m_htMots.Add(sCleSansAccent, oMot)
                        End If
                    End If

                Next i

                sMsgErrLecture = "Impossible de lire les phrases du fichier : " &
                    m_sCheminVBTxtFndIdx
                Dim oPhrase As clsPhrase
                Dim sChaine$ = ""
                Dim iMemNumParagrapheL% = 0
                Dim iNumParagrapheG% = 0
                For i = 0 To Me.iNbPhrasesG - 1

                    ' Afficher la progression de la lecture
                    If i Mod iModuloAvanvementLent = 0 Or i = Me.iNbPhrasesG - 1 Then
                        lPosFichier = fs.Position
                        AfficherMessage("Lecture de l'index (phrases) en cours... " &
                    CInt(100.0! * lPosFichier / lPosFinFichier) & "%")
                        If m_bInterrompre Then GoTo Fin
                    End If

                    oPhrase = New clsPhrase

                    If Not bLireChaine(br, sChaine) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                    oPhrase.sClePhrase = sChaine

                    If Not bLireChaine(br, sChaine) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                    oPhrase.sCleDoc = sChaine

                    oPhrase.sCodeChapitre = ""
                    If Not bVersion1p0 Then
                        If Not bLireChaine(br, sChaine) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                        oPhrase.sCodeChapitre = sChaine
                    End If

                    oPhrase.iNumParagrapheL = br.ReadInt32()
                    ' En déduire le n° de § global
                    If iMemNumParagrapheL <> oPhrase.iNumParagrapheL Then
                        iNumParagrapheG += 1
                    End If
                    oPhrase.iNumParagrapheG = iNumParagrapheG
                    iMemNumParagrapheL = oPhrase.iNumParagrapheL

                    oPhrase.iNumPhraseG = br.ReadInt32()
                    oPhrase.iNumPhraseL = br.ReadInt32()

                    If Not bLireChaine(br, sChaine) Then sMsgErr = sMsgErrLecture : GoTo Erreur
                    oPhrase.sPhrase = sChaine

                    m_colPhrases.Add(oPhrase)
                    Me.iNbPhrasesG = oPhrase.iNumPhraseG

                Next i

                LireListeDocumentsIndexesIni()
                AfficherMessage(sMsgOperationTerminee)
                bLireIndex = True

FinUsing:
            End Using 'br.Close()
            'End Using 'fs.Close()

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bLireIndex", sMsgErr)
            'Finally
            '    If Not IsNothing(fs) Then fs.Close()
        End Try

        If bRecommencer Then GoTo Recommencer

Fin:
        Sablier(bDesactiver:=True)
        Exit Function

Erreur:
        Sablier(bDesactiver:=True)
        'MsgBox(sMsgErr, MsgBoxStyle.Critical, "bLireIndex")
        AfficherMsgErreur("bLireIndex", sMsgErr)

    End Function

#End Region

#Region "Création des documents index sous Word"

    Private Function bInitMotsCourants(sCodeLangIndex$, ByRef sMotsCourants$) As Boolean

        Dim sChemin$ = Application.StartupPath & sCheminMotsCourants & "_" & sCodeLangIndex & sExtTxt
        If sCodeLangIndex = sCodeLangueFr Then
            If bFichierExiste(sChemin) Then
                sMotsCourants = sLireFichier(sChemin)
            Else
                sMotsCourants = Config.sMotsCourantsFr
            End If
        Else
            If Not bFichierExiste(sChemin, bPrompt:=True) Then Return False
            sMotsCourants = sLireFichier(sChemin)
        End If
        If Not m_bIndexerAccents Then sMotsCourants = sEnleverAccents(sMotsCourants)

        Return True

    End Function

    Private Sub ReinitDicoAccentOuPas()

        ' Réindexer les documents avec ou sans les accents
        m_htMots = New Hashtable(m_styleCompare)
        m_colPhrases = New ArrayList
        m_htDico = Nothing

        ' 11/12/2022 Ne pas oublier de réinit. ces compteurs !
        iNbPhrasesG = 0
        iNbMotsG = 0
        iNbParagG = 0

    End Sub

    Public Sub ReinitDico()
        m_htDico = Nothing ' 03/05/2014 Penser à recharger le dico si on change de langue
    End Sub

    Private Function bInitDico(sCheminDico0$) As Boolean

        If Not bFichierExiste(sCheminDico0, bPrompt:=True) Then Return False
        AfficherMessage("Chargement du dictionnaire en cours...")
        m_htDico = CreateCaseInsensitiveHashtable()
        Dim asLignes() = sLireFichier(sCheminDico0).Split(CChar(vbCrLf))
        For Each sLigne0 In asLignes
            Dim sMot = sLigne0.Trim
            Dim sCle = ""
            If m_bIndexerAccents Then
                sCle = sMot.ToLower
            Else
                sCle = sEnleverAccents(sMot)
            End If
            If m_htDico.ContainsKey(sCle) Then
                If m_bIndexerAccents Then Debug.WriteLine("Doublon : " & sCle)
            Else
                m_htDico.Add(sCle, sMot)
            End If
        Next
        Return True

    End Function

    Private Function bMotDico(sMot$) As Boolean

        If IsNothing(m_htDico) Then Return False
        If m_bIndexerAccents Then
            Return m_htDico.ContainsKey(sMot)
        End If
        Dim sMotSansAccent$ = sEnleverAccents(sMot)
        Return m_htDico.ContainsKey(sMotSansAccent)

    End Function

    Private Sub CreerDocIndexSimple(bMotsCourants As Boolean, sCodeLangIndex$,
        bNumeriques As Boolean, bMotsDico As Boolean, sCheminDico0$)

        ' Fabriquer un index simple à partir de la collection de mots indexés

        Dim sMotsCourants$ = ""
        If Not bMotsCourants Then
            If Not bInitMotsCourants(sCodeLangIndex, sMotsCourants) Then Exit Sub
        End If

        If Not bMotsDico AndAlso IsNothing(m_htDico) Then
            If Not bInitDico(sCheminDico0) Then Exit Sub
        End If

        Dim sCheminTxt = m_sCheminDossierCourant & "\" &
            sPrefixeIndexSimple & "_" & sCodeLangIndex & sExtTxt

        Dim sb As New StringBuilder
        Dim sl As New SortedList(CaseInsensitiveComparer.Default)
        Dim de As DictionaryEntry

        For Each de In Me.m_htMots
            Dim oMot As clsMot = DirectCast(de.Value, clsMot)

            Dim sCleMot$ = DirectCast(de.Key, String)
            If m_bIndexerAccents Then
                sCleMot = sCleMot.ToLower
            Else
                ' Enlever les accents comme pour la liste des mots courants
                sCleMot = sEnleverAccents(sCleMot)
            End If
            If Not bMotsCourants AndAlso InStr(sMotsCourants, " " & sCleMot & " ") > 0 Then
                Continue For
            End If

            If Not bMotsDico AndAlso bMotDico(oMot.sMot) Then Continue For

            Dim sMotGlossaire$ = oMot.sMot
            Dim sCle$ = sMotGlossaire

            If Not bNumeriques Then
                ' Exclusion des numériques
                If IsNumeric(sCle) Then Continue For
            End If

            If Not sl.Contains(sCle) Then sl.Add(sCle, sMotGlossaire)

MotSuivant:
        Next de 'oMot

        For Each de In sl
            Dim sLigne$ = DirectCast(de.Value, String)
            sb.Append(sLigne).Append(vbCrLf)
        Next de

        If Not bEcrireFichier(sCheminTxt, sb) Then Exit Sub
        ProposerOuvrirFichier(sCheminTxt)

    End Sub

    Public Sub ComparerIndexSimple(sCodesLanguesIndex$)

        ' Fabriquer une liste de mots communs à deux index simples
        '  par ex. index fr et anglais du même texte : extraction des mots propres

        Dim asCodesLangues$() = sCodesLanguesIndex.Split(";".ToCharArray())
        Dim iNbCodesLangues% = 0
        Dim sCodeLangue1$ = ""
        For Each sCodeLangue As String In asCodesLangues
            If sCodeLangue.Length = 0 Then Continue For
            If sCodeLangue1.Length = 0 Then sCodeLangue1 = sCodeLangue
            'MsgBox(sCodeLangue)
            iNbCodesLangues += 1
        Next
        If iNbCodesLangues < 2 Then
            MsgBox("Il faut au moins 2 codes langues dans la liste pour faire une intersection",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If

        Dim iNbIndex% = 0
        Dim sCheminTxt1 = m_sCheminDossierCourant & "\" &
            sPrefixeIndexSimple & "_" & sCodeLangue1 & sExtTxt

        If Not bFichierExiste(sCheminTxt1, bPrompt:=True) Then Exit Sub
        iNbIndex += 1

        ' D'abord charger l'index simple du 1er code langue
        Dim asLignes$() = asLireFichier(sCheminTxt1)
        Dim ht As Hashtable
        ht = CreateCaseInsensitiveHashtable()
        Dim htNbLang As Hashtable ' Compter le nombre de langues trouvées
        htNbLang = CreateCaseInsensitiveHashtable()
        Dim htLang As Hashtable ' Liste des langues trouvées
        htLang = CreateCaseInsensitiveHashtable()
        For Each sLigne As String In asLignes
            Dim sMot$ = sLigne
            If htNbLang.ContainsKey(sMot) Then
                ' On parcours le 1er index : on ne passe jamais ici car pas de doublon
                Dim iNbLang% = DirectCast(htNbLang(sMot), Integer)
                htNbLang(sMot) = (iNbLang + 1)
            Else
                htNbLang.Add(sMot, 1%)
                htLang.Add(sMot, sCodeLangue1)
            End If
            If ht.ContainsKey(sMot) Then Continue For
            ht.Add(sMot, sMot)
        Next

        ' Ensuite parcourir tous les codes langues présents dans la liste
        For Each sCodeLangue As String In asCodesLangues

            If sCodeLangue = sCodeLangue1 Then Continue For

            Dim sCheminTxt2 = m_sCheminDossierCourant &
                "\" & sPrefixeIndexSimple & "_" & sCodeLangue & sExtTxt
            If Not bFichierExiste(sCheminTxt2) Then Continue For

            ' Ensuite faire l'intersection de l'index simple du 1er code langue 
            '  avec ceux qui existent dans les autres langues
            iNbIndex += 1

            Dim sb As New StringBuilder
            asLignes = asLireFichier(sCheminTxt2)
            For Each sLigne As String In asLignes
                Dim sMot$ = sLigne
                If Not ht.ContainsKey(sMot) Then Continue For
                sb.Append(sMot).Append(vbCrLf) ' Mot commun aux 2 index

                If htNbLang.ContainsKey(sMot) Then
                    Dim iNbLang% = DirectCast(htNbLang(sMot), Integer)
                    htNbLang(sMot) = (iNbLang + 1)
                    Dim sLang$ = DirectCast(htLang(sMot), String)
                    htLang(sMot) = sLang & ";" & sCodeLangue
                Else
                    htNbLang.Add(sMot, 1%)
                    htLang.Add(sMot, sCodeLangue)
                End If

            Next sLigne
            Dim sCheminTxt = m_sCheminDossierCourant & "\" &
                sPrefixeIndexSimple & "_" &
                sCodeLangue1 & "_" & sCodeLangue & sExtTxt
            If Not bEcrireFichier(sCheminTxt, sb) Then Exit Sub
            ProposerOuvrirFichier(sCheminTxt)

        Next sCodeLangue

        If iNbIndex < 2 Then
            MsgBox("Il faut au moins 2 index simples dans 2 codes langues de la liste pour faire une intersection",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If

        If iNbIndex >= 3 Then
            ' Afficher le nombre de langues trouvées par mots communs (à au moins 2 langues)
            Dim sb As New StringBuilder
            Dim ht2 As New Hashtable
            Dim sbDetail As New StringBuilder
            For Each de As DictionaryEntry In htNbLang
                Dim iNbLang% = DirectCast(de.Value, Integer)
                If iNbLang < 2 Then Continue For
                Dim sMot$ = DirectCast(de.Key, String)
                ' Dans l'index final ajouter tous les mots communs à tous les index (chaque langue)
                If iNbLang = iNbIndex Then ht2.Add(sMot, sMot) ' sb.Append(sMot).Append(vbCrLf)
                Dim sLang$ = DirectCast(htLang(sMot), String)
                sbDetail.Append(sMot & ":" & iNbLang & ":" & sLang).Append(vbCrLf)
            Next
            Dim sl As New SortedList(ht2)
            For i As Integer = 0 To sl.Count - 1
                Dim sMot$ = DirectCast(sl.GetByIndex(i), String)
                sb.Append(sMot).Append(vbCrLf)
            Next
            Dim sCheminTxt = m_sCheminDossierCourant & "\" &
                sPrefixeIndexSimple & sExtTxt
            If Not bEcrireFichier(sCheminTxt, sb) Then Exit Sub
            ProposerOuvrirFichier(sCheminTxt)
            sCheminTxt = m_sCheminDossierCourant & "\" &
                sPrefixeIndexSimple & "_Detail" & sExtTxt
            If Not bEcrireFichier(sCheminTxt, sbDetail) Then Exit Sub
            'ProposerOuvrirFichier(sCheminTxt)
        End If

    End Sub

    Public Sub CreerDocIndexMajuscules()

        ' Lister les majuscules intempestives

        Dim sCheminHtml = m_sCheminDossierCourant & "\" & sPrefixeMajuscules & sExtHtm

        Dim sb As New StringBuilder
        Const sEnteteHtml$ = "<html><body>"
        Const sPiedHtml$ = "</body></html>"
        Const sSautLigneHtml$ = "<br>" & vbCrLf
        sb.Append(sEnteteHtml)
        sb.Append("<style type=" & sGm & "text/css" & sGm &
            ">SPAN.Jaune { BACKGROUND-COLOR: yellow }</style>")

        ' Mettre en couleur les majuscules intempestives trouvées dans le document
        Const sBaliseOuv$ = "<SPAN class='Jaune'>"
        Const sBaliseFerm$ = "</SPAN>"

        Dim bAuMoins1Maj As Boolean = False ' intempestive
        For Each oPhrase As clsPhrase In Me.m_colPhrases
            Dim sPhrase$ = oPhrase.sPhrase.Trim

            Dim sbPhrase As New StringBuilder
            Dim bAuMoins1MajPhrase = False ' intempestive
            Dim iMemPosDebOcc% = 0
            Dim iDebRechOcc% = 0
            Dim iLong% = 0
            Dim iMemLong% = 0
            Dim bMemOccIntempest As Boolean = False
            Do
                Dim iPosDebOcc% = sPhrase.IndexOfUppercase(iDebRechOcc)
                If iPosDebOcc = -1 Then Exit Do
                If iPosDebOcc = 0 Then
                    ' Majuscule en début de phrase : normal
                    iDebRechOcc = iPosDebOcc + 1
                    Dim sPortionAv1$ = sPhrase.Substring(0, 1)
                    sbPhrase.Append(sPortionAv1)
                    iLong = 1
                    bMemOccIntempest = False
                    GoTo Suite
                End If

                bAuMoins1Maj = True
                bAuMoins1MajPhrase = True
                Dim iLongPortionAv% = iPosDebOcc - iMemPosDebOcc - iMemLong
                Dim sPortionAv$ = sPhrase.Substring(iDebRechOcc, iLongPortionAv)
                sbPhrase.Append(sPortionAv)

                iLong = 1
                Dim sOccurrence$ = sPhrase.Substring(iPosDebOcc, 1)
                sbPhrase.Append(sBaliseOuv & sOccurrence & sBaliseFerm)
                iDebRechOcc = iPosDebOcc + iLong
                bMemOccIntempest = True

Suite:
                iMemPosDebOcc = iPosDebOcc
                iMemLong = iLong

            Loop While True
            If bAuMoins1MajPhrase Then
                Dim sFin$ = sPhrase.Substring(iMemPosDebOcc + iLong)
                sbPhrase.Append(sFin)
                Dim sAjout$ = sbPhrase.ToString
                sb.Append(sAjout)
                sb.Append(sSautLigneHtml)
            End If

        Next

        If Not bAuMoins1Maj Then
            MsgBox("Aucune majuscule intempestive trouvée dans ce document !",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If
        sb.Append(sPiedHtml)

        ' 26/10/2019 Tous les documents html doivent être en UTF8 (ça doit être l'encodage html par défaut)
        If Not bEcrireFichier(sCheminHtml, sb, bEncodageUTF8:=True) Then Exit Sub
        ProposerOuvrirFichier(sCheminHtml)

    End Sub

    Public Sub CreerDocIndexEspInsec(bTous As Boolean)

        ' Lister les espaces insécables à vérifier ou bien tous les espaces insécables

        Dim sCheminHtml = m_sCheminDossierCourant & "\" &
            sPrefixeEspacesInsecables & sExtHtm

        Dim sb As New StringBuilder
        Const sEnteteHtml$ = "<html><body>"
        Const sPiedHtml$ = "</body></html>"
        Const sSautLigneHtml$ = "<br>" & vbCrLf
        sb.Append(sEnteteHtml)
        sb.Append("<style type=" & sGm & "text/css" & sGm &
            ">SPAN.Jaune { BACKGROUND-COLOR: yellow }</style>")

        ' Mettre en couleur les espaces insécables trouvés dans le document
        Const sBaliseOuv$ = "<SPAN class='Jaune'>"
        Const sBaliseFerm$ = "</SPAN>"
        Dim cCarEspaceInsec As Char = Chr(iCodeASCIIEspaceInsecable)
        Dim sListeCarPrecedOk$ = "«—"
        Dim sListeCarSuivOk$ = "»:;?!%"

        Dim bAuMoins1EspInsec As Boolean = False
        Dim bAuMoins1EspInsecAVerif As Boolean = False
        For Each oPhrase As clsPhrase In Me.m_colPhrases
            If oPhrase.sPhrase.IndexOf(cCarEspaceInsec) = -1 Then Continue For
            Dim sPhrase$ = oPhrase.sPhrase.Trim

            Dim sbPhrase As New StringBuilder
            Dim bAuMoins1EspInsecAVerifPhrase = False
            Dim iMemPosDebOcc% = 0
            Dim iDebRechOcc% = 0
            Dim iLong% = 0
            Dim iMemLong% = 0
            Dim bMemOccAVerifier As Boolean = False ' 19/03/2016
            Do
                Dim iPosDebOcc% = sPhrase.IndexOf(cCarEspaceInsec, iDebRechOcc)
                If iPosDebOcc = -1 Then Exit Do

                ' On ne peut pas surligner un espace après un espace en html ?
                Dim sSoulignerEspaceAvInsec$ = " "

                ' 19/05/2019 Possibilité d'afficher tous les espaces insécables
                If Not bTous AndAlso iPosDebOcc >= 0 Then ' 19/03/2016 1->0
                    ' Vérifier le car. précédant
                    Dim sCarPreced$ = sPhrase.Substring(iPosDebOcc - 1, 1)
                    If sListeCarPrecedOk.Contains(sCarPreced) Then
                        iDebRechOcc = iPosDebOcc + 1
                        Dim iDec% = 0
                        If bMemOccAVerifier Then iDec = 1
                        Dim iLongPortionAv1% = iDebRechOcc - iMemPosDebOcc - iDec
                        Dim sPortionAv1$ = sPhrase.Substring(iMemPosDebOcc + iDec, iLongPortionAv1)
                        sbPhrase.Append(sPortionAv1)
                        iLong = 1
                        bMemOccAVerifier = False
                        GoTo Suite
                    End If
                    bAuMoins1EspInsec = True
                    If sCarPreced = " " Then sSoulignerEspaceAvInsec = "_"
                End If

                ' 19/05/2019 Possibilité d'afficher tous les espaces insécables
                If Not bTous AndAlso iPosDebOcc < sPhrase.Length Then
                    ' Vérifier le car. suivant
                    Dim sCarSuivant$ = sPhrase.Substring(iPosDebOcc + 1, 1)
                    If sListeCarSuivOk.Contains(sCarSuivant) Then
                        iDebRechOcc = iPosDebOcc + 1
                        Dim iLongPortionAv1% = iDebRechOcc - iMemPosDebOcc - 1
                        Dim sPortionAv1$ = sPhrase.Substring(iMemPosDebOcc + 1, iLongPortionAv1)
                        sbPhrase.Append(sPortionAv1)
                        iLong = 1
                        bMemOccAVerifier = False
                        GoTo Suite
                    End If
                    bAuMoins1EspInsec = True
                End If

                bAuMoins1EspInsec = True
                bAuMoins1EspInsecAVerif = True
                bAuMoins1EspInsecAVerifPhrase = True
                Dim iLongPortionAv% = iPosDebOcc - iMemPosDebOcc - iMemLong
                Dim sPortionAv$ = sPhrase.Substring(iDebRechOcc, iLongPortionAv)
                sbPhrase.Append(sPortionAv)

                iLong = 1
                Dim sOccurrence$ = sSoulignerEspaceAvInsec
                sbPhrase.Append(sBaliseOuv & sOccurrence & sBaliseFerm)
                iDebRechOcc = iPosDebOcc + iLong
                ' Occurrence à vérifier : non conforme à sListeCarPrecedOk et sListeCarSuivOk
                bMemOccAVerifier = True

Suite:
                iMemPosDebOcc = iPosDebOcc
                iMemLong = iLong

            Loop While True
            If bAuMoins1EspInsecAVerifPhrase Then
                Dim sFin$ = sPhrase.Substring(iMemPosDebOcc + iLong)
                sbPhrase.Append(sFin)
                Dim sAjout$ = sbPhrase.ToString
                sb.Append(sAjout)
                sb.Append(sSautLigneHtml)
            End If

        Next

        If Not bTous AndAlso Not bAuMoins1EspInsecAVerif Then
            MsgBox("Aucun espace insécable à vérifier trouvé dans ce document !",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If
        If bTous AndAlso Not bAuMoins1EspInsec Then
            MsgBox("Aucun espace insécable trouvé dans ce document !",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If
        sb.Append(sPiedHtml)

        ' 26/10/2019 Tous les documents html doivent être en UTF8 (ça doit être l'encodage html par défaut)
        If Not bEcrireFichier(sCheminHtml, sb, bEncodageUTF8:=True) Then Exit Sub
        ProposerOuvrirFichier(sCheminHtml)

    End Sub

    Public Sub CreerDocIndexCitations()

        ' Fabriquer un index des citations

        Dim sCheminTxt = m_sCheminDossierCourant & "\" &
            sPrefixeIndexCitations & sExtTxt

        Dim sb As New StringBuilder

        Dim sSepGm$ = Chr(iCodeASCIIGuillemet)
        Dim sSepQuote$ = Chr(iCodeASCIIQuote)
        Dim sSepGmO$ = Chr(iCodeASCIIGuillemetOuvrant)
        Dim sSepGmF$ = Chr(iCodeASCIIGuillemetFermant)
        Dim sSepGmO2$ = Chr(iCodeASCIIGuillemetOuvrant2)
        Dim sSepGmF2$ = Chr(iCodeASCIIGuillemetFermant2)
        Dim sSepGmO3$ = Chr(iCodeASCIIGuillemetOuvrant3)
        Dim sSepGmF3$ = Chr(iCodeASCIIGuillemetFermant3)
        Dim sSepGmO4$ = Chr(iCodeASCIIGuillemetOuvrant4)
        Dim sSepGmF4$ = Chr(iCodeASCIIGuillemetFermant4)
        Dim sSepGmO5$ = Chr(iCodeASCIIGuillemetOuvrant5)
        Dim sSepGmF5$ = Chr(iCodeASCIIGuillemetFermant5)
        Dim sListeSepCitations$ = sSepGm & sSepQuote &
            sSepGmO & sSepGmF & sSepGmO2 & sSepGmF2 & sSepGmO3 & sSepGmF3 &
            sSepGmO4 & sSepGmF4 & sSepGmO5 & sSepGmF5

        Dim acSepGm = sSepGm.ToCharArray
        Dim acSepQuote = sSepQuote.ToCharArray
        'Dim acSepGmO = sSepGmO.ToCharArray
        'Dim acSepGmF = sSepGmF.ToCharArray

        Dim bAuMoins1CitationTot As Boolean = False
        For Each oPhrase As clsPhrase In Me.m_colPhrases
            If oPhrase.sPhrase.IndexOfAny(sListeSepCitations.ToCharArray) = -1 Then _
                Continue For
            Dim sPhrase$ = oPhrase.sPhrase.Trim

            ' Voir si indicateur citation présent au moins 2 fois dans la phrase
            ExtraireCitations(sPhrase, sSepGmO, sSepGmF, sb, bAuMoins1CitationTot)
            ExtraireCitations(sPhrase, sSepGmO2, sSepGmF2, sb, bAuMoins1CitationTot)
            ExtraireCitations(sPhrase, sSepGmO3, sSepGmF3, sb, bAuMoins1CitationTot)
            ExtraireCitations(sPhrase, sSepGmO4, sSepGmF4, sb, bAuMoins1CitationTot)
            ExtraireCitations(sPhrase, sSepGmO5, sSepGmF5, sb, bAuMoins1CitationTot)

            Dim iLen% = sPhrase.Length
            Dim bAuMoins1Citation As Boolean = False
            Dim sbTmp As StringBuilder

            Dim iLenSansGm = sPhrase.Replace(sSepGm, "").Length
            Dim iNbGm% = iLen - iLenSansGm
            If iNbGm <= 1 Then GoTo Suite2
            bAuMoins1Citation = False
            sbTmp = New StringBuilder
            sbTmp.Append(sPhrase).Append(vbCrLf)
            If iNbGm = 2 Then
                Dim iPosGm1% = sPhrase.IndexOf(acSepGm)
                Dim iPosGm2% = sPhrase.LastIndexOf(acSepGm)
                Dim sCitation$ = sPhrase.Substring(iPosGm1 + 1, iPosGm2 - iPosGm1 - 1).Trim
                If sCitation.Length > 0 Then
                    bAuMoins1Citation = True
                    sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
                End If
            Else
                Dim iPos% = 0
                Do
                    Dim iPosGm1% = sPhrase.IndexOf(acSepGm, iPos)
                    If iPosGm1 = -1 Then Exit Do
                    Dim iPosGm2% = sPhrase.IndexOf(acSepGm, iPosGm1 + 1)
                    If iPosGm2 = -1 Then Exit Do
                    Dim sCitation$ = sPhrase.Substring(iPosGm1 + 1, iPosGm2 - iPosGm1 - 1).Trim
                    If sCitation.Length > 0 Then
                        bAuMoins1Citation = True
                        sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
                    End If
                    iPos = iPosGm2 + 1
                Loop While True
            End If
            If bAuMoins1Citation Then sb.Append(sbTmp) : bAuMoins1CitationTot = True

Suite2:
            ' Vérifier que le caractère suivant n'est pas 's 'd 't 'm 'll 've
            Const sListeExcepAnglais$ = "stdmlv"
            Dim iLenSansQuotes = sPhrase.Replace(sSepQuote, "").Length
            Dim iNbQuotes% = iLen - iLenSansQuotes
            If iNbQuotes <= 1 Then GoTo Suite3
            bAuMoins1Citation = False
            sbTmp = New StringBuilder
            sbTmp.Append(sPhrase).Append(vbCrLf)
            If iNbQuotes = 2 Then
                Dim iPos1% = sPhrase.IndexOf(acSepQuote)
                Dim iPos2% = sPhrase.LastIndexOf(acSepQuote)
                ' Pour les quotes, la citation doit commencer par la quote
                If iPos1 <> 0 Then GoTo Suite3
                If iPos1 < iLen - 1 Then
                    Dim sCarSuiv$ = sPhrase.Substring(iPos1 + 1, 1)
                    If sCarSuiv.IndexOfAny(sListeExcepAnglais.ToCharArray) > -1 Then _
                        GoTo Suite3
                End If
                If iPos2 < iLen - 1 Then
                    Dim sCarSuiv$ = sPhrase.Substring(iPos2 + 1, 1)
                    If sCarSuiv.IndexOfAny(sListeExcepAnglais.ToCharArray) > -1 Then _
                        GoTo Suite3
                End If
                Dim sCitation$ = sPhrase.Substring(iPos1 + 1, iPos2 - iPos1 - 1).Trim
                If sCitation.Length > 0 Then
                    bAuMoins1Citation = True
                    sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
                End If
            Else
                Dim iPos% = 0
                Dim iPos2% = -1
                Do
                    Dim iPos1% = sPhrase.IndexOf(acSepQuote, iPos)
                    If iPos1 = -1 Then Exit Do

                    ' Pour les quotes, la citation doit commencer par la quote
                    'If iPos1 <> 0 Then Continue For
                    ' Pour la 1ère citation seulement
                    If iPos1 <> 0 And iPos = 0 Then Exit Do

                    If iPos1 < iLen - 1 Then
                        Dim sCarSuiv$ = sPhrase.Substring(iPos1 + 1, 1)
                        If sCarSuiv.IndexOfAny(sListeExcepAnglais.ToCharArray) > -1 Then _
                            Exit Do
                    End If
                    iPos2 = sPhrase.IndexOf(acSepQuote, iPos1 + 1)
                    If iPos2 = -1 Then Exit Do
                    If iPos2 < iLen - 1 Then
                        Dim sCarSuiv$ = sPhrase.Substring(iPos2 + 1, 1)
                        If sCarSuiv.IndexOfAny(sListeExcepAnglais.ToCharArray) > -1 Then _
                            Exit Do
                    End If
                    Dim sCitation$ = sPhrase.Substring(iPos1 + 1, iPos2 - iPos1 - 1).Trim
                    If sCitation.Length > 0 Then
                        bAuMoins1Citation = True
                        sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
                    End If
                    iPos = iPos2 + 1
                Loop While True
            End If
            If bAuMoins1Citation Then sb.Append(sbTmp) : bAuMoins1CitationTot = True

Suite3:

            ' 19/01/2019 S'il y a juste " alors lister aussi
            If Not bAuMoins1Citation AndAlso iNbGm = 1 AndAlso sPhrase.Length > 1 Then
                Dim iPosGm% = sPhrase.IndexOf(acSepGm)
                If iPosGm > -1 AndAlso sPhrase.Length - iPosGm > 1 Then
                    Dim sCitation = sPhrase.Substring(iPosGm, sPhrase.Length - iPosGm)
                    sb.AppendLine(sCitation)
                    bAuMoins1CitationTot = True
                End If
            End If

        Next

        If Not bAuMoins1CitationTot Then
            MsgBox("Aucune citation trouvée dans ce document !",
                MsgBoxStyle.Information, sTitreMsg)
            Exit Sub
        End If
        If Not bEcrireFichier(sCheminTxt, sb) Then Exit Sub
        ProposerOuvrirFichier(sCheminTxt)

    End Sub

    Private Sub ExtraireCitations(sPhrase$, sSepGmO$, sSepGmF$,
        sb As StringBuilder, ByRef bAuMoins1CitationTot As Boolean)

        ' Extraire les citations entre une paire de guillemets ouvrant et fermant distincts

        Dim iLen% = sPhrase.Length
        Dim acSepGmO = sSepGmO.ToCharArray
        Dim acSepGmF = sSepGmF.ToCharArray

        Dim iLenSansGmO = sPhrase.Replace(sSepGmO, "").Length
        Dim iLenSansGmF = sPhrase.Replace(sSepGmF, "").Length
        Dim iNbGmO% = iLen - iLenSansGmO
        Dim iNbGmF% = iLen - iLenSansGmF
        If iNbGmO = 0 Or iNbGmF = 0 Then Exit Sub
        Dim bAuMoins1Citation = False
        Dim sbTmp As New StringBuilder
        sbTmp.Append(sPhrase).Append(vbCrLf)
        If iNbGmO = 1 And iNbGmF = 1 Then
            Dim iPosGm1% = sPhrase.IndexOf(acSepGmO)
            Dim iPosGm2% = sPhrase.LastIndexOf(acSepGmF)
            If iPosGm2 <= iPosGm1 Then Exit Sub
            Dim sCitation$ = sPhrase.Substring(iPosGm1 + 1, iPosGm2 - iPosGm1 - 1).Trim
            If sCitation.Length > 0 Then
                bAuMoins1Citation = True
                sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
            End If
        Else
            Dim iPos% = 0
            Do
                Dim iPosGm1% = sPhrase.IndexOf(acSepGmO, iPos)
                If iPosGm1 = -1 Then Exit Do
                Dim iPosGm2% = sPhrase.IndexOf(acSepGmF, iPosGm1 + 1)
                If iPosGm2 = -1 Then Exit Do
                If iPosGm2 > iPosGm1 Then
                    Dim sCitation$ = sPhrase.Substring(iPosGm1 + 1, iPosGm2 - iPosGm1 - 1).Trim
                    If sCitation.Length > 0 Then
                        bAuMoins1Citation = True
                        sbTmp.Append(vbTab & sCitation).Append(vbCrLf)
                    End If
                End If
                iPos = iPosGm2 + 1
            Loop While True
        End If
        If bAuMoins1Citation Then sb.Append(sbTmp) : bAuMoins1CitationTot = True

    End Sub

    Public Sub CreerDocIndex(sTypeIndex$, bMotsDico As Boolean,
        bMotsCourants As Boolean, sCheminDico0$, sCodeLangIndex$,
        bMotsSeulsDocIndex0 As Boolean, iMaxMotsCles%,
        bNumeriques As Boolean, sCodesLanguesIndex$,
        Optional bCreerDocWord As Boolean = True,
        Optional bProposerOuvrir As Boolean = True,
        Optional ByRef sCheminFinal$ = "")

        ' Fabriquer un index à partir de la collection de mots indexés

        If sTypeIndex = sIndexMajuscules Then
            CreerDocIndexMajuscules()
            Exit Sub
        End If
        If sTypeIndex = sIndexEspacesInsecables Then
            CreerDocIndexEspInsec(bTous:=True)
            Exit Sub
        End If
        If sTypeIndex = sIndexEspacesInsecablesAVerifier Then
            CreerDocIndexEspInsec(bTous:=False)
            Exit Sub
        End If
        If sTypeIndex = sIndexCitations Then
            CreerDocIndexCitations()
            Exit Sub
        End If
        If sTypeIndex = sIndexSimple Then
            CreerDocIndexSimple(bMotsCourants, sCodeLangIndex, bNumeriques,
                bMotsDico, sCheminDico0)
            Exit Sub
        End If
        If sTypeIndex = sIndexSimpleComparer Then
            ComparerIndexSimple(sCodesLanguesIndex)
            Exit Sub
        End If

        If sTypeIndex = sIndexTout Then
            RegenererDocs()
            Exit Sub
        End If

        If sTypeIndex = sIndexAccents Then
            AnalyseAccents(sCodeLangIndex, sCodesLanguesIndex)
            Exit Sub
        End If

        If Not bCreerDocIndexIntern(sTypeIndex, bMotsDico,
            bMotsCourants, sCheminDico0, sCodeLangIndex,
            bMotsSeulsDocIndex0, iMaxMotsCles,
            bNumeriques, sCodesLanguesIndex,
            bCreerDocWord, bProposerOuvrir, sCheminFinal) Then Exit Sub

    End Sub

    Private Function bCreerDocIndexIntern(sTypeIndex$, bMotsDico As Boolean,
        bMotsCourants As Boolean, sCheminDico0$, sCodeLangIndex$,
        bMotsSeulsDocIndex0 As Boolean, iMaxMotsCles%,
        bNumeriques As Boolean, sCodesLanguesIndex$,
        Optional bCreerDocWord As Boolean = True,
        Optional bProposerOuvrir As Boolean = True,
        Optional ByRef sCheminFinal$ = "") As Boolean

        If Not bMotsDico AndAlso IsNothing(m_htDico) Then
            If Not bInitDico(sCheminDico0) Then Return False
        End If

        ' Mots seuls sans afficher les occurrences
        Dim bMotsCles As Boolean
        Dim bMotsSeulsDocIndex As Boolean = bMotsSeulsDocIndex0 '  Config.bMotsSeulsDocIndex
        If sTypeIndex = sIndexMotsCles Then bMotsCles = True : bMotsSeulsDocIndex = True

        Dim sMotsCourants$ = Config.sMotsCourantsFr
        If Not bMotsCourants Then
            If Not bInitMotsCourants(sCodeLangIndex, sMotsCourants) Then Return False
        End If
        If Not m_bIndexerAccents Then sMotsCourants = sEnleverAccents(sMotsCourants)

        Dim bTriFreq As Boolean
        If sTypeIndex <> sIndexAlpha Then bTriFreq = True

        Dim sTitre$, sListeMax$
        Dim sExplication$ = ""
        Dim iNbDocIndexes% = Me.m_colDocs.Count()

        Dim sAccent$ = ""
        If Me.m_bIndexerAccents Then sAccent = "avec les accents " ' 06/06/2019

        sTitre = "Document index " & sAccent & "de VBTextFinder"
        If Not bMotsDico And Not bMotsCourants Then
            sTitre = "Document index " & sAccent & "(hors mots du dictionnaire et mots courants " &
                sCodeLangIndex & ") de VBTextFinder"
        ElseIf Not bMotsDico Then
            sTitre = "Document index " & sAccent & "(hors mots du dictionnaire " &
                sCodeLangIndex & ") de VBTextFinder" ' 16/01/2011 Manque une ) ici
        ElseIf Not bMotsCourants Then
            sTitre = "Document index " & sAccent & "(hors mots courants " &
                sCodeLangIndex & ") de VBTextFinder"
        End If

        sListeMax = ""
        If Not bMotsCles Then sListeMax =
            "liste des codes documents - " & iNbOccurrencesMaxListe &
            " au max. - pour les mots non fréquents <= " &
            iNbOccurencesMaxRecherchees & " occurrences"
        If iNbDocIndexes = 1 Then sListeMax = ""

        If bTriFreq Then
            sTitre &= " trié en fréquence"
            If bMotsSeulsDocIndex Then sTitre &=
                " sans les mots courants : mots clés"
            If Not bMotsSeulsDocIndex Then _
                sExplication = "Explication : Nombre d'occurrences : Mot"
            If sListeMax <> "" Then sExplication &= " (" & sListeMax & ")"
        Else
            sTitre &= " trié par ordre alphabétique"
            If Not bMotsSeulsDocIndex Then _
                sExplication = "Explication : Mot (nombre d'occurrences"
            If sListeMax <> "" Then sExplication &= " : " & sListeMax
            If Not bMotsSeulsDocIndex Then sExplication &= ")"
        End If
        sExplication = sExplication & vbLf

        ' Si le fichier existe, le supprimer avant
        Dim sCheminTxt = m_sCheminDossierCourant & "\" & sFichierVBTxtFndAlphab & sExtTxt
        If bTriFreq Then sCheminTxt = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndFreq & sExtTxt
        If bMotsCles Then sCheminTxt = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndMotsCles & sExtTxt

        Dim sCheminDoc = m_sCheminDossierCourant & "\" & sFichierVBTxtFndAlphab & sExtDoc
        If bTriFreq Then sCheminDoc = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndFreq & sExtDoc
        If bMotsCles Then sCheminDoc = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndMotsCles & sExtDoc

        If Not bFichierAccessible(sCheminTxt, bPromptFermer:=True,
            bInexistOk:=True, bPromptRetenter:=True) Then Return False
        If Not bFichierAccessible(sCheminDoc, bPromptFermer:=True,
            bInexistOk:=True, bPromptRetenter:=True) Then Return False

        ' 09/06/2019
        Dim sCheminExclusions$ = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndAlphab & "_Exclusions.txt"
        Dim sCheminInclusions$ = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndAlphab & "_Inclusions.txt"
        If Not bFichierAccessible(sCheminExclusions, bPromptFermer:=True,
            bInexistOk:=True, bPromptRetenter:=True) Then Return False
        If Not bFichierAccessible(sCheminInclusions, bPromptFermer:=True,
            bInexistOk:=True, bPromptRetenter:=True) Then Return False
        Dim bExclusions As Boolean = False
        Dim hsExcl As HashSet(Of String) = Nothing
        If bFichierExiste(sCheminExclusions) Then
            Dim asExcl = asLireFichier(sCheminExclusions, bLectureSeule:=True) ' bUnicodeUTF8:=True)
            Dim lstExcl = asExcl.ToList
            If Not m_bIndexerAccents Then lstExcl = ListToListSansAccent(lstExcl)
            If Not bListToHashSet(lstExcl, hsExcl, bPromptErr:=True) Then Return False
            If hsExcl.Count > 0 Then bExclusions = True
        End If
        Dim bInclusions As Boolean = False
        Dim hsIncl As HashSet(Of String) = Nothing
        If bFichierExiste(sCheminInclusions) Then
            Dim asIncl = asLireFichier(sCheminInclusions, bLectureSeule:=True) ', bUnicodeUTF8:=True)
            Dim lstIncl = asIncl.ToList
            If Not m_bIndexerAccents Then lstIncl = ListToListSansAccent(lstIncl)
            If Not bListToHashSet(lstIncl, hsIncl, bPromptErr:=True) Then Return False
            If hsIncl.Count > 0 Then bInclusions = True
        End If

        m_bInterrompre = False

        Dim sb As New StringBuilder

        Sablier()

        Dim oMot As clsMot
        Dim sMotGlossaire$
        Dim lNumMotIndexe, lNbMotsIndexes As Integer
        lNbMotsIndexes = Me.m_htMots.Count()
        lNumMotIndexe = 0
        Dim iNbOccEffectives%
        Dim sCleDocPhrase$, sListeRef$, sMemCleDocPhrase$
        Dim i, lFin As Integer

        Dim sl As New SortedList(CaseInsensitiveComparer.Default)

        Dim de As DictionaryEntry
        Dim lMaxOcc%, sMaxFreq$, iLenMaxFreq%
        Dim sFormatOcc$ = ""
        If bTriFreq Then

            AfficherMessage("Recherche du nombre d'occurrence max. pour la présentation...")

            ' Recherche du nbre d'occ max pour le format de présentation : nbre de 0
            For Each de In Me.m_htMots
                oMot = DirectCast(de.Value, clsMot)
                If Not bMotsDico AndAlso bMotDico(oMot.sMot) Then Continue For
                If oMot.iNbOccurrences > lMaxOcc Then lMaxOcc = oMot.iNbOccurrences
            Next de
            sMaxFreq = CStr(lMaxOcc)
            iLenMaxFreq = sMaxFreq.Length
            For i = 0 To iLenMaxFreq - 1
                sFormatOcc &= "0"
            Next i
            'sFormatOcc = VB6.String(iLenMaxFreq, "0")
        End If

        ' Projet Complexifieur : fabriquer des mots en les dérivant avec des postfixes :
        ' Logiciel de complexificationnage du langage (ou jargonasification)
        '  Complexe -> Complexité -> Complexification -> Complexificationnage -> Complexificationnement...
        '  Rater -> Ratage -> Rature...
        Dim asComplexifieurs$(7)
        If bTestComplexifieur Then
            ' iComplexifieurMinRecherche iComplexifieurMaxRecherche
            asComplexifieurs(3) = Config.sComplexifieurs3
            asComplexifieurs(4) = Config.sComplexifieurs4
            asComplexifieurs(5) = Config.sComplexifieurs5
            'asComplexifieurs(6) = Config.sComplexifieurs6
            'asComplexifieurs(7) = Config.sComplexifieurs7
            If Not m_bIndexerAccents Then
                asComplexifieurs(3) = sEnleverAccents(Config.sComplexifieurs3)
                asComplexifieurs(4) = sEnleverAccents(Config.sComplexifieurs4)
                asComplexifieurs(5) = sEnleverAccents(Config.sComplexifieurs5)
            End If
        End If

        Dim iFreqMin% = Me.m_htMots.Count \ 100

        Dim bBiGramme As Boolean = False
        If sTypeIndex = sIndexNGrammes Then bBiGramme = True
        Dim htBiG As New Hashtable
        'Dim oBG As clsBiGramme

        For Each de In Me.m_htMots
            oMot = DirectCast(de.Value, clsMot)

            'If oMot.sMot = "temps" Then Debug.WriteLine("!")

            If Not bMotsDico AndAlso bMotDico(oMot.sMot) Then
                If Not (bInclusions AndAlso hsIncl.Contains(oMot.sMot)) Then Continue For ' 09/06/2019
                'Debug.WriteLine("!")
                'Continue For
            End If

            Dim sCleMot$ = DirectCast(de.Key, String)

            If m_bIndexerAccents Then ' Conserver les accents
                sCleMot = sCleMot.ToLower
            Else
                ' Note : les accents sont déjà rétirés ici de toutes façons : pas besoin
                ' Enlever les accents comme pour la liste des mots courants
                sCleMot = sEnleverAccents(sCleMot)
            End If

            If bExclusions AndAlso hsExcl.Contains(sCleMot) Then Continue For ' 09/06/2019

            If Not bMotsCourants AndAlso InStr(sMotsCourants, " " & sCleMot & " ") > 0 Then
                If Not (bInclusions AndAlso hsIncl.Contains(sCleMot)) Then Continue For ' 09/06/2019
                'Debug.WriteLine("!")
                'Continue For
            End If

            If bBiGramme Then
                ' Test bigrammes
                Dim sMotBrut$ = DirectCast(de.Key, String)
                TestBiGrammesP1(htBiG, sMotBrut)
                GoTo MotSuivant
            End If

            If bMotsCles Then
                If InStr(sMotsCourants, " " & sCleMot & " ") > 0 Then
                    GoTo MotSuivant
                End If
                If oMot.iNbOccurrences < iFreqMin Then
                    GoTo MotSuivant
                End If
            End If

            lNumMotIndexe += 1
            If lNumMotIndexe Mod iModuloAvanvementLent = 0 Or
                lNumMotIndexe = lNbMotsIndexes Or lNumMotIndexe = 1 Then
                AfficherMessage("Création du document index en cours : " &
                    lNumMotIndexe & " / " & lNbMotsIndexes)
                If m_bInterrompre Then Exit For
            End If

            sListeRef = ""

            ' S'il n'y a qu'un seul document indexé, inutile d'indiquer toujours
            '  la même référence à ce document
            If iNbDocIndexes = 1 Then GoTo Suite

            lFin = oMot.iNbPhrases
            If lFin > iNbOccurencesMaxRecherchees Then GoTo Suite

            sMemCleDocPhrase = ""
            iNbOccEffectives = 0
            Dim iMemNumPhrase% = -1
            For i = 1 To lFin

                Dim iNumPhrase% = oMot.iLireNumPhrase(i)
                If iNumPhrase = iMemNumPhrase Then Continue For
                iMemNumPhrase = iNumPhrase

                Dim sCodeChapitre$ = ""
                sCleDocPhrase = sLireCleDocPhrase(iNumPhrase, sCodeChapitre)
                If sCleDocPhrase = sMemCleDocPhrase Then Continue For
                sMemCleDocPhrase = sCleDocPhrase

                Dim sCodeDoc$ = sLireCodeDoc(sCleDocPhrase)
                If m_bAfficherChapitreIndex AndAlso sCodeChapitre.Length > 0 Then
                    sCodeDoc &= ":" & sCodeChapitre
                End If

                iNbOccEffectives += 1
                If iNbOccEffectives > iNbOccurrencesMaxListe Then _
                    sListeRef &= ", ..." : Exit For
                If i = 1 Then
                    sListeRef = sCodeDoc
                Else
                    sListeRef &= ", " & sCodeDoc
                End If

            Next i

Suite:
            Dim sCle$

            If bTriFreq Then
                ' Tri fréquentiel : on met le nombre d'occurence du mot en premier
                If bMotsSeulsDocIndex Then
                    sMotGlossaire = oMot.sMot
                Else
                    If sListeRef <> "" Then sListeRef = " (" & sListeRef & ")"
                    sMotGlossaire = oMot.iNbOccurrences & " : " & oMot.sMot & sListeRef
                End If

                ' Ne peut pas marcher avec une SortedList car la clé n'est pas unique !
                sCle = Format(oMot.iNbOccurrences, sFormatOcc) & " : " & oMot.sMot & sListeRef

            Else
                ' Tri alphabétique : on met le mot en premier
                If bMotsSeulsDocIndex Then
                    sMotGlossaire = oMot.sMot
                Else
                    If sListeRef <> "" Then sListeRef = " : " & sListeRef
                    sMotGlossaire = oMot.sMot & " (" & oMot.iNbOccurrences & sListeRef & ")"
                End If
                sCle = oMot.sMot

            End If

            If bTestComplexifieur Then
                ' Sélectionner les mots dérivés à partir d'un mot plus simple
                '  en examinant la fin des mots
                Dim bMotDerive As Boolean
                bMotDerive = False
                Dim sMot$ = oMot.sMot
                For i = iComplexifieurMinRecherche To iComplexifieurMaxRecherche
                    If sMot.Length > i Then
                        Dim sFinMot$ = Right$(sMot, i)
                        ' Les mots accentués ne sont pas distingués
                        If Not m_bIndexerAccents Then sFinMot = sEnleverAccents(sFinMot)
                        If InStr(asComplexifieurs(i), " " & sFinMot & " ") > 0 Then _
                            bMotDerive = True : Exit For
                    End If
                Next i
                If Not bMotDerive Then GoTo MotSuivant
            End If

            If Not bNumeriques Then
                ' Exclusion des numériques
                If IsNumeric(sCle) Then Continue For
            End If

            'Try
            If Not sl.Contains(sCle) Then
                sl.Add(sCle, sMotGlossaire)
            Else
                'Catch
                ' S'il y a une erreur, c'est que HashTable est capable de distinguer
                '  Coeur de cœur, mais pas SortedList, car il n'y a pas d'équivalent de
                '  CaseInsensitiveHashCodeProvider.Default pour SortedList
            End If
            'End Try

MotSuivant:
        Next de 'oMot

        If bBiGramme Then
            TestBiGrammesP2(htBiG, sb, sl)
            GoTo Fin0
        End If

        Dim sLigne$
        Dim iNbMots% = sl.Count
        If bTriFreq Then
            For i = iNbMots - 1 To 0 Step -1
                If bMotsCles And i < iNbMots - iMaxMotsCles Then Exit For
                sLigne = DirectCast(sl.GetByIndex(i), String)
                ' C'est Word qui ajoute des sauts de ligne inopinés
                If bMotsCles Or bMotsSeulsDocIndex Then
                    sb.Append(sLigne.ToLower & " ")
                Else
                    sb.Append(sLigne).Append(vbCrLf)
                End If
            Next i
        Else
            For Each de In sl
                sLigne = DirectCast(de.Value, String)
                sb.Append(sLigne).Append(vbCrLf)
            Next de
        End If

Fin0:
        Static bWord As Boolean = True
        Dim iEncodage% = iCodePageWindowsLatin1252
        If m_bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8
        If Not bEcrireFichier(sCheminTxt, sb, iEncodage:=iEncodage) Then _
            Sablier(bDesactiver:=True) : Return False
        ' Si Word n'est pas installé, ne plus essayer de l'ouvrir
        If Not bCreerDocWord OrElse Not bWord Then GoTo Fin
        AfficherMessage("Ouverture de Microsoft Word...")
        If bCreerDocIndex2(sCheminTxt, sCheminDoc, sTitre, sExplication, lNbMotsIndexes,
            Me.m_colDocs, m_bInterrompre, bWord, m_bOptionTexteUnicode, sCodeLangIndex) Then
            AfficherMessage("Création du document index terminée.")
            If bProposerOuvrir Then ProposerOuvrirFichier(sCheminDoc)
        End If

Fin:
        Sablier(bDesactiver:=True)
        If Not bWord AndAlso bProposerOuvrir Then ProposerOuvrirFichier(sCheminTxt)
        sCheminFinal = sCheminTxt

        Return True

    End Function

    Private Function ListToListSansAccent(lst As List(Of String)) As List(Of String)

        Dim lstDest As New List(Of String)
        For Each sMot In lst
            lstDest.Add(sEnleverAccents(sMot))
        Next
        Return lstDest

    End Function

    Private Sub TestBiGrammesP1(htBiG As Hashtable, sMotBrut$)

        ' Test bigrammes
        Dim oBG As clsBiGramme
        sMotBrut = sEnleverAccents(sMotBrut)
        Dim j%
        Dim iLen% = sMotBrut.Length
        Dim iFin% = 2
        iFin = 3 ' Trigrammes
        For j = 0 To iLen - iFin
            Dim sCar1$ = sMotBrut.Chars(j)
            Dim sCar2$ = sMotBrut.Chars(j + 1)
            Dim sCar3$ = sMotBrut.Chars(j + 2)
            If Not Char.IsLetter(sCar1.Chars(0)) Or
                Not Char.IsLetter(sCar2.Chars(0)) Or
                Not Char.IsLetter(sCar3.Chars(0)) Then
                GoTo CarSuiv
            End If
            Dim sBiGramme$ = sCar1 & sCar2 & sCar3
            If htBiG.ContainsKey(sBiGramme) Then
                oBG = CType(htBiG(sBiGramme), clsBiGramme)
                oBG.iNbOccurences += 1
            Else
                oBG = New clsBiGramme
                oBG.sBiGramme = sBiGramme
                oBG.iNbOccurences = 1
                htBiG.Add(sBiGramme, oBG)
            End If
CarSuiv:
        Next j

    End Sub

    Private Sub TestBiGrammesP2(htBiG As Hashtable, sb As StringBuilder, sl As SortedList)

        Dim oBG As clsBiGramme
        Dim sFormatOcc = "0.000%"
        Dim lOccMax& = 0
        Dim lOccTot& = 0
        For Each de As DictionaryEntry In htBiG
            oBG = DirectCast(de.Value, clsBiGramme)
            lOccTot += oBG.iNbOccurences
            If oBG.iNbOccurences > lOccMax Then lOccMax = oBG.iNbOccurences
        Next de
        For Each de As DictionaryEntry In htBiG
            oBG = DirectCast(de.Value, clsBiGramme)
            Dim rFreqTot# = 100
            If lOccTot <> 0 Then rFreqTot = oBG.iNbOccurences / lOccTot ' 05/05/2018
            Dim sFreq$ = Format(rFreqTot, sFormatOcc)
            Dim sBG$ = sFreq & " : " & oBG.sBiGramme
            Dim sCle$ = sFreq & " : " & " : " & oBG.sBiGramme
            Try
                sl.Add(sCle, sBG)
            Catch
            End Try
        Next de
        Dim iNbBG% = sl.Count
        For i = iNbBG - 1 To 0 Step -1
            Dim sLigne0$ = DirectCast(sl.GetByIndex(i), String)
            sb.Append(sLigne0).Append(vbCrLf)
        Next i

    End Sub

    Private Sub RegenererDocs()

        ' Régénérer complètement les documents indexés
        '  (seuls les lignes vides sont supprimées)

        Dim sCheminTxt$ = m_sCheminDossierCourant & "\" &
            sFichierVBTxtFndTout & sExtTxt

        Dim sb As New StringBuilder

        Sablier()

        ' Utiliser le format de présentation en français, 
        '  en utilisant les préférences de l'utilisateur le cas échéant
        Dim nfi As System.Globalization.NumberFormatInfo =
            New System.Globalization.CultureInfo("fr-FR", useUserOverride:=True).NumberFormat
        nfi.NumberDecimalDigits = 0 ' Afficher des nombres entiers, sans virgule
        sb.Append("Nombre de mots indexés : " &
            Me.iNbMotsG.ToString("N", nfi)).Append(vbCrLf)
        sb.Append("Nombre de mots distincts indexés : " &
            Me.m_htMots.Count().ToString("N", nfi)).Append(vbCrLf)
        sb.Append("Nombre de phrases indexées : " &
            m_colPhrases.Count().ToString("N", nfi)).Append(vbCrLf)
        sb.Append("Nombre de paragraphes indexés : " &
            Me.iNbParagG.ToString("N", nfi)).Append(vbCrLf)

        If Me.tsDiffTps.Milliseconds <> 0 Then _
            sb.Append("Temps d'indexation : " & tsDiffTps.ToString).Append(vbCrLf)

        sb.Append(vbCrLf)
        sb.Append("Liste des documents indexés (" & Me.m_colDocs.Count() & ") :").Append(vbCrLf)
        Dim oDoc As clsDoc
        For Each oDoc In Me.m_colDocs
            sb.Append(oDoc.sChemin & " (" & oDoc.sCodeDoc & ")").Append(vbCrLf)
        Next oDoc

        If m_bIndexerChapitre Then
            sb.AppendLine(vbCrLf & "Liste des chapitres :")
            sb.Append(m_sbChapitres)
            ' Identique à m_sbChapitres :
            'For Each oDoc In Me.m_colDocs
            '    sb.AppendLine(vbCrLf & oDoc.sChemin & " (" & oDoc.sCodeDoc & ") :")
            '    For Each chapitre As clsChapitre In oDoc.colChapitres
            '        sb.AppendLine(chapitre.sCodeChapitre & " : " & chapitre.sChapitre)
            '    Next chapitre
            'Next oDoc
        End If

        sb.Append(vbCrLf)
        sb.Append("Liste des phrases indexées :").Append(vbCrLf)

        Dim iMemNumParag% = 0
        For i = 1 To Me.iNbPhrasesG ' Parcours de toutes les phrases
            Dim oPhrase As clsPhrase = DirectCast(m_colPhrases.Item(i - 1), clsPhrase)
            Dim oPhraseSuiv As clsPhrase = oPhrase
            If i < Me.iNbPhrasesG Then _
                oPhraseSuiv = DirectCast(m_colPhrases.Item(i), clsPhrase)

            If i = 1 Then sb.Append(sInfoDoc(oPhrase))
            If oPhrase.iNumParagrapheL <> iMemNumParag And
               oPhrase.iNumParagrapheL = 1 Then
                If bExporterToutAvecNumeros Then sb.Append(sInfoParag(oPhrase))
                iMemNumParag = oPhrase.iNumParagrapheL
            End If

            sb.Append(oPhrase.sPhrase)
            If oPhrase.iNumParagrapheG <> oPhraseSuiv.iNumParagrapheG Then sb.Append(vbCrLf)

            If oPhraseSuiv.sCleDoc <> oPhrase.sCleDoc Then
                sb.Append(vbCrLf).Append(vbCrLf)
                sb.Append(sInfoDoc(oPhraseSuiv))
            End If
            If oPhraseSuiv.iNumParagrapheL <> iMemNumParag And
               oPhraseSuiv.iNumParagrapheL <> oPhrase.iNumParagrapheL Then
                If bExporterToutAvecNumeros Then sb.Append(sInfoParag(oPhraseSuiv))
                iMemNumParag = oPhraseSuiv.iNumParagrapheL
            End If

        Next i
        sb.Append(vbCrLf)

        If Not bEcrireFichier(sCheminTxt, sb) Then Sablier(bDesactiver:=True) : Exit Sub
        Sablier(bDesactiver:=True)
        ProposerOuvrirFichier(sCheminTxt)

    End Sub

    Private Function sInfoDoc$(oPhrase As clsPhrase)

        Dim sCleDoc$ = oPhrase.sCleDoc
        Dim sCodeDoc$ = sLireCodeDoc(sCleDoc)
        ' 03/01/2010 Désactivé
        'If sCodeDoc <> sCleDoc Then sCleDoc &= " : " & sCodeDoc
        Dim sCleAffichee$ = sCleDoc
        If sCodeDoc <> sCleDoc Then sCleAffichee = sCleDoc & " : " & sCodeDoc

        ' 03/01/2010 Lors de la 1ère indexation, la clé du document associée aux phrases 
        '  est tjrs "Doc n°x" : il le reste ensuite
        Dim sChemin$ = ""
        If Me.m_colDocs.Contains(sCleDoc) Then
            sChemin = DirectCast(Me.m_colDocs(sCleDoc), clsDoc).sChemin & " "
            sCleAffichee = DirectCast(Me.m_colDocs(sCleDoc), clsDoc).sCodeDoc
        End If
        'sInfoDoc = "Document : " & sChemin & "(" & sCleDoc & ")" & vbCrLf & vbCrLf
        sInfoDoc = "Document : " & sChemin & "(" & sCleAffichee & ")" & vbCrLf & vbCrLf

    End Function

    Private Function sInfoParag$(oPhrase As clsPhrase)

        sInfoParag =
            "§G:" & oPhrase.iNumParagrapheG &
            ", §L:" & oPhrase.iNumParagrapheL &
            ", Ph.G:" & oPhrase.iNumPhraseG &
            " Ph.L:" & oPhrase.iNumPhraseL & " : "

    End Function

    Public Sub AnalyseAccents(sCodeLangIndex$, sCodesLanguesIndex$)

        ' D'abord vérifier que WinMerge est bien installé
        Dim sCheminWinMerge$ = ""
        If Not bLireCleBRWinMerge(sCheminWinMerge$) Then Exit Sub

        ' Vérifier si le dictionnaire est présent
        Dim sCheminDico0 = Application.StartupPath & sCheminDico & "_" &
            sCodeLangIndex & sExtTxt
        If Not bFichierExiste(sCheminDico0) Then
            MsgBox("Le dictionnaire est introuvable :" & vbLf &
                sCheminDico0, MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        ' Ensuite indexer à nouveau tous les documents avec les accents
        Dim bMemAccent = m_bIndexerAccents
        m_bIndexerAccents = True

        Dim bEchec As Boolean = False
        ReinitDicoAccentOuPas()
        Dim iNumFichier% = 0
        For Each oDoc As clsDoc In Me.m_colDocs
            iNumFichier += 1
            Dim sNumFichier$ = iNumFichier.ToString
            If Not bIndexerDocumentInterne(oDoc.sChemin, sNumFichier, oDoc.sCodeDoc) Then _
                bEchec = True : GoTo Fin
        Next oDoc

        ' Ensuite créer l'index des mots hors dico avec accent

        Dim sCheminDestAccent$ = ""
        CreerDocIndex(sIndexAlpha,
            bMotsDico:=False, bMotsCourants:=False, sCheminDico0:=sCheminDico0,
            sCodeLangIndex:=sCodeLangIndex,
            bMotsSeulsDocIndex0:=True, iMaxMotsCles:=0, bNumeriques:=False,
            sCodesLanguesIndex:=sCodesLanguesIndex, bCreerDocWord:=False, bProposerOuvrir:=False,
            sCheminFinal:=sCheminDestAccent)
        Dim sDossier$ = IO.Path.GetDirectoryName(sCheminDestAccent)
        Dim sFichier$ = IO.Path.GetFileNameWithoutExtension(sCheminDestAccent)
        Dim sFichierDest$ = sFichier & "_Accent.txt" '& IO.Path.GetExtension(sCheminDestAccent)
        Dim sCheminFinalAccent$ = sDossier & "\" & sFichierDest
        If Not bRenommerFichier(sCheminDestAccent, sCheminFinalAccent) Then _
            bEchec = True : GoTo Fin

        ' Ensuite indexer à nouveau tous les documents sans les accents
        m_bIndexerAccents = False
        ReinitDicoAccentOuPas()
        iNumFichier = 0
        For Each oDoc As clsDoc In Me.m_colDocs
            iNumFichier += 1
            Dim sNumFichier$ = iNumFichier.ToString
            If Not bIndexerDocumentInterne(oDoc.sChemin, sNumFichier, oDoc.sCodeDoc) Then _
                bEchec = True : GoTo Fin
        Next oDoc

        ' Ensuite créer l'index des mots hors dico sans accent
        Dim sCheminDestSansAccent$ = ""
        CreerDocIndex(sIndexAlpha,
            bMotsDico:=False, bMotsCourants:=False, sCheminDico0:=sCheminDico0,
            sCodeLangIndex:=sCodeLangIndex,
            bMotsSeulsDocIndex0:=True, iMaxMotsCles:=0, bNumeriques:=False,
            sCodesLanguesIndex:=sCodesLanguesIndex, bCreerDocWord:=False, bProposerOuvrir:=False,
            sCheminFinal:=sCheminDestSansAccent)
        Dim sFichier2$ = IO.Path.GetFileNameWithoutExtension(sCheminDestSansAccent)
        Dim sFichierDest2$ = sFichier2 & "_SansAccent.txt" '& IO.Path.GetExtension(sCheminDestSansAccent)
        Dim sCheminFinalSansAccent$ = sDossier & "\" & sFichierDest2
        If Not bRenommerFichier(sCheminDestAccent, sCheminFinalSansAccent) Then _
            bEchec = True : GoTo Fin

        ' Ouvrir WinMerge avec ces 2 index avec et sans accent
        Const sGm$ = """"
        Dim sCmd$ = sGm & sCheminFinalSansAccent & sGm & " " & sGm & sCheminFinalAccent & sGm
        Dim p As New Process
        p.StartInfo = New ProcessStartInfo(sCheminWinMerge)
        p.StartInfo.Arguments = sCmd
        p.Start()

Fin:
        m_bIndexerAccents = bMemAccent
        If bEchec Then
            AfficherMessage("Erreur !")
        Else
            AfficherMessage(sMsgOperationTerminee)
        End If

    End Sub

    Private Function bLireCleBRWinMerge(ByRef sCheminWinMerge$) As Boolean

        sCheminWinMerge = ""
        If Not bCleRegistreCUExiste("SOFTWARE\Thingamahoochie\WinMerge",
            "Executable", sCheminWinMerge) Then
            MsgBox("L'utilitaire WinMerge n'est pas installé (clé de registre non trouvée)",
                MsgBoxStyle.Critical, sTitreMsg & " - Analyse des accents")
            Return False
        End If

        ' Par défaut : "C:\Program Files\WinMerge\WinMergeU.exe"
        If sCheminWinMerge.Length = 0 Then Return False
        If Not bFichierExiste(sCheminWinMerge, bPrompt:=True) Then Return False
        Return True

    End Function

#End Region

End Class