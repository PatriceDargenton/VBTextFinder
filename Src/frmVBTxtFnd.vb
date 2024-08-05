
' VBTextFinder : un moteur de recherche de mot dans son contexte
' --------------------------------------------------------------

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' u pour Unsigned (non signé : entier positif)
' a pour Array (tableau) : ()
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

' Fichier frmVBTxtFnd.vb : 
' ----------------------

Public Class frmVBTextFinder ': Inherits Form : cf. classe partielle .Designer.vb

#Region "Déclarations"

    Private Const iOngletRechercher% = 0 ' 0 : 1er onglet
    Private Const iOngletWeb% = 1
    Private Const iOngletIndexer% = 2
    Private Const iOngletOutils% = 3
    Private Const iOngletConfig% = 4

    ' Menus contextuels
    Private Const sMenuCtx_TypeFichierTxt$ = "txtfile"
    Private Const sMenuCtx_TypeFichierDoc$ = "Word.Document.8" ' Word 2003 (Ne marche pas si on ne met que Word.Document)
    'Private Const sMenuCtx_TypeFichierHtml$ = "htmlfile" ' Ne fonctionne plus ?
    Private Const sMenuCtx_TypeFichierTous$ = "*" ' Tous les fichiers 28/06/2014
    Private Const sMenuCtx_TypeDossier$ = "Directory"

    ' Il vaut mieux indiquer VBTextFinder devant Indexer pour rappeler quel logiciel ajoute cette clé
    Private Const sMenuCtx_CleCmdIndexer$ = "VBTextFinder.Indexer"
    Private Const sMenuCtx_CleCmdIndexerDescription$ = "Indexer pour une recherche (VBTF)"

    ' Lorsque l'on créé un nouveau type de fichier, il faut d'abord placer un pointeur
    '  de l'extension vers le type : .idx -> VBTextFinder
    '  Attention : on suppose qu'aucun autre logiciel n'utilise cette clé (ce qui est vrai en standard)
    Private Const sMenuCtx_ExtFichierIdx$ = ".idx" ' Doit pointer vers sMenuCtx_TypeFichierIdx
    ' Description qui apparait dans l'explorateur à la place du nom générique
    '  fichier IDX
    Private Const sMenuCtx_ExtFichierIdxDescription$ = "index VBTextFinder"
    Private Const sMenuCtx_TypeFichierIdx$ = "VBTextFinder"
    Private Const sMenuCtx_TypeFichierIdxDescription$ =
        "Index VBTextFinder (fichier .idx)"
    Private Const sMenuCtx_CleCmdIndexOuvrir$ = "Ouvrir"
    Private Const sMenuCtx_CleCmdIndexOuvrirDescription$ = "Ouvrir avec VBTextFinder"

    ' Objet moteur de recherche : c'est l'objet principal
    '  dont ce formulaire est l'interface
    Private oVBTxtFnd As New clsVBTextFinder
    Private WithEvents m_msgDelegue As New clsMsgDelegue

    ' Initialiser seulement la première fois que la fenêtre est prête
    Private m_bInit As Boolean
    'Private m_bQuitter As Boolean

    'Private bClickParag As Boolean
    Private m_bSauverOption_bTexteUnicode As Boolean = True
    Private m_bSauverOption_bIndexerAccents As Boolean = True
    Private m_bSauverOption_bIndexerChapitrage As Boolean = True

    Private sCheminHtmlTmp$ = Application.StartupPath & "\VBTextFinderTmp.html"
    Private sCheminTxtTmp$ = Application.StartupPath & "\VBTextFinderTmp.txt"

#End Region

#Region "Initialisations"

    Private Sub TitrerAppli()

        'Me.Text &= " - Version " & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sVersionAppli$ = My.Application.Info.Version.Major &
            "." & My.Application.Info.Version.Minor &
            My.Application.Info.Version.Build
        Dim sTxt$ = "VBTextFinder : un moteur de recherche de mot dans son contexte" &
            " - Version " & sVersionAppli & " (" & sDateVersionAppli & ")"
        If Me.chkUnicode.Checked Then sTxt &= " - Unicode"
        If Me.chkAccents.Checked Then sTxt &= " - Accents"
        If bDebug Then sTxt &= " - Debug"
        Me.Text = sTxt

    End Sub

    Private Sub chkChapitrage_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkChapitrage.CheckedChanged
        Me.oVBTxtFnd.m_bIndexerChapitre = Me.chkChapitrage.Checked
        Me.chkAfficherChapitreIndex.Enabled = Me.chkChapitrage.Checked
    End Sub

    Private Sub chkUnicode_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkUnicode.CheckedChanged
        TitrerAppli()
        Me.oVBTxtFnd.m_bOptionTexteUnicode = Me.chkUnicode.Checked
    End Sub

    Private Sub chkAccents_CheckedChanged(sender As Object, e As EventArgs) _
        Handles chkAccents.CheckedChanged
        TitrerAppli()
        Me.oVBTxtFnd.IndexerAccents = Me.chkAccents.Checked
    End Sub

    Private Sub chkUnicode_Click(sender As Object, e As EventArgs) _
        Handles chkUnicode.Click
        ' Si on clique alors sauver l'option
        m_bSauverOption_bTexteUnicode = True
    End Sub

    Private Sub chkAccents_Click(sender As Object, e As EventArgs) _
        Handles chkAccents.Click
        ' Si on clique alors sauver l'option
        m_bSauverOption_bIndexerAccents = True
    End Sub

    Private Sub frmVBTextFinder_Load(sender As Object, e As EventArgs) _
        Handles MyBase.Load

        ' 04/05/2014 modUtilFichier peut maintenant être compilé dans une dll
        DefinirTitreApplication(sTitreMsg)

        Dim iTypeIndexSelect% = My.Settings.iIndexType

        Me.oVBTxtFnd.IndexerAccents = My.Settings.bIndexerAccents
        Me.oVBTxtFnd.m_bOptionTexteUnicode = My.Settings.bTexteUnicode

        Me.oVBTxtFnd.m_bIndexerChapitre = Me.chkChapitrage.Checked
        'Me.oVBTxtFnd.m_sChapitrage = Me.tbChapitrage.Text

        Me.oVBTxtFnd.Initialiser(Me.m_msgDelegue, Me.LstTypeAffichResult,
            Me.lstTypeIndex, iTypeIndexSelect)
        Me.tbChapitrage.Text = Me.oVBTxtFnd.m_sChapitrage

        TitrerAppli()
        Me.wbResultat.Navigate("")

    End Sub

    Private Sub frmVBTextFinder_Activated(eventSender As Object,
        eventArgs As EventArgs) Handles MyBase.Activated
        Activer()
    End Sub

    Private Sub frmVBTextFinder_Closing(sender As Object,
        e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Not Me.oVBTxtFnd.bQuitter() Then e.Cancel = True : Exit Sub

        SauverConfig(Me.Location, Me.Size, Me.WindowState)
        'm_bQuitter = True

    End Sub

    ' Note : l'appel à InitialiserFenetre() se trouve dans la fonction New()
    ' cf. frmVBTxtFnd.Designer.vb
    Private Sub InitialiserFenetre()

        ' Reprendre la taille et la position précédente de la fenêtre

        ' Positionnement de la fenêtre par le code : mode manuel
        Me.StartPosition = FormStartPosition.Manual
        ' Fixer la position et la taille de la feuille sauvées dans le fichier .exe.config
        Dim x% = My.Settings.frm_X
        Dim y% = My.Settings.frm_Y
        Dim w% = My.Settings.frm_Larg
        Dim h% = My.Settings.frm_Haut

        Me.Location = New Drawing.Point(x, y)
        Me.Size = New Size(w, h)

        'If My.Settings.frm_EtatFenetre = 2 Then Me.WindowState = FormWindowState.Maximized
        Select Case My.Settings.frm_EtatFenetre
            Case 0 : Me.WindowState = FormWindowState.Normal
            Case 1 : Me.WindowState = FormWindowState.Minimized
            Case 2 : Me.WindowState = FormWindowState.Maximized
        End Select
        If bDebug Then Me.StartPosition = FormStartPosition.CenterScreen

        Me.chkMotsCourants.Checked = My.Settings.bIndexAvecMotsCourant
        Me.chkMotsDico.Checked = My.Settings.bIndexAvecMotsDico
        Me.chkNumeriques.Checked = My.Settings.bIndexAvecNumeriques
        Me.tbCodeLangue.Text = My.Settings.sCodeLangueIndex
        Me.tbCodesLangues.Text = My.Settings.sListeCodesLanguesIndex
        Me.lbCodesLangues.DataSource = My.Settings.sListeCodesLanguesIndex.Split(";".ToCharArray)
        For i As Integer = 0 To Me.lbCodesLangues.Items.Count - 1
            Me.lbCodesLangues.SetSelected(i, False)
            If i > Me.lbCodesLangues.Items.Count Then Exit For
            If Me.lbCodesLangues.Items(i).ToString <> Me.tbCodeLangue.Text Then Continue For
            Me.lbCodesLangues.SetSelected(i, True)
            Exit For
        Next
        Me.chkListeMots.Checked = My.Settings.bIndexListeMots
        Me.mtbNbMotsCles.Text = My.Settings.NbMotsCles.ToString

        Me.chkUnicode.Checked = My.Settings.bTexteUnicode
        Me.chkAccents.Checked = My.Settings.bIndexerAccents

        Me.chkAfficherInfoResultat.Checked = My.Settings.bAfficherInfoResultat
        Me.chkAfficherInfoDoc.Checked = My.Settings.bAfficherInfoDoc
        Me.chkNumerotationGlobale.Checked = My.Settings.bNumerotationGlobale
        Me.chkAfficherNumParag.Checked = My.Settings.bAfficherNumParag
        Me.chkAfficherNumPhrase.Checked = My.Settings.bAfficherNumPhrase
        Me.chkAfficherNumOccur.Checked = My.Settings.bAfficherNumOccur
        Me.chkAfficherTiret.Checked = My.Settings.bAfficherTiret

        Me.chkHtmlGras.Checked = My.Settings.bOccurrencesGras
        Me.chkHtmlCouleurs.Checked = My.Settings.bOccurrencesCouleur
        Me.tbCouleursHtml.Text = My.Settings.sCouleursHtml

        Me.chkChapitrage.Checked = My.Settings.bIndexerChapitre
        'Me.tbChapitrage.Text = My.Settings.sChapitrage
        If Not Me.chkChapitrage.Checked Then
            Me.chkAfficherChapitreIndex.Checked = False
            Me.chkAfficherChapitreIndex.Enabled = False
        Else
            Me.chkAfficherChapitreIndex.Enabled = True
            Me.chkAfficherChapitreIndex.Checked = My.Settings.bAfficherChapitreIndex
        End If

        Me.chkUnicodeVerif.Checked = My.Settings.bVerifierUnicode ' 02/06/2019

    End Sub

    Private Sub SauverConfig(
        pt As Point,
        sz As Size,
        Optional ws As Windows.Forms.FormWindowState = FormWindowState.Normal,
        Optional bVerifierComposants As Boolean = True)

        ' Sauver la configuration (emplacement de la fenêtre) dans le fichier .exe.config

        If ws = FormWindowState.Normal Then
            My.Settings.frm_X = pt.X
            My.Settings.frm_Y = pt.Y
            My.Settings.frm_Larg = sz.Width
            My.Settings.frm_Haut = sz.Height
            My.Settings.frm_EtatFenetre = 0
        ElseIf ws = FormWindowState.Minimized Then
            My.Settings.frm_EtatFenetre = 1
        ElseIf ws = FormWindowState.Maximized Then
            My.Settings.frm_EtatFenetre = 2
        End If

        My.Settings.bIndexAvecMotsCourant = Me.chkMotsCourants.Checked
        My.Settings.bIndexAvecMotsDico = Me.chkMotsDico.Checked
        My.Settings.bIndexAvecNumeriques = Me.chkNumeriques.Checked
        My.Settings.sCodeLangueIndex = Me.tbCodeLangue.Text
        My.Settings.sListeCodesLanguesIndex = Me.tbCodesLangues.Text
        My.Settings.bIndexListeMots = Me.chkListeMots.Checked
        Integer.TryParse(Me.mtbNbMotsCles.Text, My.Settings.NbMotsCles)
        My.Settings.iIndexType = Me.lstTypeIndex.SelectedIndex

        My.Settings.bAfficherInfoResultat = Me.chkAfficherInfoResultat.Checked
        My.Settings.bAfficherInfoDoc = Me.chkAfficherInfoDoc.Checked
        My.Settings.bNumerotationGlobale = Me.chkNumerotationGlobale.Checked
        My.Settings.bAfficherNumParag = Me.chkAfficherNumParag.Checked
        My.Settings.bAfficherNumPhrase = Me.chkAfficherNumPhrase.Checked
        My.Settings.bAfficherNumOccur = Me.chkAfficherNumOccur.Checked
        My.Settings.bAfficherTiret = Me.chkAfficherTiret.Checked

        If m_bSauverOption_bTexteUnicode Then _
            My.Settings.bTexteUnicode = Me.chkUnicode.Checked
        If m_bSauverOption_bIndexerAccents Then _
            My.Settings.bIndexerAccents = Me.chkAccents.Checked

        My.Settings.bOccurrencesGras = Me.chkHtmlGras.Checked
        My.Settings.bOccurrencesCouleur = Me.chkHtmlCouleurs.Checked
        My.Settings.sCouleursHtml = Me.tbCouleursHtml.Text

        If m_bSauverOption_bIndexerChapitrage Then
            My.Settings.bIndexerChapitre = Me.chkChapitrage.Checked
            My.Settings.bAfficherChapitreIndex = Me.chkAfficherChapitreIndex.Checked
        End If
        'My.Settings.sChapitrage = Me.tbChapitrage.Text

        My.Settings.bVerifierUnicode = Me.chkUnicodeVerif.Checked ' 02/06/2019

        ' Si l'infrastructure de l'appli. est activée, l'appel peut être automatique
        ' (simple case à cocher)
        My.Settings.Save()

    End Sub

    Private Sub Activer()

        If m_bInit Then Exit Sub
        m_bInit = True

        VerifierMenuCtx()

        Me.CmdInterrompre.Enabled = False
        'Me.TxtMot.Enabled = False
        Me.CmdChercher.Enabled = False
        Me.CmdAjouterDocument.Enabled = False
        Me.AfficherMessage("Initialisation en cours...")

        If bDebug Then
            'Me.chkAccents.Checked = True
            'Me.chkUnicode.Checked = True
            'Me.oVBTxtFnd.m_bModeDirect = True
            'Me.oVBTxtFnd.m_sCheminFichierTxtDirect = Application.StartupPath & "\Tmp\Test.doc"
        End If

        If Me.oVBTxtFnd.m_bModeDirect Then

            Me.TxtCheminDocument.Text = Me.oVBTxtFnd.m_sCheminFichierTxtDirect
            Application.DoEvents()

            ' Convertir le fichier en .txt si son extension
            '  est celle d'un document convertible (.doc, .html ou .htm)
            'Dim bVerifierUnicode As Boolean = My.Settings.bVerifierUnicode
            Dim bVerifierUnicode As Boolean = Me.chkUnicodeVerif.Checked ' 01/06/2019
            If Me.oVBTxtFnd.bConvertirDocEnTxt(
                Me.oVBTxtFnd.m_sCheminFichierTxtDirect,
                bVerifierUnicode,
                Me.oVBTxtFnd.m_bAuMoinsUnTxtUnicode,
                Me.oVBTxtFnd.m_bAvertAuMoinsUnTxtUnicode,
                Me.oVBTxtFnd.m_bInfoAuMoinsUnTxtNonUnicode,
                bSablier:=True) Then
                Me.TxtCheminDocument.Text = Me.oVBTxtFnd.m_sCheminFichierTxtDirect
                Application.DoEvents()
                AjouterDocument()
                'CmdAjouterDocument_Click(New Object, New EventArgs)
            End If

        Else

            Dim bMemChapitrage As Boolean = Me.oVBTxtFnd.m_bIndexerChapitre
            Dim bMemUnicode As Boolean = Me.oVBTxtFnd.m_bOptionTexteUnicode
            Dim bMemAccents As Boolean = Me.oVBTxtFnd.IndexerAccents

            ' 24/05/2019 Voir s'il y a des info. sur l'unicode 
            ' (car on ne la sauve pas encore dans l'index)
            Me.oVBTxtFnd.LireListeDocumentsIndexesIni()

            If Me.oVBTxtFnd.bLireIndex() Then
                ' Si l'index contenait du unicode, alors passer en unicode
                Dim bUnicode As Boolean = Me.oVBTxtFnd.m_bOptionTexteUnicode
                Dim bAccents As Boolean = Me.oVBTxtFnd.IndexerAccents
                Dim bChapitrage As Boolean = Me.oVBTxtFnd.m_bIndexerChapitre
                If bUnicode <> bMemUnicode Then
                    Me.chkUnicode.Checked = bUnicode
                    Me.m_bSauverOption_bTexteUnicode = False ' Ne pas sauver l'option alors
                End If
                If bAccents <> bMemAccents Then
                    Me.chkAccents.Checked = bAccents
                    Me.m_bSauverOption_bIndexerAccents = False ' Ne pas sauver l'option alors
                End If
                If bChapitrage <> bMemChapitrage Then
                    Me.chkChapitrage.Checked = bChapitrage
                    Me.m_bSauverOption_bIndexerChapitrage = False ' Ne pas sauver l'option alors
                End If

                Me.oVBTxtFnd.ListerDocumentsIndexes(Me.TxtResultat, bListerPhrases:=False)
            Else
                ' Fichier document traité par défaut, pour l'exemple
                Dim sFiltreTxt$ = Me.oVBTxtFnd.m_sCheminDossierCourant & "\*.txt"
                Me.TxtCheminDocument.Text = sFiltreTxt
                Me.tcOnglets.SelectedIndex = iOngletIndexer
            End If

        End If

        VerifierOperationsPossibles()
        If Me.TxtMot.Enabled Then Me.TxtMot.Focus()

        ' Options passées en argument de la ligne de commande
        ' Cette fct ne marche pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArgLigneCmd$ = Microsoft.VisualBasic.Interaction.Command

        If bDebug AndAlso sArgLigneCmd.Length = 0 Then
            'Me.tcOnglets.SelectedIndex = iOngletOutils
            Me.TxtCheminDocument.Text = Application.StartupPath & "\Proverbes.txt"
            Me.TxtMot.Text = sGm & "Le temps" & sGm
            'Me.TxtMot.Text = "temps"
        End If

    End Sub

#End Region

#Region "Gestion des événements"

    Private Sub CmdChoisirFichierDoc_Click(eventSender As Object,
        eventArgs As EventArgs) Handles CmdChoisirFichierDoc.Click

        ' Gerer la boîte de dialogue pour choisir un fichier document Word à indexer

        Const sMsgFiltreDoc$ =
            "Document Texte (*.txt) : bloc-notes Windows|*.txt|" &
            "Document Word (*.doc)|*.doc|Document Html (*.htm ou *.html) : web|*.htm*|" &
            "Autre document (*.*)|*.*"
        Const sMsgTitreBoiteDlg$ =
            "Veuillez choisir un fichier texte ou un document convertible en .txt"
        Dim sInitDir$ = "", sFichier$ = ""
        ' Initialiser le chemin seulement la première fois
        Static bDejaInit As Boolean
        If Not bDejaInit Then
            bDejaInit = True
            sInitDir = Me.oVBTxtFnd.m_sCheminDossierCourant
        End If
        If bChoisirFichier(sFichier, sMsgFiltreDoc, "*.txt", sMsgTitreBoiteDlg,
            sInitDir, bMultiselect:=False) Then ' ToDo : traiter multisélect

            ' Convertir le fichier en .txt si son extension
            '  est celle d'un document convertible (.doc, .html ou .htm)

            'Dim bVerifierUnicode As Boolean = My.Settings.bVerifierUnicode
            Dim bVerifierUnicode As Boolean = Me.chkUnicodeVerif.Checked ' 01/06/2019
            If Me.oVBTxtFnd.bConvertirDocEnTxt(sFichier,
                bVerifierUnicode,
                Me.oVBTxtFnd.m_bAuMoinsUnTxtUnicode,
                Me.oVBTxtFnd.m_bAvertAuMoinsUnTxtUnicode,
                Me.oVBTxtFnd.m_bInfoAuMoinsUnTxtNonUnicode,
                bSablier:=True) Then _
                Me.TxtCheminDocument.Text = sFichier

            Me.oVBTxtFnd.Sablier(bDesactiver:=True)
        End If
        VerifierOperationsPossibles(bVerifDocumentSeul:=True)

    End Sub

    Private Sub CmdAjouterDocument_Click(eventSender As Object,
        eventArgs As EventArgs) Handles CmdAjouterDocument.Click
        AjouterDocument()
    End Sub

    Private Sub CmdChercher_Click(eventSender As Object, eventArgs As EventArgs) _
        Handles CmdChercher.Click
        Chercher()
    End Sub

    Private Sub CmdInterrompre_Click(eventSender As Object,
        eventArgs As EventArgs) Handles CmdInterrompre.Click
        Me.m_msgDelegue.m_bAnnuler = True
        Me.oVBTxtFnd.Interrompre()
    End Sub

    Private Sub TxtCheminDocument_TextChanged(eventSender As Object,
        eventArgs As EventArgs) Handles TxtCheminDocument.TextChanged
        VerifierActivationCmdIndexer()
    End Sub

    Private Sub TxtCheminDocument_DoubleClick(sender As Object, e As EventArgs) _
        Handles TxtCheminDocument.DoubleClick
        BasculerFiltreIndexationFichiers()
    End Sub

    Private Sub LstTypeIndex_Click(sender As Object, e As EventArgs) _
        Handles lstTypeIndex.Click
        AfficherDescriptionDocIndex()
    End Sub

    Private Sub LstTypeIndex_DoubleClick(eventSender As Object,
        eventArgs As EventArgs) Handles lstTypeIndex.DoubleClick
        CreerDocIndex()
    End Sub

    'Private Sub LstTypeAffichResult_SelectedIndexChanged(eventSender As Object, _
    '    eventArgs As EventArgs) Handles LstTypeAffichResult.SelectedIndexChanged
    '    ScrollParagPredef()
    'End Sub

    Private Sub LstTypeAffichResult_Click(eventSender As Object,
        eventArgs As EventArgs) Handles LstTypeAffichResult.Click
        ScrollParagPredef()
    End Sub

    Private Sub vsbZoomParag_ValueChanged(sender As Object, e As EventArgs) _
        Handles vsbZoomParag.ValueChanged
        ScrollParag()
    End Sub

    Private Sub TxtMot_TextChanged(eventSender As Object,
        eventArgs As EventArgs) Handles TxtMot.TextChanged
        VerifierOperationsPossibles()
    End Sub

    Private Sub TxtMot_KeyDown(eventSender As Object,
        eventArgs As Windows.Forms.KeyEventArgs) Handles TxtMot.KeyDown
        Dim iTouche% = eventArgs.KeyCode
        'Dim Shift% = eventArgs.KeyData \ &H10000
        AfficherMotsCompatEnCoursDeFrappe(iTouche)
    End Sub

    Private Sub TxtResultat_MouseUp(eventSender As Object,
        eventArgs As EventArgs) Handles TxtResultat.MouseUp
        ' Si on relache la souris, alors noter le curseur
        Me.oVBTxtFnd.NoterPositionCurseur(Me.TxtResultat, Me.chkAfficherInfoResultat.Checked,
            Me.chkAfficherNumParag.Checked, Me.chkAfficherNumPhrase.Checked)
    End Sub

    'Private Sub TxtResultat_Leave(sender As Object, e As EventArgs) _
    '    Handles TxtResultat.LostFocus
    '    If m_bQuitter Then Exit Sub
    '    Debug.WriteLine(Now & " : Memo pos. curseur")
    '    Me.oVBTxtFnd.NoterPositionCurseur(Me.TxtResultat, Me.chkInfoParag.Checked)
    'End Sub

    Private Sub TxtResultat_DoubleClick(eventSender As Object,
        eventArgs As EventArgs) Handles TxtResultat.DoubleClick
        HyperTexte()
    End Sub

    Private Sub chkAfficherInfoResultat_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherInfoResultat.Click
        Chercher()
    End Sub

    Private Sub chkAfficherInfoDoc_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherInfoDoc.Click
        Chercher()
    End Sub

    Private Sub chkAfficherNumParag_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherNumParag.Click
        Chercher()
    End Sub

    Private Sub chkAfficherNumPhrase_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherNumPhrase.Click
        Chercher()
    End Sub

    Private Sub chkNumerotationGlobale_Click(sender As Object, e As EventArgs) _
        Handles chkNumerotationGlobale.Click
        Chercher()
    End Sub

    Private Sub chkAfficherNumOccur_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherNumOccur.Click
        Chercher()
    End Sub

    Private Sub chkAfficherTiret_Click(sender As Object, e As EventArgs) _
        Handles chkAfficherTiret.Click
        Chercher()
    End Sub

    Private Sub chkMotsDico_Click(sender As Object, e As EventArgs) _
        Handles chkMotsDico.Click
        VerifierDico()
    End Sub

    Private Sub lbCodesLangues_Click(sender As Object, e As EventArgs) _
        Handles lbCodesLangues.Click
        Me.tbCodeLangue.Text = Me.lbCodesLangues.SelectedItem.ToString
        Me.oVBTxtFnd.ReinitDico() ' 03/05/2014 Penser à recharger le dico si on change de langue
    End Sub

    'Private Sub frmVBTextFinder_DoubleClick(sender As Object, _
    '    e As EventArgs) Handles MyBase.DoubleClick
    '    ListerDocumentsIndexes()
    'End Sub

    Private Sub cmdListeDoc_Click(sender As Object, e As EventArgs) _
        Handles cmdListeDoc.Click
        ListerDocumentsIndexes()
    End Sub

    Private Sub cmdListeDocHtml_Click(sender As Object, e As EventArgs) _
        Handles cmdListeDocHtml.Click
        ListerDocumentsIndexes(bHtml:=True)
        AfficherHtml(bVerifierIdem:=False)
    End Sub

    Private Sub tcOnglets_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles tcOnglets.SelectedIndexChanged
        AfficherHtml()
    End Sub

    Private Sub wbResultat_DocumentCompleted(sender As Object,
        e As Windows.Forms.WebBrowserDocumentCompletedEventArgs) _
        Handles wbResultat.DocumentCompleted
        ' Conserver le fichier, car on pourra l'afficher dans le navigateur externe
        'If bDebug Then Exit Sub
        'If Not bFichierExiste(sCheminHtmlTmp) Then Exit Sub
        'bSupprimerFichier(sCheminHtmlTmp)
    End Sub

    Private Sub cmdNavigExterne_Click(sender As Object, e As EventArgs) _
        Handles cmdNavigExterne.Click
        If IsNothing(oVBTxtFnd.m_sbResultatHtml) Then Exit Sub
        If Not bFichierExiste(sCheminHtmlTmp) Then Exit Sub
        OuvrirAppliAssociee(sCheminHtmlTmp)
    End Sub

    Private Sub cmdExporterTxt_Click(sender As Object, e As EventArgs) _
        Handles cmdExporterTxt.Click
        If IsNothing(oVBTxtFnd.m_sbResultatTxt) Then Exit Sub
        Dim iEncodage% = iCodePageWindowsLatin1252
        If oVBTxtFnd.m_bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8
        If Not bEcrireFichier(sCheminTxtTmp, oVBTxtFnd.m_sbResultatTxt,
            iEncodage:=iEncodage) Then Exit Sub
        OuvrirAppliAssociee(sCheminTxtTmp)
        'bSupprimerFichier(sCheminTxtTmp)
    End Sub

#End Region

#Region "Indexation"

    Private Sub BasculerFiltreIndexationFichiers()

        Static sMemChemin$

        Dim sFiltreTxt$ = Me.oVBTxtFnd.m_sCheminDossierCourant & "\*.txt"
        Dim sFiltreDoc$ = Me.oVBTxtFnd.m_sCheminDossierCourant & "\*.doc"
        Dim sFiltreHtm$ = Me.oVBTxtFnd.m_sCheminDossierCourant & "\*.htm?"

        If Me.TxtCheminDocument.Text <> sFiltreTxt And
           Me.TxtCheminDocument.Text <> sFiltreDoc And
           Me.TxtCheminDocument.Text <> sFiltreHtm Then
            sMemChemin = Me.TxtCheminDocument.Text
        End If

        If Me.TxtCheminDocument.Text = sFiltreTxt Then
            Me.TxtCheminDocument.Text = sFiltreDoc
        ElseIf Me.TxtCheminDocument.Text = sFiltreDoc Then
            Me.TxtCheminDocument.Text = sFiltreHtm
        ElseIf Me.TxtCheminDocument.Text = sFiltreHtm Then
            If sMemChemin <> "" Then
                Me.TxtCheminDocument.Text = sMemChemin
            Else
                Me.TxtCheminDocument.Text = sFiltreTxt
            End If
        Else
            Me.TxtCheminDocument.Text = sFiltreTxt
        End If

        VerifierActivationCmdIndexer()

    End Sub

    Private Sub AjouterDocument()

        ' Indexer un nouveau document

        ' Interdire la ré-entrance dans cette fonction
        Me.CmdAjouterDocument.Enabled = False
        ' Autoriser l'interruption de l'indexation
        Me.CmdInterrompre.Enabled = True
        Me.CmdChercher.Enabled = False

        Me.oVBTxtFnd.m_bIndexerChapitre = Me.chkChapitrage.Checked
        'Me.oVBTxtFnd.m_sChapitrage = Me.tbChapitrage.Text

        'Dim bVerifierUnicode As Boolean = My.Settings.bVerifierUnicode
        Dim bVerifierUnicode As Boolean = Me.chkUnicodeVerif.Checked ' 01/06/2019
        If Me.oVBTxtFnd.bIndexerDocuments(Me.TxtCheminDocument.Text, bVerifierUnicode) Then
            ' 28/08/2009 Onglet résultat de l'indexation : onglet n°1
            Me.tcOnglets.SelectedIndex = iOngletRechercher
            Me.TxtMot.Focus()
        End If

        Me.CmdInterrompre.Enabled = False
        Me.oVBTxtFnd.ListerDocumentsIndexes(Me.TxtResultat)
        VerifierOperationsPossibles()

    End Sub

#End Region

#Region "Traitements"

    Private Sub AfficherMotsCompatEnCoursDeFrappe(iTouche%)

        ' Traiter la touche Entrée sur la zone de saisie n°1
        If iTouche <> Windows.Forms.Keys.Return Then Exit Sub
        Me.CmdInterrompre.Enabled = True
        Me.CmdInterrompre.Focus()
        Chercher()
        Me.TxtMot.Focus()

    End Sub

    Private Sub Chercher()

        ' Chercher les occurrences d'un mot

        ' 01/05/2010 Inutile de proposer d'interrompre, car il n'y a qu'une seule ligne
        '  de code qui prend du temps : CtrlResultat.Text = sbResultat.ToString
        '  et on ne peut l'annuler (pas de AppendText possible ?)
        ' Oui mais alors faire une fct ActivationCmd(bDésactiver) et un booleen
        '  car on utilise Me.CmdInterrompre.Enabled pour savoir
        ' On pourra cependant annuler le ctrl web
        Me.CmdInterrompre.Enabled = True
        Me.CmdChercher.Enabled = False ' Eviter la ré-entrance dans la fonction
        Me.oVBTxtFnd.InitNouvelleRecherche() ' Effacer mémo. curseur : nouv. recherche
        ChercherDirect()
        Me.CmdInterrompre.Enabled = False
        Me.CmdChercher.Enabled = True

    End Sub

    Private Sub ChercherDirect()

        ' Faire une recherche, ou refaire une recherche avec un affichage différent

        Dim sExpression$ = Me.TxtMot.Text
        Dim bMotExiste As Boolean = False
        Dim oMot As clsMot = Nothing
        bMotExiste = Me.oVBTxtFnd.bMotExiste(sExpression, oMot)

        Dim bGuillemet As Boolean = False
        Dim sSepGm$ = Chr(iCodeASCIIGuillemet)
        If sExpression.IndexOf(sSepGm) > -1 Then bGuillemet = True

        Dim bEspace As Boolean = (sExpression.IndexOf(" ") > 1)

        If Not bGuillemet And Not bMotExiste And Not bEspace Then Exit Sub

        Dim bHtml As Boolean = False
        If Me.tcOnglets.SelectedIndex = iOngletWeb Then bHtml = True

        Me.oVBTxtFnd.m_bOccurrencesEnGras = Me.chkHtmlGras.Checked
        Me.oVBTxtFnd.m_bOccurrencesEnCouleurs = Me.chkHtmlCouleurs.Checked
        Me.oVBTxtFnd.m_sCouleursHtml = Me.tbCouleursHtml.Text
        Me.oVBTxtFnd.m_bNumerotationGlobale = Me.chkNumerotationGlobale.Checked

        If bGuillemet Or bEspace Then
            Me.oVBTxtFnd.ChercherOccurrencesMots(
                Me.TxtMot, Me.TxtResultat, Me.vsbZoomParag.Value,
                Me.chkAfficherInfoResultat.Checked, Me.chkAfficherInfoDoc.Checked,
                Me.chkAfficherNumParag.Checked, Me.chkAfficherNumPhrase.Checked,
                Me.chkAfficherNumOccur.Checked, Me.chkAfficherTiret.Checked, bHtml)
        Else
            Me.oVBTxtFnd.ChercherOccurrencesMot(
                Me.TxtMot, Me.TxtResultat, Me.vsbZoomParag.Value,
                Me.chkAfficherInfoResultat.Checked, Me.chkAfficherInfoDoc.Checked,
                Me.chkAfficherNumParag.Checked, Me.chkAfficherNumPhrase.Checked,
                Me.chkAfficherNumOccur.Checked, Me.chkAfficherTiret.Checked, bHtml)
        End If

    End Sub

    Private Sub HyperTexte()

        ' Quitter si une opération est en cours
        If Me.CmdInterrompre.Enabled Then Exit Sub
        Dim sMotSelFin$ = ""
        If Me.oVBTxtFnd.bHyperTexte((Me.TxtResultat.SelectedText), sMotSelFin) Then
            Me.TxtMot.Text = sMotSelFin
            Chercher()
        End If

    End Sub

    Private Sub AfficherHtml(Optional bVerifierIdem As Boolean = True)

        If Me.tcOnglets.SelectedIndex <> iOngletWeb Then Exit Sub

        Static iMemNbParag% = -2
        Static bMemAfficherInfoResultat, bMemAfficherNumOccur As Boolean
        Static bMemAfficherInfoDoc, bMemNumerotationGlobale As Boolean
        Static bMemAfficherNumParag, bMemAfficherNumPhrase As Boolean
        Static bMemAfficherTiret As Boolean
        Static sMemExpressions$ = ""
        Static bMemHtmlGras, bMemHtmlCouleur As Boolean
        Static sMemCouleursHtml$ = ""
        Static bMemChapitre As Boolean
        Static iMemNbDoc% = 0
        Static bMemListeDoc As Boolean

        If Not bVerifierIdem Then GoTo Suite

        ' Si l'affichage a changé alors refaire le html complètement
        If bMemAfficherInfoResultat <> chkAfficherInfoResultat.Checked OrElse
           bMemAfficherInfoDoc <> chkAfficherInfoDoc.Checked OrElse
           bMemNumerotationGlobale <> chkNumerotationGlobale.Checked OrElse
           bMemAfficherNumParag <> chkAfficherNumParag.Checked OrElse
           bMemAfficherNumPhrase <> chkAfficherNumPhrase.Checked OrElse
           bMemAfficherNumOccur <> chkAfficherNumOccur.Checked OrElse
           bMemAfficherTiret <> chkAfficherTiret.Checked OrElse
           iMemNbParag <> vsbZoomParag.Value OrElse
           sMemExpressions <> TxtMot.Text OrElse
           bMemHtmlGras <> chkHtmlGras.Checked OrElse
           bMemHtmlCouleur <> chkHtmlCouleurs.Checked OrElse
           sMemCouleursHtml <> tbCouleursHtml.Text OrElse
           bMemChapitre <> chkAfficherChapitreIndex.Checked OrElse
           iMemNbDoc <> oVBTxtFnd.iNbDocumentsIndexes OrElse
           bMemListeDoc <> (Not bVerifierIdem) Then
            Chercher()
            'If IsNothing(oVBTxtFnd.m_sbResultatHtml) Then Exit Sub
        Else
            'If Not IsNothing(oVBTxtFnd.m_sbResultatHtml) Then
            '    ' Déjà affiché correctement
            '    Exit Sub
            'End If
            Exit Sub
        End If

        bMemAfficherInfoResultat = chkAfficherInfoResultat.Checked
        bMemAfficherInfoDoc = chkAfficherInfoDoc.Checked
        bMemNumerotationGlobale = chkNumerotationGlobale.Checked
        bMemAfficherNumParag = chkAfficherNumParag.Checked
        bMemAfficherNumPhrase = chkAfficherNumPhrase.Checked
        bMemAfficherNumOccur = chkAfficherNumOccur.Checked
        bMemAfficherTiret = chkAfficherTiret.Checked

        sMemExpressions = TxtMot.Text
        iMemNbParag = vsbZoomParag.Value
        bMemHtmlGras = chkHtmlGras.Checked
        bMemHtmlCouleur = chkHtmlCouleurs.Checked
        sMemCouleursHtml = tbCouleursHtml.Text
        bMemChapitre = chkAfficherChapitreIndex.Checked
        iMemNbDoc = oVBTxtFnd.iNbDocumentsIndexes

Suite:
        ' Si on ne vérifie pas, c'est pour afficher la liste des docs
        bMemListeDoc = Not bVerifierIdem
        Me.wbResultat.Focus() ' Recevoir les raccourcis clavier
        Me.wbResultat.Navigate("")

        If IsNothing(oVBTxtFnd.m_sbResultatHtml) Then Exit Sub

        'Dim iEncodage% = iCodePageWindowsLatin1252
        'If oVBTxtFnd.m_bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8
        ' 26/10/2019 Tous les documents html doivent être en UTF8 (ça doit être l'encodage html par défaut)
        If Not bEcrireFichier(sCheminHtmlTmp, oVBTxtFnd.m_sbResultatHtml,
            bEncodageUTF8:=True) Then Exit Sub

        ' Si le texte est trop long, éviter le menu Copié/Collé : très très long !
        '  alors que navigateur externe : ok
        ' On laisse le ctrl-C mais attention avec ctrl-A ctrl-C : très très long !
        If oVBTxtFnd.m_sbResultatHtml.Length > iMaxLongChaine0 Then
            Me.wbResultat.IsWebBrowserContextMenuEnabled = False
        Else
            Me.wbResultat.IsWebBrowserContextMenuEnabled = True
        End If

        Me.wbResultat.Navigate(New System.Uri(sCheminHtmlTmp))

    End Sub

#End Region

#Region "Divers"

    Private Sub ScrollParagPredef()

        'Debug.WriteLine(Now & " : ScrollParagPredef")

        ' Gerer le type d'affichage des résultats (phrase ou paragraphe)

        ' Verifier si sélectionné, pas seulement si text = x
        'Me.bClickParag = True
        If Me.LstTypeAffichResult.SelectedIndices.Contains(4) Then     ' +-3 §
            Me.vsbZoomParag.Value = 3
        ElseIf Me.LstTypeAffichResult.SelectedIndices.Contains(3) Then ' +-2 §
            Me.vsbZoomParag.Value = 2
        ElseIf Me.LstTypeAffichResult.SelectedIndices.Contains(2) Then ' +-1 §
            Me.vsbZoomParag.Value = 1
        ElseIf Me.LstTypeAffichResult.SelectedIndices.Contains(1) Then ' § du mot trouvé
            Me.vsbZoomParag.Value = 0
        ElseIf Me.LstTypeAffichResult.SelectedIndices.Contains(0) Then ' Phrase du mot trouvé
            Me.vsbZoomParag.Value = -1
        Else
            ' § > +-3
        End If
        'Me.bClickParag = False

    End Sub

    Private Sub ScrollParag()

        ' Gerer le type d'affichage des résultats (phrase ou paragraphe)

        'Dim bNoterPosCurseur As Boolean = True
        If Me.vsbZoomParag.Value = 3 Then    ' +-3 §
            Me.LstTypeAffichResult.Text = clsVBTextFinder.sAfficherParagPM3
        ElseIf Me.vsbZoomParag.Value = 2 Then ' +-2 §
            Me.LstTypeAffichResult.Text = clsVBTextFinder.sAfficherParagPM2
        ElseIf Me.vsbZoomParag.Value = 1 Then ' +-1 §
            Me.LstTypeAffichResult.Text = clsVBTextFinder.sAfficherParagPM1
        ElseIf Me.vsbZoomParag.Value = 0 Then    ' § du mot trouvé
            Me.LstTypeAffichResult.Text = clsVBTextFinder.sAfficherParag
        ElseIf Me.vsbZoomParag.Value = -1 Then
            ' Phrase du mot trouvé
            Me.LstTypeAffichResult.Text = clsVBTextFinder.sAfficherPhrase
            'bNoterPosCurseur = False
        Else
            Dim i%
            For i = 0 To Me.LstTypeAffichResult.Items.Count - 1
                Me.LstTypeAffichResult.SetSelected(i, False)
            Next
        End If

        If Not m_bInit Then Exit Sub
        If Not Me.CmdChercher.Enabled Then GoTo Fin
        ChercherDirect()
        Exit Sub

Fin:
        Me.LblAvancement.Text = "Nombre de paragraphes affichés autour du mot : " &
            Me.vsbZoomParag.Value

    End Sub

    Private Sub VerifierOperationsPossibles(
        Optional bVerifDocumentSeul As Boolean = False)

        ' Vérifier les opérations possibles selon l'état de l'interface

        ' Si une indexation est en cours, ne pas réactiver les boutons de commande
        If Me.CmdInterrompre.Enabled Then Exit Sub

        Dim sMsgMot$ = "", sMsg$ = "", sMsgDoc$ = ""
        If Me.oVBTxtFnd.iNbDocumentsIndexes > 0 Then
            Me.TxtMot.Enabled = True
        Else
            sMsgMot = "Aucun document n'est indexé"
        End If

        ' 10/10/2009 Nouvelle méthode de recherche : parcourir les phrases
        '  toujours laisser la possibilité de chercher une expression
        Me.CmdChercher.Enabled = True
        ' Activer le bouton Chercher si le mot existe
        'Me.CmdChercher.Enabled = False
        Dim oMot As clsMot = Nothing
        If Not Me.oVBTxtFnd.bMotExiste(Me.TxtMot.Text, oMot) Then
            If Me.TxtMot.Text <> "" Then sMsgMot = "Mot non trouvé : " & Me.TxtMot.Text
        Else
            sMsgMot = "Mot trouvé : " & Me.TxtMot.Text &
                " (" & oMot.iNbOccurrences & " occurrences)"
            Me.CmdChercher.Enabled = True
        End If

        ' Vérifier si le fichier document existe
        Me.CmdAjouterDocument.Enabled = False
        If Not bFichierExisteFiltre2(Me.TxtCheminDocument.Text) Then
            If Me.TxtCheminDocument.Text <> "" Then _
                sMsgDoc = "Fichier inexistant : " & Me.TxtCheminDocument.Text
            GoTo Fin
        End If

        ' Activer le bouton Ajouter (un document à indexer)
        Me.CmdAjouterDocument.Enabled = True

Fin:
        sMsg = sMsgMot
        If sMsgDoc <> "" Then sMsg = sMsgDoc
        If bVerifDocumentSeul Then sMsg = sMsgDoc
        If sMsg <> "" Or Not bVerifDocumentSeul Then Me.LblAvancement.Text = sMsg

    End Sub

    Private Sub VerifierActivationCmdIndexer()

        ' Vérifier si le bouton Indexer peut être activé

        ' Si une indexation est en cours, ne pas réactiver les boutons de commande
        If Me.CmdInterrompre.Enabled Then Exit Sub

        Dim sMsg$ = ""
        ' Vérifier si le fichier document existe
        Me.CmdAjouterDocument.Enabled = False
        If Not bFichierExisteFiltre2(Me.TxtCheminDocument.Text) Then
            If Me.TxtCheminDocument.Text <> "" Then _
                sMsg = "Fichier inexistant : " & Me.TxtCheminDocument.Text
            GoTo Fin
        End If

        ' Activer le bouton Ajouter (un document à indexer)
        Me.CmdAjouterDocument.Enabled = True

Fin:
        Me.LblAvancement.Text = sMsg

    End Sub

    Private Sub ListerDocumentsIndexes(Optional bHtml As Boolean = False)

        Me.oVBTxtFnd.LireListeDocumentsIndexesIni()
        Me.oVBTxtFnd.ListerDocumentsIndexes(Me.TxtResultat,
            bListerPhrases:=Not bHtml, bHtml:=bHtml)
        If Not bHtml Then Me.oVBTxtFnd.AfficherFichierIni()

    End Sub

    Private Sub CreerDocIndex()

        If Me.oVBTxtFnd.iNbDocumentsIndexes = 0 Then
            Me.LblAvancement.Text = "Aucun document indexé"
            Exit Sub
        End If

        ' Quitter si une opération est en cours
        If Me.CmdInterrompre.Enabled Then Exit Sub

        Dim sCheminDico0 = ""
        If Not Me.chkMotsDico.Checked AndAlso Not bChercherDico(sCheminDico0) Then
            MsgBox("Le dictionnaire est introuvable :" & vbLf &
                sCheminDico0, MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim iNbMotsCles% = 0
        Integer.TryParse(Me.mtbNbMotsCles.Text, iNbMotsCles)
        If iNbMotsCles = 0 Then iNbMotsCles = iMaxMotsClesDef

        Me.oVBTxtFnd.m_bAfficherChapitreIndex = Me.chkAfficherChapitreIndex.Checked

        Me.CmdInterrompre.Enabled = True
        Me.oVBTxtFnd.CreerDocIndex(Me.lstTypeIndex.Text,
            Me.chkMotsDico.Checked, Me.chkMotsCourants.Checked,
            sCheminDico0, Me.tbCodeLangue.Text,
            Me.chkListeMots.Checked, iNbMotsCles, Me.chkNumeriques.Checked,
            Me.tbCodesLangues.Text)
        Me.CmdInterrompre.Enabled = False

    End Sub

    Private Sub cmdGlossaire_Click(sender As Object, e As EventArgs) Handles cmdGlossaire.Click

        If Me.TxtCheminDocument.Text.Length = 0 Then
            ' 05/05/2018
            MsgBox("Le nom du fichier à analyser n'est pas précisé !", MsgBoxStyle.Exclamation)
            Me.tcOnglets.SelectedIndex = iOngletIndexer
            Exit Sub
        End If
        If Not bFichierExiste(Me.TxtCheminDocument.Text, bPrompt:=True) Then Exit Sub

        Me.m_msgDelegue.m_bAnnuler = False
        Me.CmdInterrompre.Enabled = True
        Me.cmdGlossaire.Enabled = False
        Me.tcOnglets.Enabled = False

        Dim bTriFreq As Boolean = False
        If Me.lstTypeIndex.Text = clsVBTextFinder.sIndexFreq Then bTriFreq = True
        bCreerGlossaire(Me.TxtCheminDocument.Text, Me.lbCodesLangues.Text, m_msgDelegue,
            bTriFreq, bVoirGlossaireCourant:=False)

        Me.m_msgDelegue.m_bAnnuler = False
        Me.CmdInterrompre.Enabled = False
        Me.tcOnglets.Enabled = True
        Me.cmdGlossaire.Enabled = True

    End Sub

    Private Sub AfficherDescriptionDocIndex()

        Select Case Me.lstTypeIndex.Text
            Case clsVBTextFinder.sIndexAlpha : Me.LblAvancement.Text =
            "Double-clic pour créer le document index par ordre alphabétique"
            Case clsVBTextFinder.sIndexFreq : Me.LblAvancement.Text =
            "Double-clic pour créer le document index par ordre fréquentiel"
            Case clsVBTextFinder.sIndexMotsCles : Me.LblAvancement.Text =
            "Double-clic pour extraire les mots clés"
            Case clsVBTextFinder.sIndexCitations : Me.LblAvancement.Text =
            "Double-clic pour extraire la liste des citations"
            Case clsVBTextFinder.sIndexSimple : Me.LblAvancement.Text =
            "Double-clic pour extraire la simple liste des mots indexés (avec le suffixe du code langue sélectionné)"
            Case clsVBTextFinder.sIndexSimpleComparer : Me.LblAvancement.Text =
            "Double-clic pour faire l'intersection des index simples dans les codes langues disponibles"
            Case clsVBTextFinder.sIndexTout : Me.LblAvancement.Text =
            "Double-clic pour exporter toutes les phrases indexées"
            Case clsVBTextFinder.sIndexNGrammes : Me.LblAvancement.Text =
            "Double-clic pour exporter les N-Grammes les plus fréquents du dictionnaire de mot"
            Case clsVBTextFinder.sIndexEspacesInsecables : Me.LblAvancement.Text =
            "Double-clic pour exporter tous les espaces insécables"
            Case clsVBTextFinder.sIndexEspacesInsecablesAVerifier : Me.LblAvancement.Text =
            "Double-clic pour exporter les espaces insécables à vérifier"
            Case clsVBTextFinder.sIndexMajuscules : Me.LblAvancement.Text =
            "Double-clic pour exporter les majuscules intempestives"
            Case clsVBTextFinder.sIndexAccents : Me.LblAvancement.Text =
            "Double-clic pour analyser les accents manquants (sur les majuscules notamment)"
        End Select

    End Sub

    Private Sub VerifierDico()

        If Me.chkMotsDico.Checked Then Exit Sub

        Dim sCheminDico0$ = ""
        If bChercherDico(sCheminDico0) Then Exit Sub

        Me.chkMotsDico.Checked = True ' Rétablir
        'Me.chkMotsDico.Checked = False

        Dim sUrl$ = ""
        Select Case Me.tbCodeLangue.Text
            Case sCodeLangueFr : sUrl = sURLDicoFr
            Case sCodeLangueEn : sUrl = sURLDicoEn
            Case sCodeLangueUk : sUrl = sURLDicoUk
            Case sCodeLangueUS : sUrl = sURLDicoUs
            Case Else : sUrl = ""
        End Select
        If sUrl.Length = 0 Then
            MsgBox(
                "Le dictionnaire est introuvable :" & vbLf &
                sCheminDico0, MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If MsgBoxResult.Cancel = MsgBox(
            "Le dictionnaire est introuvable :" & vbLf &
            sCheminDico0 & vbLf &
            "Cliquez sur OK pour le télécharger :" & vbLf & sUrl,
            MsgBoxStyle.Exclamation Or MsgBoxStyle.OkCancel, sTitreMsg) Then Exit Sub

        OuvrirAppliAssociee(sUrl, bVerifierFichier:=False)

    End Sub

    Private Function bChercherDico(ByRef sCheminDicoFinal$) As Boolean

        ' 05/05/2018 Dictionnaire _Fr pour le français aussi
        Dim sCheminDico0 = Application.StartupPath & sCheminDico & "_" &
            Me.tbCodeLangue.Text & sExtTxt
        sCheminDicoFinal = sCheminDico0
        Dim bExiste0 = bFichierExiste(sCheminDico0)
        Return bExiste0
        'If Me.tbCodeLangue.Text <> sCodeLangueFr Then
        '    If Not bFichierExiste(sCheminDico0) Then Return False
        'Else
        '    ' Si Fr alors vérifier aussi le fichier de la version précédante
        '    Dim sCheminDico1 = Application.StartupPath & sCheminDicoV1Fr
        '    Dim bExiste1 = bFichierExiste(sCheminDico1)
        '    If Not bExiste0 And Not bExiste1 Then Return False
        '    If Not bExiste0 And bExiste1 Then sCheminDico0 = sCheminDico1
        'End If
        'sCheminDicoFinal = sCheminDico0
        'Return True

    End Function

    Private Sub AfficherMessage(sMsg$)

        Me.LblAvancement.Text = sMsg
        ' Laisser du temps pour le traitement des messages : affichage du message et
        '  traitement du clic éventuel sur le bouton Interrompre
        Application.DoEvents()

    End Sub

    Private Sub AfficherMsgDelegue(sender As Object,
        e As clsMsgEventArgs) Handles m_msgDelegue.EvAfficherMessage
        Me.AfficherMessage(e.sMessage)
    End Sub

    Private Sub AfficherMsgDelegue(sender As Object,
        e As clsSablierEventArgs) Handles m_msgDelegue.EvSablier
        Sablier(e.bDesactiver)
    End Sub

    Private Sub Sablier(Optional bDesactiver As Boolean = False)

        ' Me.Cursor : Curseur de la fenêtre
        ' Cursor.Current : Curseur de l'application

        If bDesactiver Then
            Me.Cursor = Cursors.Default
        Else
            Me.Cursor = Cursors.WaitCursor
        End If

        ' Curseur de l'application : il est réinitialisé à chaque Application.DoEvents
        '  ou bien lorsque l'application ne fait rien
        '  du coup, il faut insister grave pour conserver le contrôle du curseur tout en 
        '  voulant afficher des messages de progression et vérifier les interruptions...
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            ctrl.Cursor = Me.Cursor ' Curseur de chaque contrôle de la feuille
        Next ctrl
        Cursor.Current = Me.Cursor

    End Sub

#End Region

#Region "Gestion des menus contextuels"

    Private Sub VerifierMenuCtx()

        Dim sCleDescriptionCmd$ = sMenuCtx_TypeFichierIdx & "\shell\" &
            sMenuCtx_CleCmdIndexOuvrir
        If bCleRegistreCRExiste(sCleDescriptionCmd) Then
            Me.cmdAjouterMenuCtx.Enabled = False
            Me.cmdEnleverMenuCtx.Enabled = True

            ' Si la clé existe pour .doc, voir s'il faut enlever aussi celle pour tous les fichiers (*.*)
            Dim sCleDescriptionCmdTous$ = sMenuCtx_TypeFichierTous & "\shell\" &
                sMenuCtx_CleCmdIndexer
            Me.chkTous.Checked = bCleRegistreCRExiste(sCleDescriptionCmdTous)
            Me.chkTous.Enabled = False ' Interdire de décocher

        Else
            Me.cmdAjouterMenuCtx.Enabled = True
            Me.cmdEnleverMenuCtx.Enabled = False

            Me.chkTous.Enabled = True ' Autoriser à cocher
            'Me.chkTous.Checked = True ' Coché par défaut

        End If

    End Sub

    Private Sub cmdAjouterMenuCtx_Click(sender As Object,
        e As EventArgs) Handles cmdAjouterMenuCtx.Click
        AjouterMenuCtx()
        VerifierMenuCtx()
    End Sub

    Private Sub cmdEnleverMenuCtx_Click(sender As Object,
        e As EventArgs) Handles cmdEnleverMenuCtx.Click
        EnleverMenuCtx()
        VerifierMenuCtx()
    End Sub

    Private Sub AjouterMenuCtx()

        If MsgBoxResult.Cancel = MsgBox("Ajouter les menus contextuels ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question) Then Exit Sub

        AjouterMenuCtxIndexer(sMenuCtx_TypeFichierTxt)
        AjouterMenuCtxIndexer(sMenuCtx_TypeFichierDoc)
        If Me.chkTous.Checked Then AjouterMenuCtxIndexer(sMenuCtx_TypeFichierTous)
        AjouterMenuCtxIndexer(sMenuCtx_TypeDossier)

        Dim sCheminExe$ = Application.ExecutablePath
        Const bPrompt As Boolean = False
        Const sChemin$ = """%1"""

        ' Ajouter un pointeur HKCR\.idx vers HKCR\VBTextFinder
        bAjouterTypeFichier(sMenuCtx_ExtFichierIdx, sMenuCtx_TypeFichierIdx,
            sMenuCtx_ExtFichierIdxDescription)

        ' Menu contextuel pour ouvrir un index .idx
        bAjouterMenuContextuel(sMenuCtx_TypeFichierIdx, sMenuCtx_CleCmdIndexOuvrir,
            bPrompt, , sMenuCtx_CleCmdIndexOuvrirDescription, sCheminExe, sChemin,
            sMenuCtx_TypeFichierIdxDescription)

    End Sub

    Private Sub EnleverMenuCtx()

        If MsgBoxResult.Cancel = MsgBox("Enlever les menus contextuels ?",
            MsgBoxStyle.OkCancel Or MsgBoxStyle.Question) Then Exit Sub

        ' Supprimer seulement les clés ajoutées pour chacun des types de fichier
        EnleverMenuCtxIndexer(sMenuCtx_TypeFichierTxt)
        EnleverMenuCtxIndexer(sMenuCtx_TypeFichierDoc)
        If Me.chkTous.Checked Then EnleverMenuCtxIndexer(sMenuCtx_TypeFichierTous)
        EnleverMenuCtxIndexer(sMenuCtx_TypeDossier)

        ' bEnleverTypeFichier : enlever toute l'arbo HKCR\VBTextFinder
        bAjouterMenuContextuel(sMenuCtx_TypeFichierIdx, sMenuCtx_CleCmdIndexOuvrir,
            bEnlever:=True, bPrompt:=False, bEnleverTypeFichier:=True)

        ' Puis enlever le pointeur HKCR\.idx vers HKCR\VBTextFinder
        bAjouterTypeFichier(sMenuCtx_ExtFichierIdx, sMenuCtx_TypeFichierIdx,
            bEnlever:=True)

    End Sub

    Private Sub AjouterMenuCtxIndexer(sMenuCtx_TypeFichier$)

        Dim sCheminExe$ = Application.ExecutablePath
        Const bPrompt As Boolean = False
        Const sChemin$ = """%1"""
        bAjouterMenuContextuel(sMenuCtx_TypeFichier, sMenuCtx_CleCmdIndexer,
            bPrompt, , sMenuCtx_CleCmdIndexerDescription, sCheminExe, sChemin)

    End Sub

    Private Sub EnleverMenuCtxIndexer(sMenuCtx_TypeFichier$)

        bAjouterMenuContextuel(sMenuCtx_TypeFichier, sMenuCtx_CleCmdIndexer,
            bEnlever:=True, bPrompt:=False)

    End Sub

#End Region

End Class