<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVBTextFinder : Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()

        'Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()
        InitialiserFenetre()
    End Sub
    'La méthode substituée Dispose du formulaire pour nettoyer la liste des composants.
    Protected Overloads Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Requis par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer
    Friend ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TxtMot As System.Windows.Forms.ComboBox
    Friend WithEvents CmdInterrompre As System.Windows.Forms.Button
    Friend WithEvents LstTypeAffichResult As System.Windows.Forms.ListBox
    Friend WithEvents TxtResultat As System.Windows.Forms.TextBox
    Friend WithEvents CmdChercher As System.Windows.Forms.Button
    Friend WithEvents CmdChoisirFichierDoc As System.Windows.Forms.Button
    Friend WithEvents TxtCheminDocument As System.Windows.Forms.TextBox
    Friend WithEvents CmdAjouterDocument As System.Windows.Forms.Button
    Friend WithEvents LblCheminDoc As System.Windows.Forms.Label
    Friend WithEvents LblPresentation As System.Windows.Forms.Label
    Friend WithEvents LlbMot As System.Windows.Forms.Label
    Friend WithEvents LblAvancement As System.Windows.Forms.Label
    Friend WithEvents cmdAjouterMenuCtx As System.Windows.Forms.Button
    Friend WithEvents cmdEnleverMenuCtx As System.Windows.Forms.Button
 ' "Open"
    Friend WithEvents vsbZoomParag As System.Windows.Forms.VScrollBar
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Il peut être modifié à l'aide du Concepteur Windows Form.
    'Ne pas le modifier à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVBTextFinder))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CmdInterrompre = New System.Windows.Forms.Button()
        Me.LstTypeAffichResult = New System.Windows.Forms.ListBox()
        Me.TxtResultat = New System.Windows.Forms.TextBox()
        Me.CmdChercher = New System.Windows.Forms.Button()
        Me.CmdChoisirFichierDoc = New System.Windows.Forms.Button()
        Me.TxtCheminDocument = New System.Windows.Forms.TextBox()
        Me.CmdAjouterDocument = New System.Windows.Forms.Button()
        Me.vsbZoomParag = New System.Windows.Forms.VScrollBar()
        Me.cmdAjouterMenuCtx = New System.Windows.Forms.Button()
        Me.cmdEnleverMenuCtx = New System.Windows.Forms.Button()
        Me.chkTous = New System.Windows.Forms.CheckBox()
        Me.chkAfficherInfoResultat = New System.Windows.Forms.CheckBox()
        Me.chkMotsDico = New System.Windows.Forms.CheckBox()
        Me.chkMotsCourants = New System.Windows.Forms.CheckBox()
        Me.lstTypeIndex = New System.Windows.Forms.ListBox()
        Me.tbCodesLangues = New System.Windows.Forms.TextBox()
        Me.chkListeMots = New System.Windows.Forms.CheckBox()
        Me.mtbNbMotsCles = New System.Windows.Forms.MaskedTextBox()
        Me.lbCodesLangues = New System.Windows.Forms.ListBox()
        Me.tbCodeLangue = New System.Windows.Forms.TextBox()
        Me.chkNumeriques = New System.Windows.Forms.CheckBox()
        Me.cmdListeDoc = New System.Windows.Forms.Button()
        Me.chkUnicode = New System.Windows.Forms.CheckBox()
        Me.chkAccents = New System.Windows.Forms.CheckBox()
        Me.cmdExporterTxt = New System.Windows.Forms.Button()
        Me.cmdNavigExterne = New System.Windows.Forms.Button()
        Me.chkHtmlGras = New System.Windows.Forms.CheckBox()
        Me.chkHtmlCouleurs = New System.Windows.Forms.CheckBox()
        Me.tbCouleursHtml = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.chkChapitrage = New System.Windows.Forms.CheckBox()
        Me.tbChapitrage = New System.Windows.Forms.TextBox()
        Me.chkAfficherChapitreIndex = New System.Windows.Forms.CheckBox()
        Me.cmdListeDocHtml = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.chkAfficherTiret = New System.Windows.Forms.CheckBox()
        Me.chkNumerotationGlobale = New System.Windows.Forms.CheckBox()
        Me.chkAfficherInfoDoc = New System.Windows.Forms.CheckBox()
        Me.chkAfficherNumOccur = New System.Windows.Forms.CheckBox()
        Me.chkAfficherNumPhrase = New System.Windows.Forms.CheckBox()
        Me.chkAfficherNumParag = New System.Windows.Forms.CheckBox()
        Me.cmdGlossaire = New System.Windows.Forms.Button()
        Me.chkUnicodeVerif = New System.Windows.Forms.CheckBox()
        Me.TxtMot = New System.Windows.Forms.ComboBox()
        Me.LblCheminDoc = New System.Windows.Forms.Label()
        Me.LblPresentation = New System.Windows.Forms.Label()
        Me.LlbMot = New System.Windows.Forms.Label()
        Me.LblAvancement = New System.Windows.Forms.Label()
        Me.tcOnglets = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPageWeb = New System.Windows.Forms.TabPage()
        Me.wbResultat = New System.Windows.Forms.WebBrowser()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblCodeLangIndex = New System.Windows.Forms.Label()
        Me.lblNbMotsCles = New System.Windows.Forms.Label()
        Me.LblTypeIndex = New System.Windows.Forms.Label()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.lblMenuCtx = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.tcOnglets.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPageWeb.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.SuspendLayout()
        '
        'CmdInterrompre
        '
        Me.CmdInterrompre.BackColor = System.Drawing.SystemColors.Control
        Me.CmdInterrompre.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdInterrompre.Enabled = False
        Me.CmdInterrompre.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdInterrompre.Location = New System.Drawing.Point(3, 4)
        Me.CmdInterrompre.Name = "CmdInterrompre"
        Me.CmdInterrompre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdInterrompre.Size = New System.Drawing.Size(72, 25)
        Me.CmdInterrompre.TabIndex = 6
        Me.CmdInterrompre.Text = "Interrompre"
        Me.ToolTip1.SetToolTip(Me.CmdInterrompre, "Interrompre l'opération en cours")
        Me.CmdInterrompre.UseVisualStyleBackColor = False
        '
        'LstTypeAffichResult
        '
        Me.LstTypeAffichResult.BackColor = System.Drawing.SystemColors.Window
        Me.LstTypeAffichResult.Cursor = System.Windows.Forms.Cursors.Default
        Me.LstTypeAffichResult.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LstTypeAffichResult.Location = New System.Drawing.Point(8, 20)
        Me.LstTypeAffichResult.Name = "LstTypeAffichResult"
        Me.LstTypeAffichResult.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LstTypeAffichResult.Size = New System.Drawing.Size(113, 69)
        Me.LstTypeAffichResult.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.LstTypeAffichResult, "Afficher les paragraphes trouvés ou bien seulement les phrases")
        '
        'TxtResultat
        '
        Me.TxtResultat.AcceptsReturn = True
        Me.TxtResultat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtResultat.BackColor = System.Drawing.SystemColors.Window
        Me.TxtResultat.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtResultat.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtResultat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtResultat.Location = New System.Drawing.Point(27, 99)
        Me.TxtResultat.MaxLength = 0
        Me.TxtResultat.Multiline = True
        Me.TxtResultat.Name = "TxtResultat"
        Me.TxtResultat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtResultat.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TxtResultat.Size = New System.Drawing.Size(652, 272)
        Me.TxtResultat.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.TxtResultat, "Résultats de recherche : double-clic pour activer le mode hypertexte")
        '
        'CmdChercher
        '
        Me.CmdChercher.BackColor = System.Drawing.SystemColors.Control
        Me.CmdChercher.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdChercher.Enabled = False
        Me.CmdChercher.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdChercher.Location = New System.Drawing.Point(405, 68)
        Me.CmdChercher.Name = "CmdChercher"
        Me.CmdChercher.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdChercher.Size = New System.Drawing.Size(73, 25)
        Me.CmdChercher.TabIndex = 5
        Me.CmdChercher.Text = "Chercher"
        Me.ToolTip1.SetToolTip(Me.CmdChercher, "Chercher le mot ou les expressions entre guillemets dans l'index de VBTextFinder")
        Me.CmdChercher.UseVisualStyleBackColor = False
        '
        'CmdChoisirFichierDoc
        '
        Me.CmdChoisirFichierDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CmdChoisirFichierDoc.BackColor = System.Drawing.SystemColors.Control
        Me.CmdChoisirFichierDoc.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdChoisirFichierDoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdChoisirFichierDoc.Location = New System.Drawing.Point(649, 31)
        Me.CmdChoisirFichierDoc.Name = "CmdChoisirFichierDoc"
        Me.CmdChoisirFichierDoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdChoisirFichierDoc.Size = New System.Drawing.Size(25, 25)
        Me.CmdChoisirFichierDoc.TabIndex = 1
        Me.CmdChoisirFichierDoc.Text = "..."
        Me.ToolTip1.SetToolTip(Me.CmdChoisirFichierDoc, "Choisir un fichier texte de type document Bloc-notes Windows ou bien un document " & _
        "convertible")
        Me.CmdChoisirFichierDoc.UseVisualStyleBackColor = False
        '
        'TxtCheminDocument
        '
        Me.TxtCheminDocument.AcceptsReturn = True
        Me.TxtCheminDocument.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCheminDocument.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCheminDocument.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxtCheminDocument.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtCheminDocument.Location = New System.Drawing.Point(22, 34)
        Me.TxtCheminDocument.MaxLength = 0
        Me.TxtCheminDocument.Name = "TxtCheminDocument"
        Me.TxtCheminDocument.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtCheminDocument.Size = New System.Drawing.Size(621, 20)
        Me.TxtCheminDocument.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.TxtCheminDocument, "Chemin des documents à indexer (double-clic pour basculer le filtre de sélection " & _
        ": *.txt, *.doc, *.hmt?)")
        '
        'CmdAjouterDocument
        '
        Me.CmdAjouterDocument.BackColor = System.Drawing.SystemColors.Control
        Me.CmdAjouterDocument.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmdAjouterDocument.Enabled = False
        Me.CmdAjouterDocument.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CmdAjouterDocument.Location = New System.Drawing.Point(22, 72)
        Me.CmdAjouterDocument.Name = "CmdAjouterDocument"
        Me.CmdAjouterDocument.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CmdAjouterDocument.Size = New System.Drawing.Size(144, 25)
        Me.CmdAjouterDocument.TabIndex = 3
        Me.CmdAjouterDocument.Text = "Ajouter le(s) document(s)"
        Me.ToolTip1.SetToolTip(Me.CmdAjouterDocument, "Ajouter le ou les documents à l'index")
        Me.CmdAjouterDocument.UseVisualStyleBackColor = False
        '
        'vsbZoomParag
        '
        Me.vsbZoomParag.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.vsbZoomParag.LargeChange = 5
        Me.vsbZoomParag.Location = New System.Drawing.Point(8, 99)
        Me.vsbZoomParag.Maximum = 104
        Me.vsbZoomParag.Minimum = -1
        Me.vsbZoomParag.Name = "vsbZoomParag"
        Me.vsbZoomParag.Size = New System.Drawing.Size(16, 272)
        Me.vsbZoomParag.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.vsbZoomParag, "Afficher les paragraphes trouvés ou bien seulement les phrases")
        Me.vsbZoomParag.Value = -1
        '
        'cmdAjouterMenuCtx
        '
        Me.cmdAjouterMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAjouterMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAjouterMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAjouterMenuCtx.Location = New System.Drawing.Point(133, 13)
        Me.cmdAjouterMenuCtx.Name = "cmdAjouterMenuCtx"
        Me.cmdAjouterMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAjouterMenuCtx.Size = New System.Drawing.Size(25, 25)
        Me.cmdAjouterMenuCtx.TabIndex = 34
        Me.cmdAjouterMenuCtx.Text = "+"
        Me.ToolTip1.SetToolTip(Me.cmdAjouterMenuCtx, "Ajouter les menus contextuels pour indexer directement un fichier depuis l'explor" & _
        "ateur de fichiers (lancer l'appli. en tant qu'admin.)")
        Me.cmdAjouterMenuCtx.UseVisualStyleBackColor = False
        '
        'cmdEnleverMenuCtx
        '
        Me.cmdEnleverMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnleverMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnleverMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnleverMenuCtx.Location = New System.Drawing.Point(164, 13)
        Me.cmdEnleverMenuCtx.Name = "cmdEnleverMenuCtx"
        Me.cmdEnleverMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnleverMenuCtx.Size = New System.Drawing.Size(25, 25)
        Me.cmdEnleverMenuCtx.TabIndex = 35
        Me.cmdEnleverMenuCtx.Text = "-"
        Me.ToolTip1.SetToolTip(Me.cmdEnleverMenuCtx, "Enlever les menus contextuels (lancer l'appli. en tant qu'admin.)")
        Me.cmdEnleverMenuCtx.UseVisualStyleBackColor = False
        '
        'chkTous
        '
        Me.chkTous.AutoSize = True
        Me.chkTous.Checked = True
        Me.chkTous.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkTous.Location = New System.Drawing.Point(207, 18)
        Me.chkTous.Name = "chkTous"
        Me.chkTous.Size = New System.Drawing.Size(37, 17)
        Me.chkTous.TabIndex = 36
        Me.chkTous.Text = "*.*"
        Me.ToolTip1.SetToolTip(Me.chkTous, "Ajouter/enlever le menu contextuel aussi pour tous les documents (*.*)")
        Me.chkTous.UseVisualStyleBackColor = True
        '
        'chkAfficherInfoResultat
        '
        Me.chkAfficherInfoResultat.AutoSize = True
        Me.chkAfficherInfoResultat.Checked = True
        Me.chkAfficherInfoResultat.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAfficherInfoResultat.Location = New System.Drawing.Point(145, 20)
        Me.chkAfficherInfoResultat.Name = "chkAfficherInfoResultat"
        Me.chkAfficherInfoResultat.Size = New System.Drawing.Size(83, 17)
        Me.chkAfficherInfoResultat.TabIndex = 17
        Me.chkAfficherInfoResultat.Text = "Informations"
        Me.ToolTip1.SetToolTip(Me.chkAfficherInfoResultat, "Afficher les informations concernant les résultats de recherche (selon les option" & _
        "s choisies dans la config.)")
        Me.chkAfficherInfoResultat.UseVisualStyleBackColor = True
        '
        'chkMotsDico
        '
        Me.chkMotsDico.AutoSize = True
        Me.chkMotsDico.Checked = True
        Me.chkMotsDico.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMotsDico.Location = New System.Drawing.Point(15, 12)
        Me.chkMotsDico.Name = "chkMotsDico"
        Me.chkMotsDico.Size = New System.Drawing.Size(72, 17)
        Me.chkMotsDico.TabIndex = 0
        Me.chkMotsDico.Text = "Mots dico"
        Me.ToolTip1.SetToolTip(Me.chkMotsDico, "Inclure les mots du dictionnaire dans l'index")
        Me.chkMotsDico.UseVisualStyleBackColor = True
        '
        'chkMotsCourants
        '
        Me.chkMotsCourants.AutoSize = True
        Me.chkMotsCourants.Checked = True
        Me.chkMotsCourants.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMotsCourants.Location = New System.Drawing.Point(15, 32)
        Me.chkMotsCourants.Name = "chkMotsCourants"
        Me.chkMotsCourants.Size = New System.Drawing.Size(93, 17)
        Me.chkMotsCourants.TabIndex = 1
        Me.chkMotsCourants.Text = "Mots courants"
        Me.ToolTip1.SetToolTip(Me.chkMotsCourants, "Inclure les mots courants (de, la, les, ...) dans l'index")
        Me.chkMotsCourants.UseVisualStyleBackColor = True
        '
        'lstTypeIndex
        '
        Me.lstTypeIndex.BackColor = System.Drawing.SystemColors.Window
        Me.lstTypeIndex.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstTypeIndex.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstTypeIndex.Location = New System.Drawing.Point(21, 32)
        Me.lstTypeIndex.Name = "lstTypeIndex"
        Me.lstTypeIndex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstTypeIndex.Size = New System.Drawing.Size(104, 147)
        Me.lstTypeIndex.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.lstTypeIndex, "Double-clic pour créer le document index sous Word selon le type sélectionné")
        '
        'tbCodesLangues
        '
        Me.tbCodesLangues.Location = New System.Drawing.Point(237, 62)
        Me.tbCodesLangues.Name = "tbCodesLangues"
        Me.tbCodesLangues.Size = New System.Drawing.Size(160, 20)
        Me.tbCodesLangues.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.tbCodesLangues, "Définir la liste des codes langues à choisir directement")
        '
        'chkListeMots
        '
        Me.chkListeMots.AutoSize = True
        Me.chkListeMots.Location = New System.Drawing.Point(163, 32)
        Me.chkListeMots.Name = "chkListeMots"
        Me.chkListeMots.Size = New System.Drawing.Size(73, 17)
        Me.chkListeMots.TabIndex = 1
        Me.chkListeMots.Text = "Liste mots"
        Me.ToolTip1.SetToolTip(Me.chkListeMots, "Simple liste de mots pour l'index (ne pas afficher la fréquence ni les codes docu" & _
        "ments)")
        Me.chkListeMots.UseVisualStyleBackColor = True
        '
        'mtbNbMotsCles
        '
        Me.mtbNbMotsCles.Location = New System.Drawing.Point(384, 33)
        Me.mtbNbMotsCles.Mask = "9999"
        Me.mtbNbMotsCles.Name = "mtbNbMotsCles"
        Me.mtbNbMotsCles.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mtbNbMotsCles.Size = New System.Drawing.Size(32, 20)
        Me.mtbNbMotsCles.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.mtbNbMotsCles, "Nombre de mots clés")
        '
        'lbCodesLangues
        '
        Me.lbCodesLangues.FormattingEnabled = True
        Me.lbCodesLangues.Location = New System.Drawing.Point(147, 62)
        Me.lbCodesLangues.Name = "lbCodesLangues"
        Me.lbCodesLangues.Size = New System.Drawing.Size(55, 56)
        Me.lbCodesLangues.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.lbCodesLangues, "Choisir le code langue directement dans la liste")
        '
        'tbCodeLangue
        '
        Me.tbCodeLangue.Location = New System.Drawing.Point(77, 62)
        Me.tbCodeLangue.Name = "tbCodeLangue"
        Me.tbCodeLangue.Size = New System.Drawing.Size(42, 20)
        Me.tbCodeLangue.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.tbCodeLangue, "Lire les fichiers Dico et MotsCourants avec un suffixe pour le code langue (ex.: " & _
        "Dico_Fr.txt)")
        '
        'chkNumeriques
        '
        Me.chkNumeriques.AutoSize = True
        Me.chkNumeriques.Checked = True
        Me.chkNumeriques.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkNumeriques.Location = New System.Drawing.Point(163, 55)
        Me.chkNumeriques.Name = "chkNumeriques"
        Me.chkNumeriques.Size = New System.Drawing.Size(82, 17)
        Me.chkNumeriques.TabIndex = 2
        Me.chkNumeriques.Text = "Numériques"
        Me.ToolTip1.SetToolTip(Me.chkNumeriques, "Inclure les numériques dans l'index")
        Me.chkNumeriques.UseVisualStyleBackColor = True
        '
        'cmdListeDoc
        '
        Me.cmdListeDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdListeDoc.Location = New System.Drawing.Point(654, 15)
        Me.cmdListeDoc.Name = "cmdListeDoc"
        Me.cmdListeDoc.Size = New System.Drawing.Size(25, 25)
        Me.cmdListeDoc.TabIndex = 18
        Me.cmdListeDoc.Text = "?"
        Me.ToolTip1.SetToolTip(Me.cmdListeDoc, "Rappeler la liste des documents indexés")
        Me.cmdListeDoc.UseVisualStyleBackColor = True
        '
        'chkUnicode
        '
        Me.chkUnicode.AutoSize = True
        Me.chkUnicode.Location = New System.Drawing.Point(21, 61)
        Me.chkUnicode.Name = "chkUnicode"
        Me.chkUnicode.Size = New System.Drawing.Size(66, 17)
        Me.chkUnicode.TabIndex = 37
        Me.chkUnicode.Text = "Unicode"
        Me.ToolTip1.SetToolTip(Me.chkUnicode, "Utiliser l'encodage Unicode UTF-8 (au lieu du code page latin 1252), pour conserv" & _
        "er par exemple les caractères grecs (relancer l'application et réindexer).")
        Me.chkUnicode.UseVisualStyleBackColor = True
        '
        'chkAccents
        '
        Me.chkAccents.AutoSize = True
        Me.chkAccents.Location = New System.Drawing.Point(21, 84)
        Me.chkAccents.Name = "chkAccents"
        Me.chkAccents.Size = New System.Drawing.Size(65, 17)
        Me.chkAccents.TabIndex = 38
        Me.chkAccents.Text = "Accents"
        Me.ToolTip1.SetToolTip(Me.chkAccents, "Tenir compte des accents lors de l'indexation et de la recherche (relancer l'appl" & _
        "ication et réindexer).")
        Me.chkAccents.UseVisualStyleBackColor = True
        '
        'cmdExporterTxt
        '
        Me.cmdExporterTxt.Location = New System.Drawing.Point(17, 12)
        Me.cmdExporterTxt.Name = "cmdExporterTxt"
        Me.cmdExporterTxt.Size = New System.Drawing.Size(49, 25)
        Me.cmdExporterTxt.TabIndex = 1
        Me.cmdExporterTxt.Text = "Texte"
        Me.ToolTip1.SetToolTip(Me.cmdExporterTxt, "Exporter le contenu en mode texte simple")
        Me.cmdExporterTxt.UseVisualStyleBackColor = True
        '
        'cmdNavigExterne
        '
        Me.cmdNavigExterne.Location = New System.Drawing.Point(77, 12)
        Me.cmdNavigExterne.Name = "cmdNavigExterne"
        Me.cmdNavigExterne.Size = New System.Drawing.Size(49, 25)
        Me.cmdNavigExterne.TabIndex = 2
        Me.cmdNavigExterne.Text = "Html"
        Me.ToolTip1.SetToolTip(Me.cmdNavigExterne, "Consulter la page web dans le navigateur par défaut (beaucoup plus rapide qu'un c" & _
        "opié/collé lorsque la page est très volumineuse)")
        Me.cmdNavigExterne.UseVisualStyleBackColor = True
        '
        'chkHtmlGras
        '
        Me.chkHtmlGras.AutoSize = True
        Me.chkHtmlGras.Location = New System.Drawing.Point(15, 16)
        Me.chkHtmlGras.Name = "chkHtmlGras"
        Me.chkHtmlGras.Size = New System.Drawing.Size(48, 17)
        Me.chkHtmlGras.TabIndex = 39
        Me.chkHtmlGras.Text = "Gras"
        Me.ToolTip1.SetToolTip(Me.chkHtmlGras, "Mettre en évidence les occurrences en gras")
        Me.chkHtmlGras.UseVisualStyleBackColor = True
        '
        'chkHtmlCouleurs
        '
        Me.chkHtmlCouleurs.AutoSize = True
        Me.chkHtmlCouleurs.Checked = True
        Me.chkHtmlCouleurs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHtmlCouleurs.Location = New System.Drawing.Point(15, 39)
        Me.chkHtmlCouleurs.Name = "chkHtmlCouleurs"
        Me.chkHtmlCouleurs.Size = New System.Drawing.Size(67, 17)
        Me.chkHtmlCouleurs.TabIndex = 40
        Me.chkHtmlCouleurs.Text = "Couleurs"
        Me.ToolTip1.SetToolTip(Me.chkHtmlCouleurs, "Mettre en évidence les occurrences en couleurs")
        Me.chkHtmlCouleurs.UseVisualStyleBackColor = True
        '
        'tbCouleursHtml
        '
        Me.tbCouleursHtml.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbCouleursHtml.Location = New System.Drawing.Point(89, 39)
        Me.tbCouleursHtml.Name = "tbCouleursHtml"
        Me.tbCouleursHtml.Size = New System.Drawing.Size(578, 20)
        Me.tbCouleursHtml.TabIndex = 41
        Me.ToolTip1.SetToolTip(Me.tbCouleursHtml, "Liste des couleurs html des occurrences séparées par ; (n° couleur = modulo n° ex" & _
        "pression)")
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.chkHtmlCouleurs)
        Me.Panel2.Controls.Add(Me.tbCouleursHtml)
        Me.Panel2.Controls.Add(Me.chkHtmlGras)
        Me.Panel2.Location = New System.Drawing.Point(5, 117)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(687, 83)
        Me.Panel2.TabIndex = 43
        Me.ToolTip1.SetToolTip(Me.Panel2, "Options de l'affichage Html")
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.chkChapitrage)
        Me.Panel3.Controls.Add(Me.tbChapitrage)
        Me.Panel3.Location = New System.Drawing.Point(5, 222)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(687, 83)
        Me.Panel3.TabIndex = 44
        Me.ToolTip1.SetToolTip(Me.Panel3, "Options de chapitrage")
        '
        'chkChapitrage
        '
        Me.chkChapitrage.AutoSize = True
        Me.chkChapitrage.Checked = True
        Me.chkChapitrage.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkChapitrage.Location = New System.Drawing.Point(15, 39)
        Me.chkChapitrage.Name = "chkChapitrage"
        Me.chkChapitrage.Size = New System.Drawing.Size(70, 17)
        Me.chkChapitrage.TabIndex = 40
        Me.chkChapitrage.Text = "Chapitres"
        Me.ToolTip1.SetToolTip(Me.chkChapitrage, "Détecter les chapitres")
        Me.chkChapitrage.UseVisualStyleBackColor = True
        '
        'tbChapitrage
        '
        Me.tbChapitrage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbChapitrage.BackColor = System.Drawing.SystemColors.Window
        Me.tbChapitrage.Location = New System.Drawing.Point(89, 39)
        Me.tbChapitrage.Name = "tbChapitrage"
        Me.tbChapitrage.ReadOnly = True
        Me.tbChapitrage.Size = New System.Drawing.Size(578, 20)
        Me.tbChapitrage.TabIndex = 41
        Me.ToolTip1.SetToolTip(Me.tbChapitrage, "Liste des types de chapitres à détecter, avec le code a afficher, séparer par des" & _
        " ; Signe - pour ignorer un début de phrase détecté.")
        '
        'chkAfficherChapitreIndex
        '
        Me.chkAfficherChapitreIndex.AutoSize = True
        Me.chkAfficherChapitreIndex.Enabled = False
        Me.chkAfficherChapitreIndex.Location = New System.Drawing.Point(463, 32)
        Me.chkAfficherChapitreIndex.Name = "chkAfficherChapitreIndex"
        Me.chkAfficherChapitreIndex.Size = New System.Drawing.Size(70, 17)
        Me.chkAfficherChapitreIndex.TabIndex = 44
        Me.chkAfficherChapitreIndex.Text = "Chapitres"
        Me.ToolTip1.SetToolTip(Me.chkAfficherChapitreIndex, "Afficher les chapitres avec les codes documents")
        Me.chkAfficherChapitreIndex.UseVisualStyleBackColor = True
        '
        'cmdListeDocHtml
        '
        Me.cmdListeDocHtml.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdListeDocHtml.Location = New System.Drawing.Point(656, 12)
        Me.cmdListeDocHtml.Name = "cmdListeDocHtml"
        Me.cmdListeDocHtml.Size = New System.Drawing.Size(25, 25)
        Me.cmdListeDocHtml.TabIndex = 3
        Me.cmdListeDocHtml.Text = "?"
        Me.ToolTip1.SetToolTip(Me.cmdListeDocHtml, "Rappeler la liste des documents indexés")
        Me.cmdListeDocHtml.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.chkAfficherTiret)
        Me.Panel4.Controls.Add(Me.chkNumerotationGlobale)
        Me.Panel4.Controls.Add(Me.chkAfficherInfoDoc)
        Me.Panel4.Controls.Add(Me.chkAfficherNumOccur)
        Me.Panel4.Controls.Add(Me.chkAfficherNumPhrase)
        Me.Panel4.Controls.Add(Me.chkAfficherNumParag)
        Me.Panel4.Location = New System.Drawing.Point(313, 13)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(290, 88)
        Me.Panel4.TabIndex = 46
        Me.ToolTip1.SetToolTip(Me.Panel4, "Options pour la présentation des résultats de recherche")
        '
        'chkAfficherTiret
        '
        Me.chkAfficherTiret.AutoSize = True
        Me.chkAfficherTiret.Checked = True
        Me.chkAfficherTiret.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAfficherTiret.Location = New System.Drawing.Point(160, 58)
        Me.chkAfficherTiret.Name = "chkAfficherTiret"
        Me.chkAfficherTiret.Size = New System.Drawing.Size(47, 17)
        Me.chkAfficherTiret.TabIndex = 23
        Me.chkAfficherTiret.Text = "Tiret"
        Me.ToolTip1.SetToolTip(Me.chkAfficherTiret, "Afficher un tiret devant chaque occurrences trouvées")
        Me.chkAfficherTiret.UseVisualStyleBackColor = True
        '
        'chkNumerotationGlobale
        '
        Me.chkNumerotationGlobale.AutoSize = True
        Me.chkNumerotationGlobale.Checked = True
        Me.chkNumerotationGlobale.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkNumerotationGlobale.Location = New System.Drawing.Point(17, 58)
        Me.chkNumerotationGlobale.Name = "chkNumerotationGlobale"
        Me.chkNumerotationGlobale.Size = New System.Drawing.Size(126, 17)
        Me.chkNumerotationGlobale.TabIndex = 22
        Me.chkNumerotationGlobale.Text = "Numérotation globale"
        Me.ToolTip1.SetToolTip(Me.chkNumerotationGlobale, "Numérotation globale sur l'ensemble des documents (ou sinon propre à chaque docum" & _
        "ent)")
        Me.chkNumerotationGlobale.UseVisualStyleBackColor = True
        '
        'chkAfficherInfoDoc
        '
        Me.chkAfficherInfoDoc.AutoSize = True
        Me.chkAfficherInfoDoc.Location = New System.Drawing.Point(160, 35)
        Me.chkAfficherInfoDoc.Name = "chkAfficherInfoDoc"
        Me.chkAfficherInfoDoc.Size = New System.Drawing.Size(75, 17)
        Me.chkAfficherInfoDoc.TabIndex = 21
        Me.chkAfficherInfoDoc.Text = "Document"
        Me.ToolTip1.SetToolTip(Me.chkAfficherInfoDoc, "Afficher les informations concernant le document")
        Me.chkAfficherInfoDoc.UseVisualStyleBackColor = True
        '
        'chkAfficherNumOccur
        '
        Me.chkAfficherNumOccur.AutoSize = True
        Me.chkAfficherNumOccur.Location = New System.Drawing.Point(160, 12)
        Me.chkAfficherNumOccur.Name = "chkAfficherNumOccur"
        Me.chkAfficherNumOccur.Size = New System.Drawing.Size(100, 17)
        Me.chkAfficherNumOccur.TabIndex = 20
        Me.chkAfficherNumOccur.Text = "N° occurrences"
        Me.ToolTip1.SetToolTip(Me.chkAfficherNumOccur, "Afficher le n° des occurrences trouvées")
        Me.chkAfficherNumOccur.UseVisualStyleBackColor = True
        '
        'chkAfficherNumPhrase
        '
        Me.chkAfficherNumPhrase.AutoSize = True
        Me.chkAfficherNumPhrase.Checked = True
        Me.chkAfficherNumPhrase.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAfficherNumPhrase.Location = New System.Drawing.Point(17, 35)
        Me.chkAfficherNumPhrase.Name = "chkAfficherNumPhrase"
        Me.chkAfficherNumPhrase.Size = New System.Drawing.Size(88, 17)
        Me.chkAfficherNumPhrase.TabIndex = 2
        Me.chkAfficherNumPhrase.Text = "N° de phrase"
        Me.ToolTip1.SetToolTip(Me.chkAfficherNumPhrase, "Afficher les n° de phrase")
        Me.chkAfficherNumPhrase.UseVisualStyleBackColor = True
        '
        'chkAfficherNumParag
        '
        Me.chkAfficherNumParag.AutoSize = True
        Me.chkAfficherNumParag.Checked = True
        Me.chkAfficherNumParag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAfficherNumParag.Location = New System.Drawing.Point(17, 12)
        Me.chkAfficherNumParag.Name = "chkAfficherNumParag"
        Me.chkAfficherNumParag.Size = New System.Drawing.Size(110, 17)
        Me.chkAfficherNumParag.TabIndex = 1
        Me.chkAfficherNumParag.Text = "N° de paragraphe"
        Me.ToolTip1.SetToolTip(Me.chkAfficherNumParag, "Afficher les n° de paragraphe")
        Me.chkAfficherNumParag.UseVisualStyleBackColor = True
        '
        'cmdGlossaire
        '
        Me.cmdGlossaire.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGlossaire.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGlossaire.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGlossaire.Location = New System.Drawing.Point(35, 197)
        Me.cmdGlossaire.Name = "cmdGlossaire"
        Me.cmdGlossaire.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGlossaire.Size = New System.Drawing.Size(72, 25)
        Me.cmdGlossaire.TabIndex = 45
        Me.cmdGlossaire.Text = "Glossaire"
        Me.ToolTip1.SetToolTip(Me.cmdGlossaire, "Créer un glossaire des mots hors du dictionnaire de MS-Word (selon la langue choi" & _
        "sie, et pour un seul document indiqué dans l'onglet Indexation, et non une liste" & _
        " de documents)")
        Me.cmdGlossaire.UseVisualStyleBackColor = False
        '
        'chkUnicodeVerif
        '
        Me.chkUnicodeVerif.AutoSize = True
        Me.chkUnicodeVerif.Checked = True
        Me.chkUnicodeVerif.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUnicodeVerif.Location = New System.Drawing.Point(92, 61)
        Me.chkUnicodeVerif.Name = "chkUnicodeVerif"
        Me.chkUnicodeVerif.Size = New System.Drawing.Size(92, 17)
        Me.chkUnicodeVerif.TabIndex = 47
        Me.chkUnicodeVerif.Text = "Unicode vérif."
        Me.ToolTip1.SetToolTip(Me.chkUnicodeVerif, "Vérifier la présence ou l'absence de caractère unicode lors de l'indexation du do" & _
        "cument (cela peut ralentir l'étape de conversion .doc vers .txt).")
        Me.chkUnicodeVerif.UseVisualStyleBackColor = True
        '
        'TxtMot
        '
        Me.TxtMot.BackColor = System.Drawing.SystemColors.Window
        Me.TxtMot.Cursor = System.Windows.Forms.Cursors.Default
        Me.TxtMot.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxtMot.Location = New System.Drawing.Point(145, 68)
        Me.TxtMot.Name = "TxtMot"
        Me.TxtMot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxtMot.Size = New System.Drawing.Size(243, 21)
        Me.TxtMot.TabIndex = 4
        '
        'LblCheminDoc
        '
        Me.LblCheminDoc.BackColor = System.Drawing.SystemColors.Control
        Me.LblCheminDoc.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblCheminDoc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblCheminDoc.Location = New System.Drawing.Point(22, 18)
        Me.LblCheminDoc.Name = "LblCheminDoc"
        Me.LblCheminDoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblCheminDoc.Size = New System.Drawing.Size(184, 17)
        Me.LblCheminDoc.TabIndex = 15
        Me.LblCheminDoc.Text = "Chemin des documents à indexer"
        '
        'LblPresentation
        '
        Me.LblPresentation.BackColor = System.Drawing.SystemColors.Control
        Me.LblPresentation.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblPresentation.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblPresentation.Location = New System.Drawing.Point(5, 3)
        Me.LblPresentation.Name = "LblPresentation"
        Me.LblPresentation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblPresentation.Size = New System.Drawing.Size(80, 13)
        Me.LblPresentation.TabIndex = 13
        Me.LblPresentation.Text = "Présentation"
        '
        'LlbMot
        '
        Me.LlbMot.BackColor = System.Drawing.SystemColors.Control
        Me.LlbMot.Cursor = System.Windows.Forms.Cursors.Default
        Me.LlbMot.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LlbMot.Location = New System.Drawing.Point(14, -23)
        Me.LlbMot.Name = "LlbMot"
        Me.LlbMot.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LlbMot.Size = New System.Drawing.Size(97, 17)
        Me.LlbMot.TabIndex = 12
        Me.LlbMot.Text = "Mot à rechercher"
        '
        'LblAvancement
        '
        Me.LblAvancement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblAvancement.BackColor = System.Drawing.SystemColors.Control
        Me.LblAvancement.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblAvancement.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblAvancement.Location = New System.Drawing.Point(81, 4)
        Me.LblAvancement.Name = "LblAvancement"
        Me.LblAvancement.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblAvancement.Size = New System.Drawing.Size(629, 30)
        Me.LblAvancement.TabIndex = 10
        Me.LblAvancement.Text = "Avancement"
        '
        'tcOnglets
        '
        Me.tcOnglets.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tcOnglets.Controls.Add(Me.TabPage1)
        Me.tcOnglets.Controls.Add(Me.TabPageWeb)
        Me.tcOnglets.Controls.Add(Me.TabPage2)
        Me.tcOnglets.Controls.Add(Me.TabPage3)
        Me.tcOnglets.Controls.Add(Me.TabPage4)
        Me.tcOnglets.Location = New System.Drawing.Point(3, 37)
        Me.tcOnglets.Name = "tcOnglets"
        Me.tcOnglets.SelectedIndex = 0
        Me.tcOnglets.Size = New System.Drawing.Size(707, 416)
        Me.tcOnglets.TabIndex = 39
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.cmdListeDoc)
        Me.TabPage1.Controls.Add(Me.chkAfficherInfoResultat)
        Me.TabPage1.Controls.Add(Me.TxtResultat)
        Me.TabPage1.Controls.Add(Me.LstTypeAffichResult)
        Me.TabPage1.Controls.Add(Me.vsbZoomParag)
        Me.TabPage1.Controls.Add(Me.LblPresentation)
        Me.TabPage1.Controls.Add(Me.LlbMot)
        Me.TabPage1.Controls.Add(Me.TxtMot)
        Me.TabPage1.Controls.Add(Me.CmdChercher)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(699, 390)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Recherche"
        Me.TabPage1.ToolTipText = "Page de recherche"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPageWeb
        '
        Me.TabPageWeb.Controls.Add(Me.cmdListeDocHtml)
        Me.TabPageWeb.Controls.Add(Me.cmdNavigExterne)
        Me.TabPageWeb.Controls.Add(Me.cmdExporterTxt)
        Me.TabPageWeb.Controls.Add(Me.wbResultat)
        Me.TabPageWeb.Location = New System.Drawing.Point(4, 22)
        Me.TabPageWeb.Name = "TabPageWeb"
        Me.TabPageWeb.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPageWeb.Size = New System.Drawing.Size(699, 390)
        Me.TabPageWeb.TabIndex = 4
        Me.TabPageWeb.Text = "Page html"
        Me.TabPageWeb.ToolTipText = "Page d'affichage illimité"
        Me.TabPageWeb.UseVisualStyleBackColor = True
        '
        'wbResultat
        '
        Me.wbResultat.AllowWebBrowserDrop = False
        Me.wbResultat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.wbResultat.CausesValidation = False
        Me.wbResultat.Location = New System.Drawing.Point(17, 47)
        Me.wbResultat.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbResultat.Name = "wbResultat"
        Me.wbResultat.ScriptErrorsSuppressed = True
        Me.wbResultat.Size = New System.Drawing.Size(664, 326)
        Me.wbResultat.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TxtCheminDocument)
        Me.TabPage2.Controls.Add(Me.LblCheminDoc)
        Me.TabPage2.Controls.Add(Me.CmdAjouterDocument)
        Me.TabPage2.Controls.Add(Me.CmdChoisirFichierDoc)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(699, 390)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Indexation"
        Me.TabPage2.ToolTipText = "Page d'indexation"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.cmdGlossaire)
        Me.TabPage3.Controls.Add(Me.chkAfficherChapitreIndex)
        Me.TabPage3.Controls.Add(Me.Panel1)
        Me.TabPage3.Controls.Add(Me.chkNumeriques)
        Me.TabPage3.Controls.Add(Me.mtbNbMotsCles)
        Me.TabPage3.Controls.Add(Me.chkListeMots)
        Me.TabPage3.Controls.Add(Me.lblNbMotsCles)
        Me.TabPage3.Controls.Add(Me.LblTypeIndex)
        Me.TabPage3.Controls.Add(Me.lstTypeIndex)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(699, 390)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Outils"
        Me.TabPage3.ToolTipText = "Page d'outils"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.tbCodesLangues)
        Me.Panel1.Controls.Add(Me.lblCodeLangIndex)
        Me.Panel1.Controls.Add(Me.tbCodeLangue)
        Me.Panel1.Controls.Add(Me.lbCodesLangues)
        Me.Panel1.Controls.Add(Me.chkMotsDico)
        Me.Panel1.Controls.Add(Me.chkMotsCourants)
        Me.Panel1.Location = New System.Drawing.Point(146, 90)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(421, 150)
        Me.Panel1.TabIndex = 43
        '
        'lblCodeLangIndex
        '
        Me.lblCodeLangIndex.AutoSize = True
        Me.lblCodeLangIndex.Location = New System.Drawing.Point(12, 62)
        Me.lblCodeLangIndex.Name = "lblCodeLangIndex"
        Me.lblCodeLangIndex.Size = New System.Drawing.Size(58, 13)
        Me.lblCodeLangIndex.TabIndex = 4
        Me.lblCodeLangIndex.Text = "Code lang."
        '
        'lblNbMotsCles
        '
        Me.lblNbMotsCles.AutoSize = True
        Me.lblNbMotsCles.Location = New System.Drawing.Point(307, 36)
        Me.lblNbMotsCles.Name = "lblNbMotsCles"
        Me.lblNbMotsCles.Size = New System.Drawing.Size(71, 13)
        Me.lblNbMotsCles.TabIndex = 8
        Me.lblNbMotsCles.Text = "Nb. mots clés"
        '
        'LblTypeIndex
        '
        Me.LblTypeIndex.BackColor = System.Drawing.SystemColors.Control
        Me.LblTypeIndex.Cursor = System.Windows.Forms.Cursors.Default
        Me.LblTypeIndex.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LblTypeIndex.Location = New System.Drawing.Point(18, 12)
        Me.LblTypeIndex.Name = "LblTypeIndex"
        Me.LblTypeIndex.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LblTypeIndex.Size = New System.Drawing.Size(80, 17)
        Me.LblTypeIndex.TabIndex = 41
        Me.LblTypeIndex.Text = "Type d'index"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.chkUnicodeVerif)
        Me.TabPage4.Controls.Add(Me.Panel4)
        Me.TabPage4.Controls.Add(Me.Panel3)
        Me.TabPage4.Controls.Add(Me.Panel2)
        Me.TabPage4.Controls.Add(Me.chkAccents)
        Me.TabPage4.Controls.Add(Me.chkUnicode)
        Me.TabPage4.Controls.Add(Me.chkTous)
        Me.TabPage4.Controls.Add(Me.cmdEnleverMenuCtx)
        Me.TabPage4.Controls.Add(Me.lblMenuCtx)
        Me.TabPage4.Controls.Add(Me.cmdAjouterMenuCtx)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(699, 390)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Config."
        Me.TabPage4.ToolTipText = "Page de configuration"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'lblMenuCtx
        '
        Me.lblMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.lblMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMenuCtx.Location = New System.Drawing.Point(17, 20)
        Me.lblMenuCtx.Name = "lblMenuCtx"
        Me.lblMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMenuCtx.Size = New System.Drawing.Size(110, 18)
        Me.lblMenuCtx.TabIndex = 10
        Me.lblMenuCtx.Text = "Menus contextuels :"
        '
        'frmVBTextFinder
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(722, 465)
        Me.Controls.Add(Me.tcOnglets)
        Me.Controls.Add(Me.LblAvancement)
        Me.Controls.Add(Me.CmdInterrompre)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.Name = "frmVBTextFinder"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "VBTextFinder : un moteur de recherche de mot dans son contexte"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.tcOnglets.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPageWeb.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents tcOnglets As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents lblMenuCtx As System.Windows.Forms.Label
    Friend WithEvents chkTous As System.Windows.Forms.CheckBox
    Friend WithEvents chkAfficherInfoResultat As System.Windows.Forms.CheckBox
    Friend WithEvents lblCodeLangIndex As System.Windows.Forms.Label
    Friend WithEvents tbCodesLangues As System.Windows.Forms.TextBox
    Friend WithEvents chkMotsCourants As System.Windows.Forms.CheckBox
    Friend WithEvents LblTypeIndex As System.Windows.Forms.Label
    Friend WithEvents chkMotsDico As System.Windows.Forms.CheckBox
    Friend WithEvents lstTypeIndex As System.Windows.Forms.ListBox
    Friend WithEvents chkListeMots As System.Windows.Forms.CheckBox
    Friend WithEvents lblNbMotsCles As System.Windows.Forms.Label
    Friend WithEvents mtbNbMotsCles As System.Windows.Forms.MaskedTextBox
    Friend WithEvents lbCodesLangues As System.Windows.Forms.ListBox
    Friend WithEvents tbCodeLangue As System.Windows.Forms.TextBox
    Friend WithEvents chkNumeriques As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdListeDoc As System.Windows.Forms.Button
    Friend WithEvents chkUnicode As System.Windows.Forms.CheckBox
    Friend WithEvents TabPageWeb As System.Windows.Forms.TabPage
    Friend WithEvents wbResultat As System.Windows.Forms.WebBrowser
    Friend WithEvents chkAccents As System.Windows.Forms.CheckBox
    Friend WithEvents cmdExporterTxt As System.Windows.Forms.Button
    Friend WithEvents cmdNavigExterne As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents chkHtmlCouleurs As System.Windows.Forms.CheckBox
    Friend WithEvents tbCouleursHtml As System.Windows.Forms.TextBox
    Friend WithEvents chkHtmlGras As System.Windows.Forms.CheckBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents chkChapitrage As System.Windows.Forms.CheckBox
    Friend WithEvents tbChapitrage As System.Windows.Forms.TextBox
    Friend WithEvents chkAfficherChapitreIndex As System.Windows.Forms.CheckBox
    Friend WithEvents cmdListeDocHtml As System.Windows.Forms.Button
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents chkAfficherNumOccur As System.Windows.Forms.CheckBox
    Friend WithEvents chkAfficherNumPhrase As System.Windows.Forms.CheckBox
    Friend WithEvents chkAfficherNumParag As System.Windows.Forms.CheckBox
    Friend WithEvents chkAfficherInfoDoc As System.Windows.Forms.CheckBox
    Friend WithEvents chkNumerotationGlobale As System.Windows.Forms.CheckBox
    Friend WithEvents cmdGlossaire As System.Windows.Forms.Button
    Friend WithEvents chkAfficherTiret As System.Windows.Forms.CheckBox
    Friend WithEvents chkUnicodeVerif As System.Windows.Forms.CheckBox
#End Region
End Class
