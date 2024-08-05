VERSION 5.00
Begin VB.Form frmVBTextFinder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBTextFinder : un moteur de recherche de mot dans son contexte, en VB6"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmVBTextFinder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox TxtMot 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton CmdInterrompre 
      Caption         =   "Interrompre"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      ToolTipText     =   "Interrompre l'op�ration en cours"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox LstTri 
      Height          =   450
      Left            =   6720
      TabIndex        =   7
      ToolTipText     =   "Double-clic pour cr�er le document index sous Word selon le tri s�lectionn�"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox LstTypeAffichResult 
      Height          =   1035
      Left            =   8520
      TabIndex        =   8
      ToolTipText     =   "Afficher les paragraphes trouv�s ou bien seulement les phrases"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TxtResultat 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      ToolTipText     =   "R�sultats de recherche : double-clic pour activer le mode hypertexte"
      Top             =   2280
      Width           =   10095
   End
   Begin VB.CommandButton CmdChercher 
      Caption         =   "Chercher"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Chercher le mot dans l'index de VBTextFinder"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TxtCodeDoc 
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   $"frmVBTextFinder.frx":030A
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CmdChoisirFichierDoc 
      Caption         =   "..."
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "Choisir un fichier texte de type document Bloc-notes Windows ou bien un document convertible"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TxtCheminDocument 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Chemin du document � indexer"
      Top             =   360
      Width           =   6015
   End
   Begin VB.CommandButton CmdAjouterDocument 
      Caption         =   "Ajouter le document"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8520
      TabIndex        =   3
      ToolTipText     =   "Ajouter le document � l'index (�ventuellement avec un code mn�monique � gauche)"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label LblCheminDoc 
      Caption         =   "Chemin du document � indexer"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblTri 
      Caption         =   "Tri"
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LblPresentation 
      Caption         =   "Pr�sentation"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LlbMot 
      Caption         =   "Mot � rechercher"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LblIndexDoc 
      Caption         =   "Code du document"
      Height          =   255
      Left            =   6720
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label LblAvancement 
      Caption         =   "Avancement"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   8175
   End
End
Attribute VB_Name = "frmVBTextFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VBTextFinder : un moteur de recherche de mot dans son contexte
' --------------------------------------------------------------
' www.vbfrance.com/code.aspx?ID=46695
' Documentation : VBTextFinder.html :
' http://patrice.dargenton.free.fr/CodesSources/VBTextFinder.html
' Par Patrice Dargenton : patrice.dargenton@free.fr
' http://patrice.dargenton.free.fr/index.html
' http://patrice.dargenton.free.fr/CodesSources/index.html
' Version 1.01 du 07/06/2008 (ne pas indexer 2 fois le m�me document)
' Version 1.00 du 18/05/2008
' --------------------------------------------------------------

' Conventions de nommage des variables :
' b pour Boolean (bool�en vrai ou faux)
' i pour Integer : %
' l pour Long : &
' r pour nombre R�el : Single! ou Double#
' a pour Array (tableau) : ()
' o pour Object (objet ou classe)
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)

' Objet moteur de recherche : c'est l'objet principal
'  dont ce formulaire est l'interface
Private m_oVBTxtFnd As New clsVBTextFinder

' Initialiser seulement la premi�re fois que la fen�tre est pr�te
Private m_bInit As Boolean

' Moins de 1% du code est distinct entre les versions VBA et VB6
'  en utilisant la compilation conditionnelle, on ne maintient qu'une seule version !
#If bVersionVBA Then ' Version VBA (Visual Basic pour Application) Excel et Word

Private Sub UserForm_Initialize()
    Initialiser
End Sub
Private Sub UserForm_Activate()
    Activer
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not bQuitter() Then Cancel = True
End Sub

#Else ' Version VB6

Private Sub Form_Load()
    Initialiser
End Sub
Private Sub Form_Activate()
    Activer
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not bQuitter() Then Cancel = True
End Sub

#End If

Private Sub Initialiser()
    
    ' On passe Me (le formulaire) pour changer Me.MousePointer
    '  (curseur de souris en sablier)
    m_oVBTxtFnd.Initialiser Me.LblAvancement, Me.LstTypeAffichResult, Me.LstTri, Me
    
End Sub

Private Sub Activer()

    If m_bInit Then Exit Sub
    m_bInit = True

    If m_oVBTxtFnd.m_bModeDirect Then
    
        Me.TxtCheminDocument = m_oVBTxtFnd.m_sCheminFichierTxtDirect
        DoEvents
        ' Convertir le fichier en .txt si son extension
        '  est celle d'un document convertible (.doc, .html ou .htm)
        If Not m_oVBTxtFnd.bConvertirDocEnTxt() Then Exit Sub
            'm_oVBTxtFnd.m_sCheminFichierTxtDirect) Then Exit sub
        Me.TxtCheminDocument = m_oVBTxtFnd.m_sCheminFichierTxtDirect
        DoEvents
        CmdAjouterDocument_Click
    
    Else

        If m_oVBTxtFnd.bLireIndex() Then
            m_oVBTxtFnd.ListerDocumentsIndexes Me.TxtResultat
        Else
            ' Fichier document trait� par d�faut, pour l'exemple
            Me.TxtCheminDocument = sLireCheminApplication & "\Proverbes.txt"
            Me.TxtCodeDoc = "PROV" ' Cl� du document par d�faut
        End If
    
    End If
    
    VerifierOperationsPossibles
    If Me.TxtMot.Enabled Then Me.TxtMot.SetFocus
    
End Sub

Private Function bQuitter() As Boolean
    
    bQuitter = m_oVBTxtFnd.bQuitter()

End Function

#If bVersionVBA Then

Private Sub LstTri_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

#Else

Private Sub LstTri_DblClick()

#End If

    ' Quitter si une op�ration est en cours
    If Me.CmdInterrompre.Enabled Then Exit Sub
    Me.CmdInterrompre.Enabled = True
    Me.CmdInterrompre.SetFocus
    m_oVBTxtFnd.CreerDocIndex Me.LstTri
    Me.CmdInterrompre.Enabled = False

End Sub

Private Sub CmdChoisirFichierDoc_Click()
    
    ' Gerer la bo�te de dialogue pour choisir un fichier document Word � indexer
    
    ' Handle de la fen�tre propri�taire de la bo�te de dialogue
    '  (ce n'est pas indispensable)
    Dim lhWnd&
    #If bVersionWord Then
        lhWnd = 0
    #ElseIf bVersionVBA Then
        lhWnd = Application.hWnd
    #Else
        lhWnd = Me.hWnd
    #End If
    
    Const sMsgFiltreDoc$ = _
        "Document Texte (*.txt) : bloc-notes Windows" & vbNullChar & "*.txt" & vbNullChar & _
        "Document Word (*.doc)" & vbNullChar & "*.doc" & vbNullChar & _
        "Document Html (*.htm ou *.html) : web" & vbNullChar & "*.htm*" & vbNullChar & _
        "Autre document (*.*)" & vbNullChar & "*.*"
    Const sMsgTitreBoiteDlg$ = _
        "Veuillez choisir un fichier texte ou un document convertible en .txt"
    Dim sInitDir$, sFichier$
    ' Initialiser le chemin seulement la premi�re fois
    Static bDejaInit As Boolean
    If Not bDejaInit Then
        bDejaInit = True
        sInitDir = m_oVBTxtFnd.m_sCheminDossierCourant
    End If
    If bChoisirUnFichierAPI(sFichier, sMsgFiltreDoc, _
        sMsgTitreBoiteDlg, sInitDir, lhWnd) Then
        ' Convertir le fichier en .txt si son extension
        '  est celle d'un document convertible (.doc, .html ou .htm)
        m_oVBTxtFnd.m_sCheminFichierTxtDirect = sFichier
        If Not m_oVBTxtFnd.bConvertirDocEnTxt() Then Exit Sub
        Me.TxtCheminDocument = m_oVBTxtFnd.m_sCheminFichierTxtDirect 'sFichier
        m_oVBTxtFnd.Sablier bDesactiver:=True
    End If
    VerifierOperationsPossibles bVerifDocumentSeul:=True
    
End Sub

Private Sub CmdAjouterDocument_Click()
    
    ' Indexer un nouveau document
    
    ' Interdire la r�-entrance dans cette fonction
    Me.CmdAjouterDocument.Enabled = False
    ' Autoriser l'interruption de l'indexation
    Me.CmdInterrompre.Enabled = True
    Me.CmdInterrompre.SetFocus
    Me.CmdChercher.Enabled = False
    
    If m_oVBTxtFnd.bIndexerDocument( _
        Me.TxtCheminDocument, Me.TxtCodeDoc) Then
        Me.TxtCodeDoc = ""
        Me.CmdInterrompre.Enabled = False
        m_oVBTxtFnd.ListerDocumentsIndexes Me.TxtResultat
        VerifierOperationsPossibles
        Me.TxtMot.SetFocus
    End If
    
    Me.CmdInterrompre.Enabled = False
    
End Sub

Private Sub LstTypeAffichResult_Click()
    
    ' Gerer le type d'affichage des r�sultats (phrase ou paragraphe)
    
    If Not m_bInit Then Exit Sub
    If Not Me.CmdChercher.Enabled Then Exit Sub
    CmdChercher_Click
    
End Sub

Private Sub CmdChercher_Click()
    
    ' Chercher les occurrences d'un mot
    
    Me.CmdInterrompre.Enabled = True
    Me.CmdInterrompre.SetFocus
    Me.CmdChercher.Enabled = False ' Eviter la r�-entrance dans la fonction
    m_oVBTxtFnd.ChercherOccurrencesMot Me.TxtMot, Me.TxtResultat, Me.LstTypeAffichResult
    Me.CmdInterrompre.Enabled = False
    Me.TxtMot.SetFocus
    Me.CmdChercher.Enabled = True

End Sub

Private Sub TxtMot_Change()
    
    VerifierOperationsPossibles
    
End Sub

Private Sub TxtMot_Click()
    
    VerifierOperationsPossibles
    
End Sub

#If bVersionVBA Then

Private Sub TxtMot_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
#Else

Private Sub TxtMot_KeyDown(KeyCode As Integer, Shift As Integer)

#End If
    
    ' Traiter la touche Entr�e sur la zone de saisie n�1
    If KeyCode = vbKeyReturn Then CmdChercher_Click: Exit Sub

End Sub

Private Sub TxtCheminDocument_Change()

    VerifierOperationsPossibles bVerifDocumentSeul:=True
    
End Sub

Private Sub TxtCodeDoc_Change()
    
    ' Code du document = cl� unique du document
    
    VerifierOperationsPossibles bVerifDocumentSeul:=True
    
End Sub

#If bVersionVBA Then

Private Sub TxtCodeDoc_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

#Else

Private Sub TxtCodeDoc_DblClick()

#End If

    m_oVBTxtFnd.ListerDocumentsIndexes Me.TxtResultat

End Sub

Private Sub CmdInterrompre_Click()
    
    m_oVBTxtFnd.Interrompre
    
End Sub

Private Sub VerifierOperationsPossibles( _
    Optional bVerifDocumentSeul As Boolean = False)
    
    ' V�rifier les op�rations possibles selon l'�tat de l'interface

    ' Si une indexation est en cours, ne pas r�activer les boutons de commande
    If Me.CmdInterrompre.Enabled Then Exit Sub

    Dim sMsg$, sMsgMot$, sMsgDoc$
    If m_oVBTxtFnd.iNbDocumentsIndexes > 0 Then
        Me.TxtMot.Enabled = True
    Else
        sMsgMot = "Aucun document n'est index�"
    End If

    ' Activer le bouton Chercher si le mot existe
    Me.CmdChercher.Enabled = False
    Dim oMot As clsMot
    If Not m_oVBTxtFnd.bMotExiste(Me.TxtMot, oMot) Then
        If Me.TxtMot <> "" Then sMsgMot = "Mot non trouv� : " & Me.TxtMot
    Else
        sMsgMot = "Mot trouv� : " & Me.TxtMot & " (" & _
            oMot.lNbOccurences & " occurrences)"
        Me.CmdChercher.Enabled = True
    End If

    ' V�rifier si le fichier document existe
    Me.CmdAjouterDocument.Enabled = False
    If Not bFichierExiste(Me.TxtCheminDocument) Then
        If Me.TxtCheminDocument <> "" Then _
            sMsgDoc = "Fichier inexistant : " & Me.TxtCheminDocument
        GoTo Fin
    End If
    
    ' Activer le bouton Ajouter (un document � indexer)
    '  si le code document n'existe pas d�j�
    If m_oVBTxtFnd.bCodeDocExiste(Me.TxtCodeDoc) Then
        sMsgDoc = "Code document d�j� utilis� : " & Me.TxtCodeDoc: GoTo Fin
    Else
        If bVerifDocumentSeul Then
            Dim sCodeDoc$
            sCodeDoc = Me.TxtCodeDoc
            If sCodeDoc = "" Then sCodeDoc = m_oVBTxtFnd.sCodeDocDefaut
            sMsgDoc = "Code mn�monique associ� au document : " & sCodeDoc
        End If
    End If
    Me.CmdAjouterDocument.Enabled = True
    
Fin:
    sMsg = sMsgMot
    If sMsgDoc <> "" Then sMsg = sMsgDoc
    If bVerifDocumentSeul Then sMsg = sMsgDoc
    If sMsg <> "" Or Not bVerifDocumentSeul Then Me.LblAvancement.Caption = sMsg
    
End Sub

#If bVersionVBA Then

Private Sub TxtResultat_MouseUp(ByVal Button%, ByVal Shift%, ByVal X!, ByVal Y!)

    ' Lancement automatique d'une recherche avec la version VBA : la gestion
    '  des �v�nements ne fonctionne pas aussi bien qu'en VB6 :
    '  Me.TxtResultat.SelText n'est pas encore renseign� dans l'�v�nement
    '  TxtResultat_DblClick, on est donc oblig� de traiter TxtResultat_MouseUp,
    '  et dans ce cas, on ne lance pas de recherche automatiquement pour �viter
    '  un conflit avec la position de la barre de d�filement
    
    ' Quitter si une op�ration est en cours
    If Me.CmdInterrompre.Enabled Then Exit Sub
    Dim sMotSelFin$
    If m_oVBTxtFnd.bHyperTexte(Me.TxtResultat.SelText, sMotSelFin) Then
        Me.TxtMot = sMotSelFin
        'CmdChercher_Click
    End If
    
End Sub
    
#Else

Private Sub TxtResultat_DblClick()
    
    ' Quitter si une op�ration est en cours
    If Me.CmdInterrompre.Enabled Then Exit Sub
    Dim sMotSelFin$
    If m_oVBTxtFnd.bHyperTexte(Me.TxtResultat.SelText, sMotSelFin) Then
        Me.TxtMot = sMotSelFin
        CmdChercher_Click
    End If

End Sub

#End If
