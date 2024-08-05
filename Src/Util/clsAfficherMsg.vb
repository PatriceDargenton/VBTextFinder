
' Fichier clsAfficherMsg.vb : Classes de gestion des messages via des délégués
' -------------------------

Imports System.IO

Public Class clsTickEventArgs : Inherits EventArgs
    ' Classe pour l'événement Tick : avancement d'une unité de temps : TIC-TAC
    '  utile pour mettre à jour l'heure en cours, ou pour scruter une annulation
    Public Sub New()
    End Sub
End Class

Public Class clsMsgEventArgs : Inherits EventArgs
    ' Classe pour l'événement Message
    Private m_sMsg$ = "" 'Nothing
    Public Sub New(sMsg$)
        'If sMsg Is Nothing Then Throw New NullReferenceException
        If sMsg Is Nothing Then sMsg = ""
        Me.m_sMsg = sMsg
    End Sub
    Public ReadOnly Property sMessage$()
        Get
            Return Me.m_sMsg
        End Get
    End Property
End Class

Public Class clsFECEventArgs : Inherits EventArgs
    ' Classe pour l'événement Fichier En Cours (FEC)
    Private m_iNumFichierEnCours% = 0
    Public Sub New(iNumFichierEnCours%)
        Me.m_iNumFichierEnCours = iNumFichierEnCours
    End Sub
    Public ReadOnly Property iNumFichierEnCours%()
        Get
            Return Me.m_iNumFichierEnCours
        End Get
    End Property
End Class

Public Class clsFSIEventArgs : Inherits EventArgs
    ' Classe pour l'événement FileSystemInfo
    Private m_fsi As FileSystemInfo
    Public ReadOnly Property fsi() As FileSystemInfo
        Get
            Return Me.m_fsi
        End Get
    End Property
    Public Sub New(fsi As FileSystemInfo)
        Me.m_fsi = fsi
    End Sub
End Class

Public Class clsAvancementEventArgs : Inherits EventArgs

    ' Classe pour l'événement Avancement

    Private m_sMsg$ = ""
    Private m_lAvancement& = 0

    Public Sub New(sMsg$)
        If sMsg Is Nothing Then sMsg = ""
        Me.m_sMsg = sMsg
    End Sub
    Public Sub New(lAvancement&)
        Me.m_lAvancement = lAvancement
    End Sub
    Public Sub New(lAvancement&, sMsg$)
        Me.m_lAvancement = lAvancement
        If sMsg Is Nothing Then sMsg = ""
        Me.m_sMsg = sMsg
    End Sub
    Public ReadOnly Property sMessage$()
        Get
            Return Me.m_sMsg
        End Get
    End Property
    Public ReadOnly Property lAvancement&()
        Get
            Return Me.m_lAvancement
        End Get
    End Property
End Class

Public Class clsSablierEventArgs : Inherits EventArgs
    ' Classe pour l'événement Sablier
    Private m_bDesactiver As Boolean = False
    Public Sub New(bDesactiver As Boolean)
        Me.m_bDesactiver = bDesactiver
    End Sub
    Public ReadOnly Property bDesactiver() As Boolean
        Get
            Return Me.m_bDesactiver
        End Get
    End Property
End Class

Public Class clsMsgDelegue

    ' Classe de gestion des messages via des délégués

    'Const bDoEvents As Boolean = False ' 16/10/2016 Pas de différence constatée !
    Const bDoEvents As Boolean = True ' 04/02/2018 Il faut activer pour gérer l'annulation

    'Private Delegate Sub GestEvTick(sender As Object, e As clsTickEventArgs)
    'Public Event EvTick As GestEvTick
    Public Event EvTick As EventHandler(Of clsTickEventArgs) ' CA1003

    'Private Delegate Sub GestEvAfficherMessage(sender As Object, e As clsMsgEventArgs)
    'Public Event EvAfficherMessage As GestEvAfficherMessage
    Public Event EvAfficherMessage As EventHandler(Of clsMsgEventArgs) ' CA1003

    'Private Delegate Sub GestEvAfficherFEC(sender As Object, e As clsFECEventArgs)
    'Public Event EvAfficherNumFichierEnCours As GestEvAfficherFEC
    Public Event EvAfficherNumFichierEnCours As EventHandler(Of clsFECEventArgs) ' CA1003

    'Private Delegate Sub GestEvAfficherFSI(sender As Object, e As clsFSIEventArgs)
    'Public Event EvAfficherFSIEnCours As GestEvAfficherFSI
    Public Event EvAfficherFSIEnCours As EventHandler(Of clsFSIEventArgs) ' CA1003

    'Private Delegate Sub GestEvAfficherAvancement(sender As Object, e As clsAvancementEventArgs)
    'Public Event EvAfficherAvancement As GestEvAfficherAvancement
    Public Event EvAfficherAvancement As EventHandler(Of clsAvancementEventArgs) ' CA1003

    'Private Delegate Sub GestEvSablier(sender As Object, e As clsSablierEventArgs)
    'Public Event EvSablier As GestEvSablier
    Public Event EvSablier As EventHandler(Of clsSablierEventArgs) ' CA1003

    Public m_bAnnuler As Boolean
    Public m_bErr As Boolean ' 21/03/2016

    Public Sub New()
    End Sub

    Public Sub AfficherMsg(sMsg$)
        Dim e As New clsMsgEventArgs(sMsg)
        RaiseEvent EvAfficherMessage(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

    Public Sub AfficherFichierEnCours(iNumFichierEnCours%)
        Dim e As New clsFECEventArgs(iNumFichierEnCours)
        RaiseEvent EvAfficherNumFichierEnCours(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

    Public Sub AfficherFSIEnCours(fsi As FileSystemInfo)
        Dim e As New clsFSIEventArgs(fsi)
        RaiseEvent EvAfficherFSIEnCours(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

    Public Sub AfficherAvancement(lAvancement&, sMsg$)
        Dim e As New clsAvancementEventArgs(lAvancement, sMsg)
        RaiseEvent EvAfficherAvancement(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

    Public Sub Tick()
        Dim e As New clsTickEventArgs()
        RaiseEvent EvTick(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

    Public Sub Sablier(Optional bDesactiver As Boolean = False)
        Dim e As New clsSablierEventArgs(bDesactiver)
        RaiseEvent EvSablier(Me, e)
        If bDoEvents Then TraiterMsgSysteme_DoEvents()
    End Sub

End Class