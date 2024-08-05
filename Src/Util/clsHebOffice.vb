
Option Strict Off ' Pour oWkb.Close()

' clsHebOffice : classe pour héberger une application Office (Word, Excel, ...)
'  basée sur clsExcelHost, cf. XLDOTNET :

' XLDOTNET : QUITTER EXCEL SANS LAISSER D'INSTANCE EN RAM
' https://codes-sources.commentcamarche.net/source/27541

#Region "Informations"

' D'après :

'   ======================================================================================
'   clsExcelHost : Classe pour héberger Excel
'   ============

' Title: EXCEL.EXE Process Killer
' Description: After many weeks of trying to figure out why the EXCEL.EXE Process 
'  does not want to go away from the Task Manager, I wrote this class that will ensure 
'  that the correct EXCEL.EXE Process is closed. This is after using Excel.Application 
'  via Automation from a VB.NET/ASP.NET application.
' This file came from Planet-Source-Code.com... the home millions of lines of source code
' You can view comments on this code/and or vote on it at: 
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=1998&lngWId=10

' The author may have retained certain copyrights to this code...
'  please observe their request and the law by reviewing all copyright conditions 
'  at the above URL.

'   Author: I.W Coetzer 2004/01/22
'   *Thanks Dan for the process idea.
'   Classe commentée et légèrement modifiée par Patrice Dargenton le 05/11/2004
'   *Solution to the EXCEL.EXE Process that does not want to go away from task manager.
'
'   ======================================================================================

#End Region

#Region "clsHebOffice"

Public Class clsHebOffice

    Public m_oApp As Object = Nothing 'Protected

    Private m_iIdProcess% = 0
    Public m_bAppliDejaOuverte As Boolean = False
    Public m_bInterdireAppliAvant As Boolean = True
    Public m_sNomProcess$ = ""

    Public Sub New(sNomProcess$, sClasseObjet$,
        Optional bInterdireAppliAvant As Boolean = True,
        Optional bReutiliserInstance As Boolean = False)

        ' Exemple :
        'Private Const sClasseObjetWord$ = "Word.Application"
        'Private Const sNomProcessWord$ = "Word"

        'Private Const sClasseObjetExcel$ = "Excel.Application"
        'Private Const sNomProcessExcel$ = "Excel"

        Me.m_iIdProcess = 0
        Me.m_bAppliDejaOuverte = False
        Me.m_bInterdireAppliAvant = bInterdireAppliAvant
        Me.m_sNomProcess = sNomProcess
        Dim sNomProcessMaj$ = sNomProcess.ToUpper

        ' Liste des processus avant le mien
        Dim aProcessAv() As Process = Process.GetProcesses()

        Dim j%
        For j = 0 To aProcessAv.GetUpperBound(0)
            If aProcessAv(j).ProcessName = sNomProcessMaj Then
                Me.m_bAppliDejaOuverte = True
                Exit For
            End If
        Next j
        If bInterdireAppliAvant And Me.m_bAppliDejaOuverte Then Exit Sub

        ' Créer le processus demandé
        Try

            If Me.m_bAppliDejaOuverte And bReutiliserInstance Then
                ' Pb : on récupère n'importe quelle instance
                '  il faudrait plutôt conserver l'instance qu'on a créée
                Me.m_oApp = GetObject(, sClasseObjet)
            Else
                Me.m_oApp = CreateObject(sClasseObjet)
            End If

        Catch Ex As Exception
            'AfficherMsgErreur2(Ex, "clsHebOffice:New", _
            '    sNomProcess & " n'est pas installé !")
            MsgBox(sClasseObjet & " n'est pas installé !" & vbLf &
                Ex.Message, MsgBoxStyle.Critical,
                "Lancement de " & sNomProcess)
            Me.m_oApp = Nothing
            Exit Sub
        End Try

        ' Liste des processus après le mien : la différence me donnera l'Id du mien
        Dim aProcessAp() As Process = Process.GetProcesses()

        Dim i%
        Dim bMonProcess As Boolean
        For j = 0 To aProcessAp.GetUpperBound(0)
            If aProcessAp(j).ProcessName = sNomProcessMaj Then
                bMonProcess = True
                ' Parcours des processus avant le mien
                For i = 0 To aProcessAv.GetUpperBound(0)
                    If aProcessAv(i).ProcessName = sNomProcessMaj Then
                        If aProcessAp(j).Id = aProcessAv(i).Id Then
                            ' S'il existait avant, ce n'était pas le mien
                            bMonProcess = False
                            Exit For
                        End If
                    End If
                Next i
                If bMonProcess = True Then
                    ' Maintenant que j'ai son Id, je pourrai le tuer
                    '  cette méthode marche toujours !
                    Me.m_iIdProcess = aProcessAp(j).Id
                    Exit For
                End If
            End If
        Next j

    End Sub

    Public Sub Quitter()

        If Me.m_iIdProcess = 0 Then Exit Sub

        If Not bMonInstanceOuverte() Then
            ' 28/08/2009 L'instance n'est plus ouverte, mais voir s'il faut libérer les variables
            'Try ' 27/02/2011 Déjà Try catch dans la fct LibererObjetCom
            LibererObjetCom(Me.m_oApp)
            'Me.m_oApp = Nothing : Déjà fait
            'Catch ex As Exception
            '    Debug.WriteLine(ex)
            'End Try
            Exit Sub
        End If

        LibererObjetCom(Me.m_oApp) ' 27/02/2011

        ' 27/02/2011 Cette ligne peut echouer si le process est déjà quitté :
        '  "Un processus ayant l'ID x n'est pas exécuté"
        'Dim monProc As Process = Process.GetProcessById(Me.m_iIdProcess)
        Dim monProc As Process = Nothing
        Try
            monProc = Process.GetProcessById(Me.m_iIdProcess)
        Catch 'ex As Exception
            ' Le processus vient de se terminer, il n'y a plus rien à faire
            Exit Sub
        End Try

        ' Même si l'instance a été fermée, monProc est toujours valide :
        '  ce test n'est pas suffisant
        If Not IsNothing(monProc) Then
            Try

                ' 15/05/2009 Libérer avant de tuer le processus
                ' Pour Excel l'objet oXL a déjà été libéré,
                '  mais il faut aussi libérer m_oApp ? c'est pourtant le même pointeur !?
                'LibererObjetCom(Me.m_oApp) 27/02/2011
                'Me.m_oApp = Nothing : Déjà fait

                ' Si l'instance ne nous appartient pas, on ne peut pas la fermer
                '  mais on ne reçoit aucune exception !
                ' 27/02/2011 If Not monProc.HasExited : inutile de tuer alors
                If Not monProc.HasExited Then monProc.Kill()
                ' On ne peut pas interroger immédiatement ExitCode, seule solution :
                '  vérifier si l'appli est toujours ouverte avec l'iIdProcess
                'If monProc.ExitCode = -1 Then
                '    ' MainModule vaut alors {"Accès refusé"}
                '    Debug.WriteLine("Impossible de fermer " & _
                '        Me.m_sNomProcess & " : " & monProc.MainModule.ToString)
                'End If
            Catch ex As Exception
                Debug.WriteLine(ex)
            End Try
        End If

    End Sub

    Public Function bMonInstanceOuverte() As Boolean

        ' Vérifier si l'instance que j'ai utilisée est encore ouverte
        '  (elle a pu être fermée par l'utilisateur si on l'autorise)
        If Me.m_iIdProcess = 0 Then Return False

        ' 28/08/2009 Avec Word cela ne marche pas, car Word déjà quitté
        ' D'abord on vérifie s'il ne reste plus aucune instance
        If Not bOuvert(Me.m_sNomProcess) Then Return False
        Dim monProc As Process
        Try ' Puis on teste si on peut récupérer l'instance
            monProc = Process.GetProcessById(Me.m_iIdProcess)
        Catch
            ' On ne peut pas : l'instance est déjà fermée
            ' "Un processus ayant l'ID xxxx n'est pas exécuté."
            Return False
        End Try

        ' Même si l'instance a été fermée, monProc est toujours valide :
        '  cette fonction n'est pas suffisante
        'If IsNothing(monProc) Then Exit Function
        'bMonInstanceOuverte = True

        ' 15/05/2009
        Try
            Return Not monProc.HasExited()
        Catch 'ex As Exception
            ' On vient juste de fermer
            Return False
        End Try

    End Function

    Public Shared Function bOuvert(sNomProcess$) As Boolean

        ' Vérifier si l'application est déjà ouverte 
        ' (pour le cas où cela poserait problème, faire la vérification au départ)

        Dim sNomProcessMaj$ = sNomProcess.ToUpper

        ' Liste des processus avant le mien
        Dim aProcessAv() As Process = Process.GetProcesses()

        Dim j%
        For j = 0 To aProcessAv.GetUpperBound(0)
            If aProcessAv(j).ProcessName = sNomProcessMaj Then Return True
        Next j
        Return False

    End Function

    Public Shared Sub LibererObjetCom(ByRef oCom As Object)
        ' ByRef car on fait oCom = Nothing)

        ' D'abord Quitter ou Fermer, puis ReleaseComObject puis oCom = Nothing

        ' Pour Excel :
        ' Quit Excel and clean up.
        ' oBook.Close(false, oMissing, oMissing);
        ' System.Runtime.InteropServices.Marshal.ReleaseComObject (oBook);
        ' oBook = null;
        ' System.Runtime.InteropServices.Marshal.ReleaseComObject (oBooks);
        ' oBooks = null;
        ' oExcel.Quit();
        ' System.Runtime.InteropServices.Marshal.ReleaseComObject (oExcel);
        ' oExcel = null;

        If IsNothing(oCom) Then Exit Sub
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(oCom)
        Catch ex As Exception
            Debug.WriteLine(ex)
        Finally
            oCom = Nothing
        End Try

    End Sub

End Class

#End Region

#Region "clsHebExcel"

Public Class clsHebExcel : Inherits clsHebOffice

    ' clsHebExcel : classe pour héberger Excel, basée sur clsHebOffice

    Private Const sClasseObjetExcel$ = "Excel.Application"
    Private Const sNomProcessExcel$ = "Excel"

    Public oXL As Object = Nothing

    Public Sub New(Optional bInterdireAppliAvant As Boolean = True,
        Optional bReutiliserInstance As Boolean = False)
        MyBase.New(sNomProcessExcel, sClasseObjetExcel,
            bInterdireAppliAvant, bReutiliserInstance)
        Me.oXL = Me.m_oApp
    End Sub

    Public Overloads Shared Function bOuvert() As Boolean
        bOuvert = clsHebOffice.bOuvert(sNomProcessExcel)
    End Function

    Public Sub Fermer(ByRef oSht As Object, ByRef oWkb As Object, bQuitter As Boolean,
        Optional bFermerClasseur As Boolean = True,
        Optional bLibererXLSiResteOuvert As Boolean = True)

        ' Liberer correctement le classeur, et le femer si demandé, 
        '  et quitter Excel si demandé

        If bFermerClasseur AndAlso Not IsNothing(oWkb) Then
            'msgDelegue.AfficherMsg("Fermeture du classeur...")
            Try
                oWkb.Close(SaveChanges:=False) ' Si Excel 2007 veut sauver qqch.: Non merci.
            Catch ex As Exception
                Debug.WriteLine(ex)
            End Try
        End If
        LibererObjetCom(oSht)
        LibererObjetCom(oWkb)

        ' Conserver Excel ouvert (par exemple pour visualiser l'actualisation d'un classeur)
        '  on libère oXL dans le cas général (sauf si on doit continuer d'utiliser l'instance
        '  par ex. pour effectuer d'autres traitements)
        If Not bQuitter Then
            If bLibererXLSiResteOuvert Then LibererObjetCom(Me.oXL)
            Exit Sub
        End If

        If Not IsNothing(Me.oXL) Then
            Try
                'msgDelegue.AfficherMsg("Fermeture d'Excel...")
                If Me.bMonInstanceOuverte() Then Me.oXL.Quit()
            Catch ex As Exception
                ' L'application a été fermée par l'utilisateur, on n'y a plus accès
                '  ou bien on tente d'utiliser l'objet Me.oXL qui a déjà été libéré
                '  "Un objet COM qui a été séparé de son RCW sous-jacent ne peut pas être utilisé."
                Debug.WriteLine(ex)
            End Try
            'msgDelegue.AfficherMsg("Libération d'Excel...")
            LibererObjetCom(Me.oXL)
        End If

        Me.Quitter()

    End Sub

End Class

#End Region

#Region "clsHebWord"

Public Class clsHebWord : Inherits clsHebOffice

    ' clsHebWord : classe pour héberger Word, basée sur clsHebOffice

    Private Const sClasseObjetWord$ = "Word.Application"
    Private Const sNomProcessWrd$ = "Winword" '"Word"

    Public oWrd As Object = Nothing

    Public Sub New(Optional bInterdireAppliAvant As Boolean = True)
        MyBase.New(sNomProcessWrd, sClasseObjetWord, bInterdireAppliAvant)
        oWrd = Me.m_oApp
    End Sub

    Public Overloads Shared Function bOuvert() As Boolean
        bOuvert = clsHebOffice.bOuvert(sNomProcessWrd)
    End Function

End Class

#End Region

#Region "clsHebNav"

Public Class clsHebNav

    ' clsHebNav : classe pour héberger un navigateur (Internet Explorer ou Firefox)

    Private Const sNomProcessIE$ = "iexplore"
    Private Const sNomProcessFireFox$ = "firefox"

    Public oAppNav As Object = Nothing
    Private m_iIdProcess%

    Public Sub New(sURL$)

        Me.m_iIdProcess = 0
        ' Liste des processus avant le mien
        Dim aProcessAv() As Process = Process.GetProcesses()

        OuvrirAppliAssociee(sURL, bVerifierFichier:=False)

        ' Liste des processus après le mien : la différence me donnera l'Id du mien
        Dim aProcessAp() As Process = Process.GetProcesses()

        Dim i%, j%
        Dim bMonProcessNav As Boolean
        For j = 0 To aProcessAp.GetUpperBound(0)
            Dim sNomProcess$ = aProcessAp(j).ProcessName
            If sNomProcess = sNomProcessIE Or sNomProcess = sNomProcessFireFox Then
                bMonProcessNav = True
                ' Parcours des processus avant le mien
                For i = 0 To aProcessAv.GetUpperBound(0)
                    Dim sNomProcess1$ = aProcessAv(i).ProcessName
                    If sNomProcess1 = sNomProcessIE Or
                       sNomProcess1 = sNomProcessFireFox Then
                        If aProcessAp(j).Id = aProcessAv(i).Id Then
                            ' S'il existait avant, ce n'était pas le mien
                            bMonProcessNav = False
                            Exit For
                        End If
                    End If
                Next i
                If bMonProcessNav = True Then
                    ' Maintenant que j'ai son Id, je pourrai le controler
                    Me.m_iIdProcess = aProcessAp(j).Id
                    Exit For
                End If
            End If
        Next j

    End Sub

    Public Function bOuvert() As Boolean

        ' On peut savoir si l'utilisateur a fermé le navigateur ouvert 
        '  par l'application 
        If Me.m_iIdProcess = 0 Then Return False
        Try
            Return Not Process.GetProcessById(Me.m_iIdProcess).HasExited()
        Catch 'ex As Exception
            ' On vient juste de fermer
            Return False
        End Try

    End Function

    Public Sub Quitter()

        If Me.m_iIdProcess = 0 Then Exit Sub

        'Process.GetProcessById(Me.m_iIdProcess).Kill()

        Dim monProc As Process = Process.GetProcessById(Me.m_iIdProcess)
        ' Même si l'instance a été fermée, monProc est toujours valide :
        '  ce test n'est pas suffisant
        If Not IsNothing(monProc) Then
            Try

                ' 15/05/2009 Libérer avant de tuer le processus
                LibererObjetCom(Me.oAppNav)
                'Me.oAppNav = Nothing : Déjà fait

                ' Si l'instance ne nous appartient pas, on ne peut pas la fermer
                '  mais on ne reçoit aucune exception !
                monProc.Kill()
                ' On ne peut pas interroger immédiatement ExitCode, seule solution :
                '  vérifier si l'appli est toujours ouverte avec l'iIdProcess
                'If monProc.ExitCode = -1 Then
                '    ' MainModule vaut alors {"Accès refusé"}
                '    Debug.WriteLine("Impossible de fermer " & _
                '        Me.m_sNomProcess & " : " & monProc.MainModule.ToString)
                'End If
            Catch ex As Exception
                Debug.WriteLine(ex)
            End Try
        End If

    End Sub

    Public Shared Sub LibererObjetCom(ByRef oCom As Object)
        ' ByRef car on fait oCom = Nothing)

        ' D'abord Quitter ou Fermer, puis ReleaseComObject puis oCom = Nothing

        If IsNothing(oCom) Then Exit Sub
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(oCom)
        Catch
        Finally
            oCom = Nothing
        End Try

    End Sub

End Class

#End Region