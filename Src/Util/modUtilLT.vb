
' Fichier modUtilLT.vb : Module de fonctions utilitaires en liaison tardive
' ---------------------

Option Strict Off ' Module non strict

Module modUtilitairesLiaisonTardive

    Public Function bConvertirDocEnTxt2(sCheminFichierSelect$,
        sCheminFichierTxt$, sCheminDossierCourant$,
        msgDelegue As clsMsgDelegue, bOptionTexteUnicode As Boolean,
        bVerifierUnicode As Boolean,
        ByRef bTxtUnicode As Boolean,
        ByRef bAvertUnicode As Boolean,
        ByRef bInfoTxtNonUnicode As Boolean) As Boolean

        ' Convertir un fichier .doc ou .html en .txt

        Dim oWrdH As clsHebWord = Nothing
        bTxtUnicode = False
        bAvertUnicode = False
        bInfoTxtNonUnicode = False

        Try

            oWrdH = New clsHebWord(bInterdireAppliAvant:=False)
            If IsNothing(oWrdH.oWrd) Then Return False

            Const wdCRLF% = 0
            Const wdFormatText% = 2

            msgDelegue.AfficherMsg("Ouverture de Microsoft Word...")
            Application.DoEvents() : Cursor.Current = Cursors.WaitCursor

            oWrdH.oWrd.Visible = False

            msgDelegue.AfficherMsg("Ouverture du fichier " &
                sCheminFichierSelect & "...")
            Application.DoEvents() : Cursor.Current = Cursors.WaitCursor

            oWrdH.oWrd.Documents.Open(sCheminFichierSelect)

            msgDelegue.AfficherMsg("Conversion en .txt du fichier " &
                sCheminFichierSelect & "...")
            Application.DoEvents() : Cursor.Current = Cursors.WaitCursor

            oWrdH.oWrd.ChangeFileOpenDirectory(sCheminDossierCourant)

            ' 28/04/2018 Correction du bug "Espace mémoire insuffisant" 
            ' (voir la ligne CharacterUnitFirstLineIndent = 0 plus bas)
            ' 19/01/2019 Maintenant ce code provoque une autre erreur : désactivation !
            Const bSupprSignets As Boolean = False
            If bSupprSignets Then
                msgDelegue.AfficherMsg("Suppression des signets du fichier " &
                    sCheminFichierSelect & "...")
                oWrdH.oWrd.ActiveDocument.Bookmarks.ShowHidden = True
                Dim objBkm As Object = Nothing ' As Bookmark
                For Each objBkm In oWrdH.oWrd.ActiveDocument.Bookmarks
                    'Try
                    objBkm.Delete()
                    'Catch 'ex As Exception
                    '' L'exception System.Runtime.InteropServices.COMException s'est produite
                    '' ErrorCode=-2146822463
                    '' HResult=-2146822463
                    '' Message=L'objet a été supprimé.
                    'Exit For
                    'End Try
                Next
                oWrdH.oWrd.ActiveDocument.Bookmarks.ShowHidden = False
            End If

            ' 02/05/2010 Ne pas ajouter d'espace de présentation : 
            '  AddBiDiMarks:=False et supprimer tous les retraits
            oWrdH.oWrd.Selection.WholeStory()
            With oWrdH.oWrd.Selection.ParagraphFormat

                ' 18/05/2014 Si le document est en mode Plan alors repasser en mode affichage Page
                Const wdMasterView% = 5 ' Membre de Word.WdViewType
                Const wdPrintView% = 3
                Const wdPaneNone% = 0 ' Membre de Word.WdSpecialPane
                If oWrdH.oWrd.ActiveWindow.View.SplitSpecial = wdPaneNone Then
                    If oWrdH.oWrd.ActiveWindow.ActivePane.View.Type = wdMasterView Then
                        oWrdH.oWrd.ActiveWindow.ActivePane.View.Type = wdPrintView
                    End If
                Else
                    If oWrdH.oWrd.ActiveWindow.View.Type = wdMasterView Then
                        oWrdH.oWrd.ActiveWindow.View.Type = wdPrintView
                    End If
                End If

                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .FirstLineIndent = 0 'CentimetersToPoints(0)

                ' Solution trouvée à ce bug : supprimer tous les signets cachés 
                ' (ceux de la table des matières)
                ' "Microsoft Word" "Espace mémoire insuffisant" "Une fois terminée, cette action ne pourra pas être annulée" Continuer ?
                ' "Word has insufficient memory" "You will not be able to undo this action once it is completed" "Do you want to continue?"
                .CharacterUnitFirstLineIndent = 0

            End With

            ' 19/09/2009 Il faut préciser AllowSubstitutions:=False
            '  sinon des substitutions peuvent avoir lieu par exemple de « en "
            'http://msdn.microsoft.com/fr-fr/library/microsoft.office.tools.word.document.saveas%28VS.80%29.aspx

            ' 15/05/2010
            Dim iEncodage% = iCodePageWindowsLatin1252
            If bOptionTexteUnicode Then iEncodage = iEncodageUnicodeUTF8

            msgDelegue.AfficherMsg("Ecriture du fichier " & sCheminFichierTxt & "...")
            Application.DoEvents() : Cursor.Current = Cursors.WaitCursor
            ' 02/05/2010 AddBiDiMarks:=False : Ne pas ajouter d'espace de présentation
            oWrdH.oWrd.ActiveDocument.SaveAs(
                FileName:=sCheminFichierTxt,
                FileFormat:=wdFormatText,
                Encoding:=iEncodage,
                LineEnding:=wdCRLF,
                AllowSubstitutions:=False,
                AddBiDiMarks:=False)

            ' 23/11/2018 Si l'option Unicode n'est pas activé, tester quand même et comparer
            Dim sCheminU$ = ""
            If bVerifierUnicode Then
                msgDelegue.AfficherMsg("Vérication Unicode...")
                Application.DoEvents() : Cursor.Current = Cursors.WaitCursor
                sCheminU = sDossierParent(sCheminFichierTxt) & "\" &
                    IO.Path.GetFileNameWithoutExtension(sCheminFichierTxt) & "_Utmp00.txt"
                If Not bOptionTexteUnicode Then
                    oWrdH.oWrd.ActiveDocument.SaveAs(
                        FileName:=sCheminU,
                        FileFormat:=wdFormatText,
                        Encoding:=iEncodageUnicodeUTF8,
                        LineEnding:=wdCRLF,
                        AllowSubstitutions:=False,
                        AddBiDiMarks:=False)
                    Dim sTexteU$ = sLireFichier(sCheminU, bLectureSeule:=True, bUnicodeUTF8:=True)
                    Dim sTexte$ = sLireFichier(sCheminFichierTxt, bLectureSeule:=True)
                    If sTexteU <> sTexte Then
                        bAvertUnicode = True
                        bTxtUnicode = True
                        'MsgBox("Le texte contient des caractères Unicode et l'option n'est pas activée", _
                        '    MsgBoxStyle.Exclamation, m_sTitreMsg)
                    End If
                Else
                    oWrdH.oWrd.ActiveDocument.SaveAs(
                        FileName:=sCheminU,
                        FileFormat:=wdFormatText,
                        Encoding:=iCodePageWindowsLatin1252,
                        LineEnding:=wdCRLF,
                        AllowSubstitutions:=False,
                        AddBiDiMarks:=False)
                    Dim sTexte$ = sLireFichier(sCheminU, bLectureSeule:=True)
                    Dim sTexteU$ = sLireFichier(sCheminFichierTxt, bLectureSeule:=True, bUnicodeUTF8:=True)
                    If sTexteU = sTexte Then
                        bInfoTxtNonUnicode = True
                        'MsgBox("Le texte ne contient pas de caractères Unicode", _
                        '    MsgBoxStyle.Exclamation, m_sTitreMsg)
                    Else
                        bTxtUnicode = True
                    End If
                End If
            End If

            msgDelegue.AfficherMsg("Fermeture du fichier " & sCheminFichierTxt & "...")
            Application.DoEvents() : Cursor.Current = Cursors.WaitCursor
            oWrdH.oWrd.ActiveDocument.Close()

            If bVerifierUnicode Then bSupprimerFichier(sCheminU, bPromptErr:=True)

            sCheminFichierSelect = sCheminFichierTxt
            msgDelegue.AfficherMsg(sMsgOperationTerminee)
            Return True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirDocEnTxt2")
            Return False

        Finally
            If Not IsNothing(oWrdH.oWrd) Then
                'oWrdH.oWrd.Quit()
                ' Ne pas sauvegarder les changements s'il y a eu une erreur (document déjà ouvert)
                oWrdH.oWrd.Quit(SaveChanges:=False)
                oWrdH.oWrd = Nothing
                oWrdH.Quitter()
                oWrdH = Nothing
            End If
        End Try

    End Function

End Module