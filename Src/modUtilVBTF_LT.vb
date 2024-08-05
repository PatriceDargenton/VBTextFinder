
' Fichier modUtilLT.vb : Module de fonctions utilitaires pour VBTextFinder en liaison tardive
' ---------------------

Option Strict Off ' Module non strict

Module modUtilVBTF_LT

    Public Const sCodeLangueFr$ = "Fr"
    Public Const sCodeLangueEn$ = "En"
    Public Const sCodeLangueUk$ = "Uk"
    Public Const sCodeLangueUS$ = "Us"
    Public Const sCodeLangueEs$ = "Es"

    Public Function iConvCodeLangue%(sCodeLangue$)

        Const wdFrench% = 1036    ' (&H40C)
        Const wdEnglishUK% = 2057 ' (&H809) 
        Const wdEnglishUS% = 1033 ' (&H409)
        Const wdSpanish% = 1034   ' (&H40A)

        Dim iLang% = wdFrench
        Select Case sCodeLangue
            Case sCodeLangueEn : iLang = wdEnglishUS ' Compatibilité ascendante
            Case sCodeLangueUS : iLang = wdEnglishUS
            Case sCodeLangueUk : iLang = wdEnglishUK
            Case sCodeLangueEs : iLang = wdSpanish
            Case Else : iLang = wdFrench
        End Select
        Return iLang

    End Function

    Public Function bCreerDocIndex2(sCheminTxt$, sCheminDoc$,
        sTitre$, sExplication$, lNbMotsIndexes%,
        m_colDocs As Collection, m_bInterrompre As Boolean,
        ByRef bWord As Boolean, bTexteUnicode As Boolean, sCodeLangue$) As Boolean

        ' Creer le document index de VBTextFinder
        ' bWord est retourné dans le cas où Word n'est pas installé : ne plus retenter

        bCreerDocIndex2 = False
        Dim oWrdH As clsHebWord = Nothing
        Dim bErrBugWord As Boolean = False

        Try

            oWrdH = New clsHebWord(bInterdireAppliAvant:=False)
            If IsNothing(oWrdH.oWrd) Then
                bWord = False
                Return False
            End If
            bWord = True

            'http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.documents.open(office.11).aspx
            ' 1252 : Code page de Windows (Latin) pour récupérer les caractères unicodes,
            '  par exemple coeur avec oe collé
            'oWrdH.oWrd.Documents.Open(CType(sCheminTxt, Object), _
            '    Encoding:=iCodePageWindowsLatin1252)
            Dim iEncodage% = iCodePageWindowsLatin1252
            If bTexteUnicode Then iEncodage = iEncodageUnicodeUTF8
            Try
                oWrdH.oWrd.Documents.Open(sCheminTxt, Encoding:=iEncodage, ReadOnly:=True)
            Catch ex As Exception
                bErrBugWord = True
                'Debug.WriteLine(ex)
                'AfficherMsgErreur2(ex)

                ' La solution est ici : il faut nettoyer le dossier %temp%
                ' "C:\Documents and Settings\[Compte utilisateur]\Local Settings\Temp"
                ' http://us.generation-nt.com/there-insufficient-memory-save-document-now-help-118779201.html

                ' Il y a plusieurs raisons suggérées pour cette erreur
                ' (trop de fontes installées, pb. doc. maitre, Drag & drop, 
                '  "Launching User", ...), mais aucune ne correspond, Exemples :
                ' Doc. maitre Word 2003 : http://support.microsoft.com/kb/822511
                ' (mais dans notre cas il s'agit de Word 2002)
                ' Drag & drop : http://support.microsoft.com/kb/829933
                ' Pb avec ctrl Organization chart : solution proposée : trop de fontes installées
                '  http://support.microsoft.com/default.aspx/kb/238307/en-us?p=1
                ' Launching User : http://p2p.wrox.com/vbscript/1951-5097-there-insufficient-memory.html

                ' On ne récupère le code erreur spécifique de l'exception qu'en mode Debug dans l'IDE,
                '  en mode Release, on n'a que le texte, mais ne pas faire de test dessus :
                ' "Mémoire insuffisante. Enregistrez votre document immédiatement."
                '  car la langue peut changer :
                ' "There is insufficient memory. Save the document now."
                ' N° erreur VB6 : 5097
                ' En mode Debug dans l'IDE, on n'a plus d'info.:
                ' ErrorCode = -2146823191
                ' Exception System.Runtime.InteropServices.COMException
                ' La rubrique 24577 n'existe pas :
                ' HelpLink="C:\Program Files\Microsoft Office\Office10\1036\wdmain10.chm#24577"
                ' Source="Microsoft Word"
                ' StackTrace: Microsoft.VisualBasic.CompilerServices.LateBinding.InternalLateCall
                '             Microsoft.VisualBasic.CompilerServices.NewLateBinding.LateCall
                ' InnerException : vide

            End Try


            'Format 0:wdOpenFormatAuto, 5:wdOpenFormatUnicodeText, 4:wdOpenFormatText
            'Const wdOpenFormatAuto% = 0
            'Const wdOpenFormatAuto As Object = 0
            'oWrdH.oWrd.Documents.Open(FileName:=sCheminTxt, _
            '    ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, _
            '    PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
            '    WritePasswordDocument:="", WritePasswordTemplate:="", _
            '    Format:=wdOpenFormatAuto, Encoding:=1252)

            With oWrdH.oWrd.ActiveDocument.Content

                .InsertBefore(vbLf)

                ' Présentation de l'index
                .InsertBefore(sExplication)
                .InsertBefore("Nombre de mots distincts indexés : " & lNbMotsIndexes & vbLf)

                .InsertBefore(vbLf)

                ' Afficher la liste des documents indexés
                'Dim de As DictionaryEntry
                'For Each de In m_colDocs
                '    oDoc = DirectCast(de.Value, clsDoc)
                '    .InsertBefore(oDoc.sChemin & " (" & oDoc.sCle & ")" & vbLf)
                'Next de
                Dim oDoc As clsDoc
                For Each oDoc In m_colDocs
                    .InsertBefore(oDoc.sChemin & " (" & oDoc.sCodeDoc & ")" & vbLf)
                Next oDoc
                .InsertBefore("Liste des documents indexés :" & vbLf)

                .InsertBefore(vbLf)

                If m_bInterrompre Then .InsertBefore(
                    "(création du document index interrompue)" & vbLf)

                .InsertBefore(sTitre & vbLf)

                ' Mettre en largeur de page maximale pour améliorer la lisibilité du document
                Const wdPageFitBestFit% = 2
                oWrdH.oWrd.ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit

                ' Enlever la mise en forme -> Texte brut
                ' --------------------------------------
                .WholeStory()
                '.ClearFormatting() ' Marche en WordBasic mais pas en WordAutomation !
                ' .Style : Ne marche qu'en mode liaison précoce !!!
                '.Style.Item = oWrd.ActiveDocument.Styles.Item("Normal")
                Try
                    Const wdStyleNormal% = -1
                    .Style = wdStyleNormal

                    ' 03/05/2014 Fixer la langue dans le document, avec vérification activée
                    Dim iLang% = iConvCodeLangue%(sCodeLangue)
                    .LanguageID = iLang
                    .NoProofing = False ' Ne pas vérifier le texte : non
                    oWrdH.oWrd.CheckLanguage = False ' Détection auto de la langue : non

                Catch
                End Try
                ' Même en appliquant le style Normal, il faut aussi faire cela
                .Font.Name = "Times New Roman"
                .Font.Size = 12
                ' --------------------------------------

            End With

            Const wdFormatDocument% = 0 ' Attention : Int32, sinon ça ne marche pas !!!
            oWrdH.oWrd.ActiveDocument.SaveAs(sCheminDoc, wdFormatDocument)
            bCreerDocIndex2 = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bConvertirDocEnTxt2")

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

        If bErrBugWord Then
            Dim sLecteur$ = sLecteurDossier(Environment.GetFolderPath(
                Environment.SpecialFolder.System))
            Dim sUser$ = Environment.UserName
            MsgBox("Le dossier temporaire est trop volumineux !" & vbLf &
                "Il provoque une erreur 5097 (Mémoire insuffisante) dans Word." & vbLf &
                "Veuillez nettoyer le dossier caché :" & vbLf &
                sLecteur & "\Documents and Settings\" & sUser & "\Local Settings\Temp",
                MsgBoxStyle.Information)
        End If

    End Function

End Module