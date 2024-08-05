
' VBDico : faire un glossaire des mots hors dictionnaire
'  en parcourant un document Word ou compatible (.doc, .html, ...)
' http://patrice.dargenton.free.fr/CodesSources/VBDico.html

Option Strict Off

#Const bLiaisonPrecoce = False
#If bLiaisonPrecoce Then
    ' Cela requiert de télécharger : office primary interop assemblies
    '  et d'ajouter la référence à Microsoft.Office.Interop.Word.dll
    Imports Microsoft.Office.Interop
    Imports Microsoft.Office.Interop.Word
    Imports Microsoft.Office.Interop.Word.WdSaveOptions ' wdDoNotSaveChanges
    Imports Microsoft.Office.Interop.Word.WdSortFieldType
    Imports Microsoft.Office.Interop.Word.WdSortOrder
    Imports Microsoft.Office.Interop.Word.WdPageFit
    Imports Microsoft.Office.Interop.Word.WdLanguageID
    Imports Microsoft.Office.Interop.Word.WdUnits ' wdStory
#End If

Module modGlossaire

    ' Pour pouvoir localiser la ligne ayant provoqué une erreur, mettre bTrapErr = False
    Private Const bTrapErr As Boolean = True

    ' Nombre max. d'occurrence du mot affiché dans les listes du glossaire
    ' 0 pour en afficher aucune 18/05/2014
    Private Const iNbOccurencesMaxListe% = 0 ' 14

    ' 11/12/2022 Ne pas oublier de réinit. ce dictionnaire à chaque appel
    ' Collection de mots indexés par sMotHorsDico
    'Private m_colMots As New DicoTri(Of String, clsMot)
    Private m_colMots As DicoTri(Of String, clsMot)

    Private m_sTypeParcours, m_sTitreDoc, m_sCheminDoc, m_sTypeIndex As String
    Private m_bGlossaireInterrompu As Boolean

    Private Const sTypeParcoursMots As String = "Mots"
    Private Const sTypeParcoursParagraphes As String = "Paragraphes"
    Private Const sTypeParcoursSections As String = "Sections"
    Private Const sTypeParcoursDef As String = sTypeParcoursParagraphes

    Private Const sTriAlpha As String = "Alphabétique"
    Private Const sTriFreq As String = "Fréquentiel"
    Private Const sTriDef As String = sTriAlpha

    Private Const sIndexMotsHorsDico As String = "Hors-dico"
    Private Const sIndexMotsTous As String = "Tous"
    Private Const sIndexDef As String = sIndexMotsHorsDico

    Private m_msgDelegue As clsMsgDelegue

    Class clsMot

        ' ClsMot : classe pour indexer les mots hors dictionnaire

        Public sMot As String ' Clé de la collection : mot hors dictionnaire
        Public lNbOccurences As Integer ' Nombre d'occurrences du mot hors dictionnaire

        ' Sections ou paragraphes selon le type de parcours du document
        Public sListeSections, sMemSection As String
        Public lNbSectionsDistinctes As Integer

    End Class

    Public Function bCreerGlossaire(sCheminFichierDoc$, sCodeLangue$,
        msgDelegue As clsMsgDelegue, bTriFreq As Boolean,
        Optional ByRef bVoirGlossaireCourant As Boolean = False) As Boolean

        ' Déclaration d'un objet Word
        ' Pour utiliser la liaison précoce, on déclare bLiaisonPrecoce = -1
        '  dans les arguments de compilation conditionnelle

#If bLiaisonPrecoce Then

        ' Liaison précoce : déclaration avant la compilation
        ' C'est plus pratique lorsqu'il y a plusieurs variables liées à Word
        '  (Range, Paragraph, Section...)
        ' - avantage : intellisense pour déboguer la programmation Word
        '   (obtenir la liste des méthodes et constantes possibles sur un objet)
        ' - inconvénient : le programme requiert une référence à
        '   "Microsoft Word 10 Object Library" : Word XP (2002) doit donc être installé,
        '   sinon, le logiciel plante si vous ne changez pas la référence
        '   pour mettre votre version de Word, ou bien si Word n'est pas installé

        Dim oWrd As Word.Application = Nothing

#Else

        ' Liaison tardive : déclaration au moment de l'exécution
        ' - avantage : le programme ne plante pas si Word n'est pas installé,
        '   et un message approprié peut donc être affiché ; marche avec toutes les versions
        '   de Word qui gèrent le code VBA utilisé ici (je n'ai pas testé avec Word 97)
        ' - inconvénient : pas d'intellisense pour déboguer, et le code est moins clair

        Dim oWrd As Object = Nothing

#End If

        ' 11/12/2022 Ne pas oublier de réinit. ce dictionnaire
        m_colMots = New DicoTri(Of String, clsMot)

        m_msgDelegue = msgDelegue
        Dim LstParcours$, LstTypeIndex$
        Dim ChkAfficherGlossaire As Boolean = True
        LstParcours = sTypeParcoursDef
        LstTypeIndex = sIndexDef

        Dim sCheminGlossaireDoc$ = sDossierParent(sCheminFichierDoc) & "\Glossaire" & sExtDoc
        If Not bFichierAccessible(sCheminGlossaireDoc, bPromptFermer:=True, bInexistOk:=True) Then _
            Return False

        Dim bMemCheckSpellingAsYouType As Boolean
        Dim bMemCheckGrammarAsYouType As Boolean
        Dim bMemCheckGrammarWithSpelling As Boolean
        Dim bMemAllowAccentedUppercase As Boolean

        Dim bAucunMotHorsDico As Boolean = False
        Dim bOk As Boolean = False

        Dim oWrdH As clsHebWord = Nothing

        Try

            AfficherMessage("Ouverture de Microsoft Word...")
            oWrdH = New clsHebWord(bInterdireAppliAvant:=False)
            If IsNothing(oWrdH.oWrd) Then Return False
            oWrd = oWrdH.oWrd

            ' Optimisation : désactiver la correction grammatical
            '  avant d'ouvrir de gros documents
            ' Comment faire pour trouver les bonnes options de Word ? Réponse :
            '  utiliser l'enregistreur de macro de Word et examiner le code produit !
            With oWrd.Options
                ' Mémoriser les options pour pouvoir les rétablir une fois terminé
                bMemCheckSpellingAsYouType = .CheckSpellingAsYouType
                bMemCheckGrammarAsYouType = .CheckGrammarAsYouType
                bMemCheckGrammarWithSpelling = .CheckGrammarWithSpelling
                bMemAllowAccentedUppercase = .AllowAccentedUppercase
                .CheckSpellingAsYouType = False
                .CheckGrammarAsYouType = False
                .CheckGrammarWithSpelling = False
                .AllowAccentedUppercase = True ' 08/05/2019
            End With

            If Not bParcourirDoc(oWrd, sCheminFichierDoc, sCodeLangue, LstParcours, LstTypeIndex,
                bAucunMotHorsDico) Then Return False

            m_bGlossaireInterrompu = m_msgDelegue.m_bAnnuler
            ' Possibilité d'interrompre le parcours du document mais pas
            '  l'affichage du glossaire tel quel
            m_msgDelegue.m_bAnnuler = False
            CreerDocGlossaire(oWrd, m_colMots, bTriFreq,
                m_bGlossaireInterrompu, bMultiDocs:=False, sCodeLangue:=sCodeLangue)
            AfficherMessage("Création du glossaire terminée.")

            Const wdFormatDocument% = 0
            oWrdH.oWrd.ActiveDocument.SaveAs(sCheminGlossaireDoc, wdFormatDocument)
            bOk = True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "CreerGlossaire")
            Return False

        Finally
            If Not IsNothing(oWrdH.oWrd) Then
                ' Rétablir les options de Word
                oWrd.Options.CheckSpellingAsYouType = bMemCheckSpellingAsYouType
                oWrd.Options.CheckGrammarAsYouType = bMemCheckGrammarAsYouType
                oWrd.Options.CheckGrammarWithSpelling = bMemCheckGrammarWithSpelling
                oWrd.Options.AllowAccentedUppercase = bMemAllowAccentedUppercase

                'oWrdH.oWrd.Quit()
                ' Ne pas sauvegarder les changements s'il y a eu une erreur (document déjà ouvert)
                oWrdH.oWrd.Quit(SaveChanges:=False)
                oWrdH.oWrd = Nothing
                oWrdH.Quitter()
                oWrdH = Nothing
            End If
        End Try

        If bAucunMotHorsDico Then
            MsgBox("Aucun mot n'est absent du dictionnaire !", MsgBoxStyle.Exclamation)
        Else
            OuvrirAppliAssociee(sCheminGlossaireDoc)
        End If

        Return bOk

    End Function

    Private Function bParcourirDoc(oWrd As Object,
        sCheminFichierDoc$, sCodeLangue$, LstParcours$, LstTypeIndex$,
        ByRef bAucunMotHorsDico As Boolean) As Boolean

        Dim oColMotsHorsDico As Object ' Collection des mots hors dico
        Const wdStory% = 6

        ' Voir le glossaire courant
        'If bVoirGlossaireCourant Then GoTo CreationGlossaire

        ' Recréer le glossaire
        m_bGlossaireInterrompu = False

        ' Supprimer les mots ignorés par le correcteur orthog. dans le dictionnaire actif
        oWrd.ResetIgnoreAll()

        AfficherMessage("Ouverture du document : " & sCheminFichierDoc & "...")
        Try
            oWrd.Documents.Open(sCheminFichierDoc)
        Catch ex As Exception
            Dim sLecteur$ = sLecteurDossier(Environment.GetFolderPath(
                Environment.SpecialFolder.System))
            Dim sUser$ = Environment.UserName
            Dim sMsg$ = "Le dossier temporaire est trop volumineux !" & vbLf &
                "Il provoque une erreur 5097 (Mémoire insuffisante) dans Word." & vbLf &
                "Veuillez nettoyer le dossier caché :" & vbLf &
                sLecteur & "\Documents and Settings\" & sUser & "\Local Settings\Temp"
            CopierPressePapier(sMsg)
            MsgBox(sMsg, MsgBoxStyle.Information)
            Return False
        End Try

        Dim lNbMotsHorsDico, lNumMotHorsDico As Integer
        Dim lNbMotsHorsDicoIndexes As Integer
        Dim sMotHorsDico As String

        Dim lNbSections, lNumSection As Integer
        Dim lNbParag, lNumParag As Integer
        Dim lNumMot, lNbMots, lNbMotsIndexes As Integer
        Dim sMot As String
        With oWrd.ActiveDocument
            ' Recommencer la vérification de l'orthographe après la réinitialisation du dico
            .SpellingChecked = False

            Dim sExtension$ = IO.Path.GetExtension(sCheminFichierDoc).ToLower
            If Not (sExtension = sExtDoc OrElse sExtension.StartsWith(sExtHtm)) Then
                ' Fixer la langue dans le document (en cas d'import d'un fichier texte), 
                '  avec vérification activée
                ' Si la langue n'est pas installée dans Word, les mots hors dico ne seront pas détectés
                Dim iLang% = iConvCodeLangue%(sCodeLangue)
                .Content.LanguageID = iLang
                .Content.NoProofing = False ' Ne pas vérifier le texte : non
                oWrd.CheckLanguage = False ' Détection auto de la langue : non
            End If

            AfficherMessage("Repagination du document en cours...")
            oWrd.Selection.EndKey(Unit:=wdStory)

            ' Type de parcours d'un document :
            ' Mots        : Méthode la plus lente
            ' Paragraphes : Méthode la plus rapide
            ' Sections    : Méthode la plus sûr pour les gros documents

            Select Case LstParcours

                Case sTypeParcoursSections

                    AfficherMessage("Comptage du nombre de sections...")
                    lNbSections = .Sections.Count

                    ' Parcourir les sections du document
                    For Each oSection In .Sections

                        lNumSection = lNumSection + 1
                        AfficherMessage("Détection des mots hors dico dans la section n°" & lNumSection & " / " & lNbSections & "...")
                        If m_msgDelegue.m_bAnnuler Then Exit For

                        ' Test d'optimisation en faisant la vérification section par section
                        '  dans un nouveau document Word : cela ne change rien
                        '.Documents.Add
                        'oSection.Range.Copy
                        '.Range.Paste
                        'Set oColMotsHorsDico = .SpellingErrors

                        oColMotsHorsDico = oSection.Range.SpellingErrors
                        lNbMotsHorsDico = oColMotsHorsDico.Count
                        If lNbMotsHorsDico = 0 Then GoTo SectionSuivante

                        AfficherMessage("Parcours des mots hors dico dans la section n°" & lNumSection & " / " & lNbSections & "...")
                        If m_msgDelegue.m_bAnnuler Then Exit For

                        lNumMotHorsDico = 0
                        For Each oMotHorsDico In oColMotsHorsDico

                            lNumMotHorsDico = lNumMotHorsDico + 1
                            AfficherMessage("Indexation des mots hors dico dans la section n°" & lNumSection & " / " & lNbSections & ", mot n°" & lNumMotHorsDico & " / " & lNbMotsHorsDico & ", mots indexés : " & lNbMotsHorsDicoIndexes)
                            If m_msgDelegue.m_bAnnuler Then Exit For
                            sMotHorsDico = oMotHorsDico.Text
                            AjouterMot(m_colMots, sMotHorsDico, CStr(lNumSection))

                            lNbMotsHorsDicoIndexes = m_colMots.Count()

                        Next oMotHorsDico

                        ' Test d'optimisation : suite et fin
                        ' Enlever le message presse-papier rempli en copiant juste 1 caractère
                        '.Characters(1).Copy
                        '.Close wdDoNotSaveChanges

SectionSuivante:
                    Next oSection

                    If lNbMotsHorsDicoIndexes = 0 Then bAucunMotHorsDico = True 'GoTo AucunMotHorsDico


                Case sTypeParcoursParagraphes

                    AfficherMessage("Comptage du nombre de paragraphes...")
                    lNbParag = .Paragraphs.Count

                    For Each oParagraphe In .Paragraphs

                        lNumParag = lNumParag + 1
                        AfficherMessage("Indexation des mots hors dico : paragraphe n°" &
                            lNumParag & " / " & lNbParag & ", mots indexés : " & lNbMotsHorsDicoIndexes)
                        If m_msgDelegue.m_bAnnuler Then Exit For

                        oColMotsHorsDico = oParagraphe.Range.SpellingErrors
                        lNbMotsHorsDico = oColMotsHorsDico.Count
                        If lNbMotsHorsDico = 0 Then GoTo ParagrapheSuivant

                        lNumMotHorsDico = 0
                        For Each oMotHorsDico In oColMotsHorsDico
                            lNumMotHorsDico = lNumMotHorsDico + 1
                            sMotHorsDico = oMotHorsDico.Text
                            Dim sSection$ = lNumParag.ToString
                            AjouterMot(m_colMots, sMotHorsDico, sSection)
                        Next oMotHorsDico

                        lNbMotsHorsDicoIndexes = m_colMots.Count()

ParagrapheSuivant:
                    Next oParagraphe

                    If lNbMotsHorsDicoIndexes = 0 Then bAucunMotHorsDico = True 'GoTo AucunMotHorsDico


                Case sTypeParcoursMots

                    AfficherMessage("Comptage du nombre de mots hors dico en tout...")

                    If LstTypeIndex = sIndexMotsHorsDico Then

                        ' Indexer les mots hors dictionnaire

                        lNbMotsHorsDico = .SpellingErrors.Count
                        If lNbMotsHorsDico = 0 Then bAucunMotHorsDico = True 'GoTo AucunMotHorsDico

                        For lNumMotHorsDico = 1 To lNbMotsHorsDico
                            If lNumMotHorsDico Mod 10 = 0 Or lNumMotHorsDico = lNbMotsHorsDico Or lNumMotHorsDico = 1 Then
                                lNbMotsHorsDicoIndexes = m_colMots.Count()
                                AfficherMessage("Indexation des mots hors dico : " & lNumMotHorsDico & " / " & lNbMotsHorsDico & ", mots indexés : " & lNbMotsHorsDicoIndexes)
                                If m_msgDelegue.m_bAnnuler Then Exit For
                            End If
                            sMotHorsDico = .SpellingErrors(lNumMotHorsDico).Text
                            AjouterMot(m_colMots, sMotHorsDico)
                        Next lNumMotHorsDico

                    Else

                        ' Indexer tous les mots
                        lNbMots = .Words.Count
                        If lNbMots = 0 Then bAucunMotHorsDico = True 'GoTo AucunMotHorsDico

                        For lNumMot = 1 To lNbMots
                            If lNumMot Mod 10 = 0 Or lNumMot = lNbMots Or lNumMot = 1 Then
                                lNbMotsIndexes = m_colMots.Count()
                                AfficherMessage("Indexation des mots : " & lNumMot & " / " & lNbMots & ", mots indexés : " & lNbMotsIndexes)
                                If m_msgDelegue.m_bAnnuler Then Exit For
                            End If
                            sMot = Trim(.Words(lNumMot).Text)
                            If bSignesPonctuation(sMot) Then GoTo MotSuivant

                            AjouterMot(m_colMots, sMot)

MotSuivant:
                        Next lNumMot

                    End If

                    lNbMotsHorsDicoIndexes = m_colMots.Count()

            End Select

            ' Conserver le nom du document
            m_sTitreDoc = .Name
            m_sCheminDoc = .FullName
            m_sTypeParcours = LstParcours
            m_sTypeIndex = LstTypeIndex

            ' Fermer le document d'origine, on en a plus besoin
            .Close()
        End With
        Return True

    End Function

    'Private Sub RetablirOptionsWord(oWrd As Object, _
    '    bMemCheckSpellingAsYouType As Boolean, _
    '    bMemCheckGrammarAsYouType As Boolean, _
    '    bMemCheckGrammarWithSpelling As Boolean)

    '    If Not (oWrd Is Nothing) Then
    '        ' Rétablir les options de Word
    '        oWrd.Options.CheckSpellingAsYouType = bMemCheckSpellingAsYouType
    '        oWrd.Options.CheckGrammarAsYouType = bMemCheckGrammarAsYouType
    '        oWrd.Options.CheckGrammarWithSpelling = bMemCheckGrammarWithSpelling
    '    End If

    'End Sub

    Private Function bSignesPonctuation(ByRef sMot As String) As Boolean

        ' Indiquer si le mot ne contient que des signes de ponctuation

        Dim i, iLen As Short
        iLen = Len(sMot)
        For i = 1 To iLen
            If Not bSignePonctuation(Mid(sMot, i, 1)) Then Return False
        Next i
        Return True ' Ce mot ne contient que des signes de ponctuation

    End Function

    Private Function bSignePonctuation(ByRef sCar As String) As Boolean

        ' Indiquer si le caractère est un signe de ponctuation

        Dim iCode As Short
        iCode = Asc(sCar)
        Select Case iCode
            Case Asc("A") To Asc("Z") ' Majuscule
            Case Asc("a") To Asc("z") ' Minuscule
            Case Else ' Ponctuation et chiffre
                Return True
        End Select
        Return False

    End Function

    ' Tout ça pour conserver l'intellisense en mode débug...
#If bLiaisonPrecoce Then

    Private Sub CreerDocGlossaire(oWrd As Word.Application, m_colMots As DicoTri(Of String, clsMot), _
        ByRef bTriFreq As Boolean, ByRef bGlossaireInterrompu As Boolean, _
        ByRef bMultiDocs As Boolean, sCodeLangue$)

#Else

    Private Sub CreerDocGlossaire(ByRef oWrd As Object, ByRef m_colMots As DicoTri(Of String, clsMot),
        ByRef bTriFreq As Boolean, ByRef bGlossaireInterrompu As Boolean,
        ByRef bMultiDocs As Boolean, sCodeLangue$)

        Const wdSortFieldAlphanumeric As Integer = 0
        Const wdSortOrderAscending As Integer = 0
        Const wdSortFieldNumeric As Integer = 1
        Const wdSortOrderDescending As Integer = 1
        Const wdPageFitBestFit As Integer = 2

#End If

        ' Fabrication du glossaire à partir de la collection de mots indexés

        Dim sTitre$ = "", sMethode$ = "", sExplication$ = "", sTxtMotHorsDico$ = ""
        Dim sTitreComptage$ = "", sListeMax$ = "", sDetailExplic$ = "", sChemin$ = ""

        'If m_sCheminDoc = "" Then GoTo CreationGlossaireMultiDoc

        sTitreComptage = "Nombre de mots distincts hors dictionnaire : "
        If iNbOccurencesMaxListe > 0 Then sListeMax = " <= " & iNbOccurencesMaxListe & ")" ' & vbLf
        sTxtMotHorsDico = " hors dico"
        If m_sTypeIndex = sIndexMotsTous Then
            sTxtMotHorsDico = ""
            sTitreComptage = "Nombre de mots distincts : "
        End If
        sChemin = "Chemin : " & m_sCheminDoc

        Select Case m_sTypeParcours
            Case sTypeParcoursSections
                sMethode = "sections distinctes"

            Case sTypeParcoursParagraphes
                sMethode = "paragraphes distincts"

            Case sTypeParcoursMots
                sExplication = "Explication : Mot" & sTxtMotHorsDico & " (nombre d'occurrences)" & vbLf
                sTitre = "Glossaire de " & m_sTitreDoc
                If bTriFreq Then
                    sTitre = sTitre & " en fréquence"
                    sExplication = "Explication : Nombre d'occurrences : Mot" & sTxtMotHorsDico & vbLf
                End If
        End Select

        If m_sTypeParcours <> sTypeParcoursMots Then

            If iNbOccurencesMaxListe = 0 Then
                sDetailExplic = " (nombre d'occurrences)"
            Else
                sDetailExplic = " (nombre d'occurrences : liste des numéros de " & sMethode & sListeMax
            End If
            sExplication = "Explication : Mot" & sTxtMotHorsDico & sDetailExplic & vbLf
            sTitre = "Glossaire de " & m_sTitreDoc
            If bTriFreq Then
                sTitre = sTitre & " en fréquence"
                If iNbOccurencesMaxListe = 0 Then
                    sDetailExplic = ""
                Else
                    sDetailExplic = " (liste des numéros de " & sMethode & sListeMax
                End If
                sExplication = "Explication : Nombre d'occurrences : Mot" &
                    sTxtMotHorsDico & sDetailExplic & vbLf
            End If

        End If

        Dim lNbMotsHorsDicoIndexes As Integer
        Dim lMethodeTri, lOrdreTri As Integer
        Dim oMot As clsMot
        Dim sMotGlossaire As String
        Dim lNumMotHorsDicoIndexe As Integer
        With oWrd

            ' Ajouter la collection de mots dans un nouveau document Word
            .Documents.Add()
            lNbMotsHorsDicoIndexes = m_colMots.Count()
            lNumMotHorsDicoIndexe = 0
            For Each oMot In m_colMots.Trier()

                lNumMotHorsDicoIndexe = lNumMotHorsDicoIndexe + 1
                If lNumMotHorsDicoIndexe Mod 100 = 0 Or lNumMotHorsDicoIndexe = lNbMotsHorsDicoIndexes Or lNumMotHorsDicoIndexe = 1 Then
                    AfficherMessage("Création du " & sTitre & " : " & lNumMotHorsDicoIndexe & " / " & lNbMotsHorsDicoIndexes)
                    System.Windows.Forms.Application.DoEvents()
                    If m_msgDelegue.m_bAnnuler Then Exit For
                End If

                If bTriFreq Then
                    ' Tri fréquentiel : on met le nombre d'occurence du mot en premier
                    If oMot.sListeSections <> "" Then
                        sMotGlossaire = oMot.lNbOccurences & " : " & oMot.sMot & " (" & oMot.sListeSections & ")" & vbLf
                    Else
                        ' Cas indexation avec la boucle sur les mots '
                        '  (pas de pointeur dans ce cas)
                        sMotGlossaire = oMot.lNbOccurences & " : " & oMot.sMot & vbLf

                        ' Plus besoin de _ avec Sort FieldNumber:="Mot 1" au lieu de §
                        ' On met _ pour forcer le tri numérique sur la seule fréquence
                        '  et non sur le contenu du mot qui peut être numérique parfois
                        'sMotGlossaire = oMot.lNbOccurences & " _" & oMot.sMot & vbLf
                    End If
                Else
                    ' Tri alphabétique : on met le mot en premier
                    If oMot.sListeSections <> "" Then
                        sMotGlossaire = oMot.sMot & " (" & oMot.lNbOccurences & " : " & oMot.sListeSections & ")" & vbLf
                    Else ' Cas indexation avec la boucle sur les mots
                        sMotGlossaire = oMot.sMot & " (" & oMot.lNbOccurences & ")" & vbLf
                    End If
                End If

                .ActiveDocument.Content.InsertAfter(sMotGlossaire)

            Next oMot

            ' Trier le nouveau document Word par ordre alphabétique ou bien numérique
            AfficherMessage("Tri du glossaire...")
            .ActiveDocument.Content.WholeStory() ' Sélection de tout le document

            lMethodeTri = wdSortFieldAlphanumeric : lOrdreTri = wdSortOrderAscending
            If bTriFreq Then lMethodeTri = wdSortFieldNumeric : lOrdreTri = wdSortOrderDescending

            Dim iLang% = iConvCodeLangue%(sCodeLangue)
            If lNbMotsHorsDicoIndexes > 0 Then .ActiveDocument.Content.Sort(FieldNumber:="Mot 1",
                SortFieldType:=lMethodeTri, SortOrder:=lOrdreTri, LanguageID:=iLang)

            ' Présentation du glossaire
            If bTriFreq Then .ActiveDocument.Content.InsertBefore(vbLf)
            .ActiveDocument.Content.InsertBefore(sExplication)
            .ActiveDocument.Content.InsertBefore(sTitreComptage & lNbMotsHorsDicoIndexes & vbLf)

            ' Sinon afficher le chemin du document analysé
            If sChemin <> "" Then .ActiveDocument.Content.InsertBefore(sChemin & vbLf)

            If bGlossaireInterrompu Then
                .ActiveDocument.Content.InsertBefore("(Glossaire interrompu)" & vbLf)
            End If

            .ActiveDocument.Content.InsertBefore(sTitre & vbLf)

            ' Mettre en largeur de page maximale pour améliorer la lisibilité du document
            .ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitBestFit

            ' Fixer la langue dans le document, avec vérification activée
            oWrd.ActiveDocument.Content.LanguageID = iLang
            oWrd.ActiveDocument.Content.NoProofing = False ' Ne pas vérifier le texte : non
            .CheckLanguage = False ' Détection auto de la langue : non

        End With

    End Sub

    Public Sub AjouterMot(ByRef m_colMots As Dictionary(Of String, clsMot), sMot$,
        Optional ByRef sSection$ = "")

        ' Indexation des mots hors dico (ou bien tous les mots)
        '  pour ne conserver que les mots distincts

        ' sSection : En mode simple document, numéro de la section ou du paragraphe
        '  selon le type de parcours du document
        ' sSection : En mode multidocument, code mnémonique du document
        '  (par exemple LM pour LisezMoi)

        Dim mot As clsMot
        'Dim sCle$ = sMot.ToLower : la fonction suivante convertie en minuscule par défaut : bMinuscule:=True
        ' 19/05/2018 Cette fonction ne dépend plus de Unicode de toutes façons : bTexteUnicode:=False
        Dim sCle$ = sEnleverAccents(sMot)

        If m_colMots.ContainsKey(sCle) Then
            ' Mot déjà existant : incrémenter le nombre d'occurrences
            mot = m_colMots(sCle)
            With mot
                .lNbOccurences += 1
                If sSection <> "" And sSection <> .sMemSection Then
                    .lNbSectionsDistinctes += 1
                    .sMemSection = sSection
                    If .lNbSectionsDistinctes < iNbOccurencesMaxListe - 1 Then
                        .sListeSections &= ", " & sSection
                    ElseIf .lNbSectionsDistinctes = iNbOccurencesMaxListe - 1 Then
                        .sListeSections &= ", " & sSection & "..."
                    ElseIf .lNbSectionsDistinctes = 1 AndAlso iNbOccurencesMaxListe = 1 Then
                        .sListeSections &= "..."
                    End If
                End If
            End With
        Else
            ' Mot inexistant dans le dico, on l'ajoute
            mot = New clsMot
            With mot
                .sMot = sMot
                .lNbOccurences = 1
                If sSection <> "" Then
                    If iNbOccurencesMaxListe > 0 Then .sListeSections = sSection
                    .sMemSection = sSection
                End If
            End With
            m_colMots.Add(sCle, mot)
        End If

    End Sub

    Private Sub AfficherMessage(ByRef sMsg As String)

        'Me.LblAvancement.Text = sMsg
        ' Laisser du temps pour le traitement des messages : affichage du message et
        '  traitement du clic éventuel sur le bouton Interrompre
        'System.Windows.Forms.Application.DoEvents()
        m_msgDelegue.AfficherMsg(sMsg)

    End Sub

End Module