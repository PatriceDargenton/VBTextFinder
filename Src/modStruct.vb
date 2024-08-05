
Friend Class clsDoc

    ' clsDoc : classe pour indexer la liste des documents index�s

    ' Cl� de la collection : code mn�monique du document index�
    '  (ce code est pr�cis� dans les r�sultats de recherche)
    Public sCle$ ' sCodeDoc
    Public sCodeDoc$ ' Cl� �dit�e dans le fichier ini (nouveau !)
    Public sChemin$ ' Chemin du document index�
    Public bTxtUnicode As Boolean ' Encodage unicode ? sinon encodage par d�faut 26/01/2019
    ' Nombre de mots index�s du document index� (pas encore utilis�)
    'Public lNbMotsIndexes&

    Public colChapitres As New Collection
    'Public colChapitres As Collection

End Class

Friend Class clsChapitre

    ' clsChapitre : classe pour indexer la liste des chapitres des documents index�s

    ' Cl� de la collection : code mn�monique du chapitre
    '  (ce code est pr�cis� dans les r�sultats de recherche)
    Public sCle$ ' CleDoc:CodeChapitre
    Public sCodeChapitre$
    Public sCleDoc$ ' Cl� d'origine du document : Doc n�1, ...
    Public sCodeDoc$ ' Cl� �dit�e du document dans le fichier ini
    Public sChapitre$ ' Chemin du document index�

End Class

Friend Class clsPhrase

    ' clsPhrase : classe pour indexer les phrases

    Public sClePhrase$ ' Cl� de la collection : num�ro de la phrase global
    Public iNumPhraseG% ' Num�ro de la phrase global des documents index�s
    Public iNumPhraseL% ' Num�ro de la phrase local au document index�
    Public sPhrase$ ' Phrase stock�e en int�gralit�
    Public sCleDoc$ ' Code mn�monique du document dans lequel figure la phrase
    Public sCodeChapitre$ ' 19/06/2010
    Public iNumParagrapheL% ' Num�ro du paragraphe local dans lequel figure la phrase
    Public iNumParagrapheG% ' Num�ro du paragraphe global dans lequel figure la phrase

End Class

Friend Class clsMot

    ' clsMot : classe pour indexer les mots

    ' Cl� de la collection : le mot lui-m�me
    '  on est oblig� de conserver la cl� en tant que membre publique de la classe
    '  car il n'existe aucun moyen (en VB6) d'y acc�der, par exemple dans une boucle For Each
    '  (sauf en bidouillant avec des pointeurs)
    Public sMot$
    Public iNbOccurrences% ' Nombre d'occurrences du mot
    'Public lNbPhrases% ' Nombre de phrases dans lesquelles ce mot figure
    'Private m_alNumPhrases%() ' Tableau des n� de phrase dans lesquelles ce mot figure
    ' Si le mot figure plusieurs fois dans la m�me phrase, 
    '  on duplique quand m�me le n� de phrase
    Public aiNumPhrase As New ArrayList

    Public Function iLireNumPhrase%(ByRef iIndex%)

        ' L'index commence � 1 vu de l'ext�rieur de la classe
        '  (phrase n�1 = 1�re phrase) mais commence � 0 en interne
        'iLireNumPhrase = m_alNumPhrases(lIndex - 1)
        iLireNumPhrase = DirectCast(Me.aiNumPhrase.Item(iIndex - 1), Integer)

    End Function

    Public Function iNbPhrases%()

        iNbPhrases = Me.aiNumPhrase.Count

    End Function

    Public Sub RedimPhrases(iNbPhrases0%)

        'lNbPhrases = lNbPhrases0
        'ReDim m_alNumPhrases(lNbPhrases - 1)
        Me.aiNumPhrase = New ArrayList(iNbPhrases0)

    End Sub

    Public Sub AjouterNumPhrase3(iNumPhrases%)

        ' Ajouter une r�f�rence de phrase contenant ce mot
        Me.aiNumPhrase.Add(iNumPhrases)

    End Sub

End Class

Friend Class clsBiGramme

    ' clsBiGramme : classe pour comptabiliser la fr�quence des bigrammes

    ' Cl� de la collection : le bigramme
    Public sBiGramme$
    Public iNbOccurences% ' Nombre d'occurrences du bigramme

End Class