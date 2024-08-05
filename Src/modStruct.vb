
Friend Class clsDoc

    ' clsDoc : classe pour indexer la liste des documents indexés

    ' Clé de la collection : code mnémonique du document indexé
    '  (ce code est précisé dans les résultats de recherche)
    Public sCle$ ' sCodeDoc
    Public sCodeDoc$ ' Clé éditée dans le fichier ini (nouveau !)
    Public sChemin$ ' Chemin du document indexé
    Public bTxtUnicode As Boolean ' Encodage unicode ? sinon encodage par défaut 26/01/2019
    ' Nombre de mots indexés du document indexé (pas encore utilisé)
    'Public lNbMotsIndexes&

    Public colChapitres As New Collection
    'Public colChapitres As Collection

End Class

Friend Class clsChapitre

    ' clsChapitre : classe pour indexer la liste des chapitres des documents indexés

    ' Clé de la collection : code mnémonique du chapitre
    '  (ce code est précisé dans les résultats de recherche)
    Public sCle$ ' CleDoc:CodeChapitre
    Public sCodeChapitre$
    Public sCleDoc$ ' Clé d'origine du document : Doc n°1, ...
    Public sCodeDoc$ ' Clé éditée du document dans le fichier ini
    Public sChapitre$ ' Chemin du document indexé

End Class

Friend Class clsPhrase

    ' clsPhrase : classe pour indexer les phrases

    Public sClePhrase$ ' Clé de la collection : numéro de la phrase global
    Public iNumPhraseG% ' Numéro de la phrase global des documents indexés
    Public iNumPhraseL% ' Numéro de la phrase local au document indexé
    Public sPhrase$ ' Phrase stockée en intégralité
    Public sCleDoc$ ' Code mnémonique du document dans lequel figure la phrase
    Public sCodeChapitre$ ' 19/06/2010
    Public iNumParagrapheL% ' Numéro du paragraphe local dans lequel figure la phrase
    Public iNumParagrapheG% ' Numéro du paragraphe global dans lequel figure la phrase

End Class

Friend Class clsMot

    ' clsMot : classe pour indexer les mots

    ' Clé de la collection : le mot lui-même
    '  on est obligé de conserver la clé en tant que membre publique de la classe
    '  car il n'existe aucun moyen (en VB6) d'y accéder, par exemple dans une boucle For Each
    '  (sauf en bidouillant avec des pointeurs)
    Public sMot$
    Public iNbOccurrences% ' Nombre d'occurrences du mot
    'Public lNbPhrases% ' Nombre de phrases dans lesquelles ce mot figure
    'Private m_alNumPhrases%() ' Tableau des n° de phrase dans lesquelles ce mot figure
    ' Si le mot figure plusieurs fois dans la même phrase, 
    '  on duplique quand même le n° de phrase
    Public aiNumPhrase As New ArrayList

    Public Function iLireNumPhrase%(ByRef iIndex%)

        ' L'index commence à 1 vu de l'extérieur de la classe
        '  (phrase n°1 = 1ère phrase) mais commence à 0 en interne
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

        ' Ajouter une référence de phrase contenant ce mot
        Me.aiNumPhrase.Add(iNumPhrases)

    End Sub

End Class

Friend Class clsBiGramme

    ' clsBiGramme : classe pour comptabiliser la fréquence des bigrammes

    ' Clé de la collection : le bigramme
    Public sBiGramme$
    Public iNbOccurences% ' Nombre d'occurrences du bigramme

End Class