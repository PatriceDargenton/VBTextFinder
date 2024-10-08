VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsMot : classe pour indexer les mots

' Cl� de la collection : le mot lui-m�me
'  on est oblig� de conserver la cl� en tant que membre publique de la classe
'  car il n'existe aucun moyen d'y acc�der, par exemple dans une boucle For Each
'  (sauf en bidouillant avec des pointeurs)
Public sMot$
Public lNbOccurences&     ' Nombre d'occurrences du mot
Public lNbPhrases&        ' Nombre de phrases dans lesquelles ce mot figure
Private m_alNumPhrases&() ' Tableau des n� de phrase dans lesquelles ce mot figure

Public Function lLireNumPhrase&(lIndex&)
    
    ' L'index commence � 1 vu de l'ext�rieur de la classe
    '  (phrase n�1 = 1�re phrase) mais commence � 0 en interne
    lLireNumPhrase = m_alNumPhrases(lIndex - 1)

End Function
    
Public Sub AjouterNumPhrase(lNumPhrases&)
    
    ' Ajouter une r�f�rence de phrase (par son num�ro global) contenant ce mot
    
    lNbPhrases = lNbPhrases + 1
    ReDim Preserve m_alNumPhrases(0 To lNbPhrases - 1)
    m_alNumPhrases(lNbPhrases - 1) = lNumPhrases
    
End Sub
