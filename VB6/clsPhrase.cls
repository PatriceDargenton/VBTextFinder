VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhrase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsPhrase : classe pour indexer les phrases

Public sClePhrase$     ' Cl� de la collection : num�ro de la phrase global
Public lNumPhraseG&    ' Num�ro de la phrase global des documents index�s
Public lNumPhraseL&    ' Num�ro de la phrase local au document index�
Public sPhrase$        ' Phrase stock�e en int�gralit�
Public sCodeDoc$       ' Code mn�monique du document dans lequel figure la phrase
Public lNumParagraphe& ' Num�ro du paragraphe dans lequel figure la phrase
