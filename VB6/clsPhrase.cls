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

Public sClePhrase$     ' Clé de la collection : numéro de la phrase global
Public lNumPhraseG&    ' Numéro de la phrase global des documents indexés
Public lNumPhraseL&    ' Numéro de la phrase local au document indexé
Public sPhrase$        ' Phrase stockée en intégralité
Public sCodeDoc$       ' Code mnémonique du document dans lequel figure la phrase
Public lNumParagraphe& ' Numéro du paragraphe dans lequel figure la phrase
