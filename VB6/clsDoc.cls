VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsDoc : classe pour indexer la liste des documents index�s

' Cl� de la collection : code mn�monique du document index�
'  (ce code est pr�cis� dans les r�sultats de recherche)
Public sCle$    ' sCodeDoc
Public sCodeDoc$ ' Cl� �dit�e dans le fichier ini (nouveau !)
Public sChemin$ ' Chemin du document index�
' Nombre de mots index�s du document index� (pas encore utilis�)
'Public lNbMotsIndexes&
