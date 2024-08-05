Attribute VB_Name = "modConfig"
Option Explicit

' Module de configuration

Public Const sTitreMsg$ = "VBTextFinder"

' Nombre de références maximum indiquées pour chaque mot du document index
Public Const iNbOccurrencesMaxListe% = 14

' Nombre de références maximum recherchées (pour les mots trop fréquents)
Public Const iNbOccurencesMaxRecherchees% = 100

Public Const sListeSeparateurPhrase$ = ".:?!;|"
Public Const sListeSeparateurMot$ = sListeSeparateurPhrase & _
    " ,&~'()[]{}<>`’-+±*/\@=°%#¦$€£§"
' Séparateurs de mot supplémentaires
Public Const iCodeASCIITabulation% = 9
Public Const iCodeASCIIEspaceInsecable% = 160
Public Const iCodeASCIIGuillement% = 34  '"
Public Const iCodeASCIIGuillementOuvrant% = 171
Public Const iCodeASCIIGuillementFermant% = 187

Public Const iModuloAvanvement% = 100 ' Affichage périodique de l'avancement

' Faire une sauvegarde de sécurité à chaque indexation d'un nouveau document
' sFichierVBTxtFndTmp = "VBTxtFnd.tmp"
Public Const bSauvegardeSecurite As Boolean = False




