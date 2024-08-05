Attribute VB_Name = "modConfig"
Option Explicit

' Module de configuration

Public Const sTitreMsg$ = "VBTextFinder"

' Nombre de r�f�rences maximum indiqu�es pour chaque mot du document index
Public Const iNbOccurrencesMaxListe% = 14

' Nombre de r�f�rences maximum recherch�es (pour les mots trop fr�quents)
Public Const iNbOccurencesMaxRecherchees% = 100

Public Const sListeSeparateurPhrase$ = ".:?!;|"
Public Const sListeSeparateurMot$ = sListeSeparateurPhrase & _
    " ,&~'()[]{}<>`�-+�*/\@=�%#�$���"
' S�parateurs de mot suppl�mentaires
Public Const iCodeASCIITabulation% = 9
Public Const iCodeASCIIEspaceInsecable% = 160
Public Const iCodeASCIIGuillement% = 34  '"
Public Const iCodeASCIIGuillementOuvrant% = 171
Public Const iCodeASCIIGuillementFermant% = 187

Public Const iModuloAvanvement% = 100 ' Affichage p�riodique de l'avancement

' Faire une sauvegarde de s�curit� � chaque indexation d'un nouveau document
' sFichierVBTxtFndTmp = "VBTxtFnd.tmp"
Public Const bSauvegardeSecurite As Boolean = False




