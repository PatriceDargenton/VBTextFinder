
' Fichier modConfig.vb : Module de configuration
' ----------------------

Module Config

    Public Const bSupprimerEspInsec As Boolean = False ' 06/03/2016 Faire une option ?

    ' 05/05/2018 Nouveau dictionnaire basé sur frgut + DELA + LibreOffice :
    ' DELA : http://infolingu.univ-mlv.fr/DonneesLinguistiques/Dictionnaires/telechargement.html
    ' LibreOffice : www.dicollecte.org
    'Public Const sCheminDicoV1Fr$ = "\Dico\liste.de.mots.francais.frgut.txt"

    Public Const sCheminDico$ = "\Dico\Dico" '_Fr.txt"
    ' Trop long dans l'explorateur !
    'Public Const sURLDico$ = "http://www.pallier.org/ressources/dicofr/liste.de.mots.francais.frgut.txt" 
    ' 800 Ko : ok !
    'Public Const sURLDicoFr$ = "http://patrice.dargenton.free.fr/CodesSources/VBTextFinder/DicoVBTF.zip"
    ' 05/05/2018
    Public Const sURLDicoFr$ = "http://patrice.dargenton.free.fr/CodesSources/VBTextFinder/Dico_Fr.zip"

    ' AGID is an Automatically Generated Inflection Database from an insanely large word list.
    ' http://downloads.sourceforge.net/wordlist/agid-4.zip
    Public Const sURLDicoEn$ = "http://patrice.dargenton.free.fr/CodesSources/VBTextFinder/Dico_En.zip"
    ' Ce sont les mêmes dico. pour l'instant
    Public Const sURLDicoUk$ = "http://patrice.dargenton.free.fr/CodesSources/VBTextFinder/Dico_Uk.zip"
    Public Const sURLDicoUs$ = "http://patrice.dargenton.free.fr/CodesSources/VBTextFinder/Dico_Us.zip"

    Public Const bCompatVB6RechercheAussiAvecAccents As Boolean = True

    ' sMotsCourants ne contient pas les accents : 
    '  les mots clés ne fonctionneront plus si on indexe les accents
    ' Si Unicode alors conserver les accents et tous les caractères exotiques
    'Public bIndexerAccents As Boolean = False ' Ignorer    les accents
    'Public Const bIndexerAccents As Boolean = False ' Ignorer    les accents
    'Public Const bIndexerAccents As Boolean = True ' Distinguer les accents

    ' Exporter tous les documents avec les n° de § et de ph. global et local
    ' (pour vérifier que l'affichage des n° fonctionne bien)
    Public Const bExporterToutAvecNumeros As Boolean = False

    ' Nombre de références maximum indiquées pour chaque mot du document index
    Public Const iNbOccurrencesMaxListe% = 12

    ' Nombre de références maximum recherchées (pour les mots trop fréquents)
    Public Const iNbOccurencesMaxRecherchees% = 100

    Public Const iNbCarChapitreMax% = 8 '10 '5

    ' Si le fichier externe est présent alors il remplace la liste codée en dur
    Public Const sCheminSeparateursPhrase$ = "\Dico\SeparateursPhrase.txt"
    Public Const sCheminSeparateursMot$ = "\Dico\SeparateursMot.txt"
    Public Const sListeSeparateursPhrase$ = ".:?!;|¡¿"
    Public Const sListeSeparateursMot$ = " ,&~'`´‘’()[]{}<>–-+±*/¦\@=°%#$€£§…"

    Public Const sCheminChapitrage$ = "\Dico\Chapitrage.txt"
    Public Const sCheminChapitrageExcel$ = "\Dico\ChapitrageExcel.txt"
    Public Const sCheminChapitrageAccess$ = "\Dico\ChapitrageAccess.txt"
    Public Const sChapitrageDef$ = "Chapitre;Chap;Livre;Livre"
    Public Const sChapitrageXLDef$ = "Feuille Excel n°;Feuil."
    Public Const sChapitrageMdbDef$ =
        "Structure Table Access n°;Struc.Table;" &
        "Table Access n°;Table;" &
        "Module VBA Access n°;ModVBA;" &
        "Formulaire VBA Access n°;FrmVBA;" &
        "Etat VBA Access n°;EtatVBA;" &
        "Définition Requête Access n°;DefRq;" &
        "Requête Access n°;Rq;" &
        "Requêtes systèmes Access;RqSys"

    Public Const iMaxMotsClesDef% = 50

    ' Si le fichier externe est présent alors il remplace la liste codée en dur
    ' 28/08/2009 On tient maintenant compte du code langue : Si Fr : MotsCourants_Fr.txt
    Public Const sCheminMotsCourants$ = "\Dico\MotsCourants" '.txt"
    Public Const sMotsCourantsFr$ = " de la le l et les est à dans il que nous en des qui du d un une se ce qu ne pour a pas avec au par vous je n s c sont on ils sur ces tout plus ou cette son mais même si moi elle notre comme y tous lui être leur ses ont sa sans alors très peut aux celui ainsi où toutes mon ceux me bien dit fait tu grand doit deux toute quand cela nos était car j leurs autre lorsque aussi faut etc avons toujours donc autres dire grande chose jusqu là devons entre etre temps après cet jamais m faire parce votre ai chaque mêmes vers beaucoup rien été avoir elles fois avait eux maintenant seulement encore ni trouve sous fut sommes jour quelque non mes suis dont contre sera soit afin peu avant ma ceci ci moment point état tant devant ici t toi lorsqu or veut déjà ton aucun celle vos avez êtes selon "

    ' Séparateurs de mot supplémentaires : ne figurent pas dans la première liste
    Public Const iCodeASCIITabulation% = 9

    ' https://murviel-info.com/specialchars.php
    Public Const iCodeASCIIEspaceInsecable% = 160 ' Non-breaking space &nbsp;

    ' 13/07/2019 Rétabli pour le mode Unicode
    Public Const iCodeUTF16EspaceInsecable% = 8201 ' Alt+8201 espace fine &thinsp;

    ' Cocher l'option Unicode pour pouvoir utiliser ces car.:
    Public Const iCodeUTF16EspaceFineInsecable% = 8239 ' Alt+8239 = 0x202F = espace fine insécable
    'Public Const iCodeASCIIEspaceInsecable4% = 8194 ' espace demi-cadratin &ensp;
    'Public Const iCodeASCIIEspaceInsecable5% = 8195 ' espace cadratin &emsp;
    'Public Const iCodeASCIIEspaceInsecable6% = 255 ' 

    Public Const iCodeASCIIGuillemet% = 34 ' "
    Public Const iCodeASCIIQuote% = 39 '
    Public Const iCodeASCIIGuillemetOuvrant% = 171 ' «
    Public Const iCodeASCIIGuillemetFermant% = 187 ' »
    Public Const iCodeASCIIGuillemetOuvrant2% = 145 ' ‘
    Public Const iCodeASCIIGuillemetFermant2% = 146 ' ’
    Public Const iCodeASCIIGuillemetOuvrant3% = 147 ' “
    Public Const iCodeASCIIGuillemetFermant3% = 148 ' ”
    Public Const iCodeASCIIGuillemetOuvrant4% = 96 ' `
    Public Const iCodeASCIIGuillemetFermant4% = 180 ' ´
    Public Const iCodeASCIIGuillemetOuvrant5% = 139 ' ‹
    Public Const iCodeASCIIGuillemetFermant5% = 155 ' ›
    Public Const sGm$ = Chr(iCodeASCIIGuillemet)

    Public Const iModuloAvanvementTresLent% = 10000
    Public Const iModuloAvanvementLent% = 1000
    Public Const iModuloAvanvement% = 100 ' Affichage périodique de l'avancement
    Public Const iModuloAvanvementRapide% = 10

    ' Faire une sauvegarde de sécurité à chaque indexation d'un nouveau document
    ' sFichierVBTxtFndTmp = "VBTxtFnd.tmp"
    Public Const bSauvegardeSecurite As Boolean = False

    Public Const bTestComplexifieur As Boolean = False
    Public Const iComplexifieurMinRecherche% = 3
    Public Const iComplexifieurMaxRecherche% = 5
    Public Const sComplexifieurs3$ = " ure oir age "
    Public Const sComplexifieurs4$ = " cité isme naire ogie ance ible tion "
    Public Const sComplexifieurs5$ = " ilité iaire ateur sseur ement " 'logie tible 
    'Public Const sComplexifieurs6$ = " ssible " 'ssance nement 
    'Public Const sComplexifieurs7$ = " ssement "

    'Public Const iNbCouleursHtml% = 5
    'http://htmlhelp.com/cgi-bin/color.cgi
    Public Const sCouleursHtmlDef$ = "yellow;lightgreen;lightblue;silver;turquoise"

End Module