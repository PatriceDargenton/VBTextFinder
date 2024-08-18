# VBTextFinder
Un moteur de recherche de mot dans son contexte
---

VBTextFinder permet de retrouver l'ensemble des phrases (et paragraphes) contenant un mot donné que l'on recherche dans un corpus documentaire préalablement indexé (un corpus est un ensemble cohérent de documents). Pour retrouver dans quel document a été trouvé chaque occurrence, un code mnémonique unique par document est stocké dans un fichier ini.

## Table des matières
- [Utilisation](#utilisation)
- [Fonctionnalités](#fonctionnalités)
- [Limitations](#limitations)
- [Projets](#projets)
- [Versions](#versions)
- [Liens](#liens)

## Utilisation
Lancer VBTextFinder en mode administrateur, ajouter le menu contextuel via le bouton + dans l'onglet Config., quitter VBTextFinder, puis selectionner un dossier ou un fichier Word dans l'explorateur de fichier, et avec le bouton droit de la souris, faite "Indexer pour une recherche (VBTF)" (avec Windows 11, il faudra appuyer sur la touche majuscule pour afficher directement ce menu, sinon il faudra d'abord afficher le menu Autres options).

## Fonctionnalités

- Exportation : les résultats de recherche peuvent être copiés / collés dans le presse papier ;
- Choix du nombre de paragraphes à afficher avant ou après chaque phrase afin de mieux saisir le contexte dans certain cas ;
- Gestion des fichiers Word .doc ou bien Html : il y a une procédure pour convertir automatiquement le fichier en .txt en utilisant l'automation de Word ; on peut aussi utiliser les menus contextuels dans l'explorateur de Windows afin d'indexer rapidement un document à la volée ;
- Affichage d'un indicateur si le mot est trouvé directement pendant la frappe, avec le nombre d'occurrences correspondant ;
- Mémorisation de l'historique des mots recherchés, seulement pour une session donnée ;
- Mode hypertexte : il est activé lorsque l'on fait un double-clic sur un mot dans les phrases affichées dans la zone des résultats de recherche ;
- Sauvegarde : si au moment de quitter l'application on souhaite conserver l'index, une sauvegarde de l'index est effectuée dans le fichier VBTextFinder.idx. Au lancement de VBTextFinder, si ce fichier est présent (dans le dossier de l'application), on propose de le recharger ;
- Création d'un index des mots dans Word (avec la liste des codes document le contenant) selon un tri alphabétique ou fréquentiel au choix ;
- Recherche d'expressions ;
- Affichage des occurrences en html, et mise en évidence en couleurs ou en gras des occurrences ;
- Gestion des accents en option (les accents sont ignorés par défaut) ;
- Sélection de plusieurs fichiers à indexer (ou a convertir au préalable), par exemple *.txt, *.doc ou *.html ;
- Création d'une liste de mot clés : il s'agit de la liste des n mots les plus fréquents, une fois que l'on a retiré les mots non signifiants tel que : de le la et... ;
- Gestion d'un dictionnaire : en décochant "Mots dico" les index seront générés sans les mots qui font partie du dictionnaire, ce qui permet de lister tous les mots propres indexés ;
- Gestion des mots courants : en décochant "Mots courants" les index seront générés sans les mots courants ;
- Outil pour extraire les citations ;
- Détection des chapitres.

## Limitations
- Le surlignage des mots trouvés est parfois un peu décallé, du fait de l'encodage des lettres sur un ou deux octets.

## Projets

## Versions

Voir le [Changelog.md](Changelog.md)

## Liens

Documentation d'origine complète : [VBTextFinder.html](http://patrice.dargenton.free.fr/CodesSources/VBTextFinder.html)