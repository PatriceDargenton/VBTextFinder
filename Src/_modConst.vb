
Module _modConst

    Public ReadOnly sNomAppli$ = My.Application.Info.Title '"VBTextFinder"
    Public ReadOnly sTitreMsg$ = sNomAppli
    Private Const sDateVersionVBTF$ = "05/08/2024"
    Public Const sDateVersionAppli$ = sDateVersionVBTF

#If DEBUG Then
    ' Pour la suite, bDebug est plus simple à écrire que #If Debug
    Public Const bDebug As Boolean = True
#Else
        Public Const bDebug As Boolean = False
#End If

    Public Const sExtDoc$ = ".doc"
    Public Const sExtHtm$ = ".htm"
    Public Const sExtTxt$ = ".txt"

    Public Const sMsgOperationTerminee$ = "Opération terminée."

    Public Const iIndiceNulString% = -1

End Module