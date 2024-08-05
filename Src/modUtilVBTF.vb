
' Fichier modUtilVBTF.vb : Module utilitaire pour VBTextFinder
' ----------------------

Imports System.Text ' Pour StringBuilder
'Imports System.Text.Encoding ' Pour GetEncoding
Imports System.Runtime.CompilerServices ' Pour l'attribut <Extension()>

Module modUtilVBTF

    ' Longueur max. d'une chaîne dans un contrôle
    Public Const iMaxLongChaine0 As Short = 32767 '32000 En VB6, 7 ?, 8 ? mais plus 9 !
    'Public Const iMaxLongChaine% = iMaxLongChaine0
    ' 01/05/2010 (147 483 647 au lieu de 32767)
    Public Const iMaxLongChaine% = Int32.MaxValue

    Public Function bEcrireChaine(bw As IO.BinaryWriter, sChaine$) As Boolean

        ' Ecrire une chaîne de longueur variable dans un fichier binaire

        Dim iLongChaine As Short
        iLongChaine = CShort(sChaine.Length)
        bw.Write(iLongChaine)
        bw.Write(sChaine.ToCharArray())
        bEcrireChaine = True

    End Function

    Public Function bLireChaine(br As IO.BinaryReader, ByRef sChaine$) As Boolean

        ' Lire une chaîne de longueur variable dans un fichier binaire
        '  pour cela, il faut d'abord sauvegarder la longueur de la chaîne
        ' On utilise ByRef pour éviter de réallouer la chaîne en RAM

        Dim iLongChaine As Short ' Int16
        ' Lire d'abord la longueur de la chaîne
        iLongChaine = br.ReadInt16()
        'FileGet(iNumFichier, iLongChaine)
        'If iLongChaine <= 0 Then Exit Function
        ' C'est surement une erreur si la chaîne est trop longue
        If iLongChaine > iMaxLongChaine0 Then
            Return False
        End If
        'sChaine = Space(iLongChaine) ' = String(iLongChaine, " ")
        'sChaine = br.ReadString ' Ne fonctionne pas toujours avec les accents
        sChaine = br.ReadChars(iLongChaine)
        If sChaine.Length <> iLongChaine Then
            Return False
        End If
        'FileGet(iNumFichier, sChaine)
        Return True

    End Function

    Public Function bCarNumerique(cCar As Char) As Boolean

        ' Vérifier si le car. est numérique (romain ou pas)

        bCarNumerique = False
        If Char.IsDigit(cCar) Then
            bCarNumerique = True
        Else
            Dim cCarMaj As Char = Char.ToUpper(cCar)
            Dim iPosCarRomain% = cCarMaj.ToString.IndexOfAny("IVXLCD".ToCharArray)
            If iPosCarRomain > -1 Then bCarNumerique = True
        End If

    End Function

    Public Function sRognerDernierCar$(sTexte$, sCar$)

        Dim sTexte2$ = sTexte.TrimEnd
        If sTexte2.EndsWith(sCar) Then
            sRognerDernierCar = Left$(sTexte2, sTexte2.Length - 1)
        Else
            sRognerDernierCar = sTexte2
        End If

    End Function

    Public Function Reverse$(s$)
        Dim charArray() As Char = s.ToCharArray
        Array.Reverse(charArray)
        Return New String(charArray)
    End Function

    Public Function sSupprimerNumeriquesEnFinDeMot$(sMot$)

        ' Si le dernier car. n'est pas numérique retourner le mot tel quel
        Dim cDernierCar As Char = sMot.ElementAt(sMot.Length - 1)
        If Not Char.IsDigit(cDernierCar) Then Return sMot

        Dim sMotInv$ = Reverse(sMot)
        Dim iPosDernierNumInv% = 0
        For Each cCar As Char In sMotInv.ToCharArray()
            If Not Char.IsDigit(cCar) Then Exit For
            iPosDernierNumInv += 1
        Next
        If iPosDernierNumInv = 0 Then Return sMot
        Dim iPosDernierNum% = sMot.Length - iPosDernierNumInv
        ' Si le mot est entièrement numérique, alors le laisser tel quel
        If iPosDernierNum = 0 Then Return sMot
        Dim sMotSansNum$ = sMot.Substring(0, iPosDernierNum)
        Return sMotSansNum

    End Function

    <Extension()>
    Public Function IndexOfUppercase%(sTexte$, Optional iDeb% = 0)

        ' Retourner l'index de la 1ère majuscule trouvée après iDeb, sinon -1
        Dim bTrouve As Boolean = False
        Dim iIdx% = 0
        For Each cChar0 In sTexte
            If iIdx >= iDeb AndAlso Char.IsUpper(cChar0) Then bTrouve = True : Exit For
            iIdx += 1
        Next
        If Not bTrouve Then Return -1
        Return iIdx

    End Function

    <Extension()>
    Public Function VBSplit(sTexte$, acSepMot() As Char) As String()

        ' Découper le texte sTexte en tableau de String, selon les séparateurs indiqués
        ' (comme la fonction String.Split)

        Dim ac = sTexte.ToCharArray
        Dim lst As New List(Of String)
        Dim sb As New StringBuilder
        Dim iNbCar% = sTexte.Length
        Dim iNumCar% = 0
        For Each c In ac
            iNumCar += 1
            Dim bSep As Boolean = False
            For Each cSep In acSepMot
                If c = cSep Then bSep = True : Exit For
            Next
            If bSep Then
                lst.Add(sb.ToString)
                If iNumCar = iNbCar Then lst.Add("")
                sb = New StringBuilder
            Else
                sb.Append(c)
            End If
        Next
        If sb.Length > 0 Then lst.Add(sb.ToString)
        Dim astr = lst.ToArray
        Return astr

    End Function

End Module