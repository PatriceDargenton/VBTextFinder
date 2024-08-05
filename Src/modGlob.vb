
Module modGlob

    Public m_sTitreMsg$ = "modUtilFichier"

    Public Sub DefinirTitreApplication(sTitreMsg As String)
        m_sTitreMsg = sTitreMsg
    End Sub

    Public Function bChoisirFichier(ByRef sCheminFichier$, sFiltre$, sExtDef$,
        sTitre$, Optional sInitDir$ = "",
        Optional bDoitExister As Boolean = True,
        Optional bMultiselect As Boolean = False) As Boolean

        ' Afficher une boite de dialogue pour choisir un fichier
        ' Exemple de filtre : "|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        ' On peut indiquer le dossier initial via InitDir, ou bien via le chemin du fichier

        Static bInit As Boolean = False

        Dim ofd As New OpenFileDialog
        With ofd
            If Not bInit Then
                bInit = True
                If sInitDir.Length = 0 Then
                    If sCheminFichier.Length = 0 Then
                        .InitialDirectory = Application.StartupPath
                    Else
                        .InitialDirectory = IO.Path.GetDirectoryName(sCheminFichier)
                    End If
                Else
                    .InitialDirectory = sInitDir
                End If
            End If
            If Not String.IsNullOrEmpty(sCheminFichier) Then .FileName = sCheminFichier
            .CheckFileExists = bDoitExister ' 14/10/2007
            .DefaultExt = sExtDef
            .Filter = sFiltre
            .Multiselect = bMultiselect
            .Title = sTitre
            .ShowDialog()

            If .FileName <> "" Then sCheminFichier = .FileName : Return True
            Return False

        End With

    End Function

End Module