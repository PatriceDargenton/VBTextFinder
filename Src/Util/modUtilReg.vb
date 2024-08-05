
' Fichier modUtilReg.vb : Module de gestion de la base de registre
' ---------------------

Imports Microsoft.Win32

Module modUtilReg

    ' Microsoft Win32 to Microsoft .NET Framework API Map : Registry Functions
    ' http://msdn.microsoft.com/en-us/library/aa302340.aspx#win32map_registryfunctions

    Public Const sDossierShell$ = "shell"
    Public Const sDossierCmd$ = "command"

    Public Function bAjouterTypeFichier(sExtension$, sTypeFichier$,
        Optional sDescriptionExtension$ = "",
        Optional bEnlever As Boolean = False) As Boolean

        ' Ajouter(/Enlever) dans la base de registre un type de fichier ClassesRoot
        '  pour associer une extension de fichier à une application par défaut
        ' (via le double-clic ou bien le menu contextuel Ouvrir)
        ' Exemple : associer .dat à mon application.exe

        Try

            If bEnlever Then
                If bCleRegistreCRExiste(sExtension) Then
                    Registry.ClassesRoot.DeleteSubKeyTree(sExtension)
                End If
            Else
                If Not bCleRegistreCRExiste(sExtension) Then
                    Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sExtension)
                        rk.SetValue("", sTypeFichier)
                        If sDescriptionExtension.Length > 0 Then
                            rk.SetValue("Content Type", sDescriptionExtension)
                        End If
                    End Using 'rk.Close()
                End If
            End If
            Return True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bAjouterTypeFichier")
            Return False
        End Try

    End Function

    Public Function bAjouterMenuContextuel(sTypeFichier$, sCmd$,
        Optional bPrompt As Boolean = True,
        Optional bEnlever As Boolean = False,
        Optional sDescriptionCmd$ = "",
        Optional sCheminExe$ = "",
        Optional sCmdDef$ = """%1""",
        Optional sDescriptionTypeFichier$ = "",
        Optional bEnleverTypeFichier As Boolean = False) As Boolean

        ' Ajouter un menu contextuel dans la base de registre
        '  de type ClassesRoot : fichier associé à une application standard
        ' Exemple : ajouter le menu contextuel "Convertir en Html" sur les fichiers projet VB6
        ' sTypeFichier = "VisualBasic.Project"
        ' sCmd = "ConvertirEnHtml"
        ' sDescriptionCmd = "Convertir en Html"
        ' sCheminExe = "C:\Program Files\VB2Html\VB2Html.exe"

        Try

            ' D'abord vérifier si la clé principale existe
            If Not bCleRegistreCRExiste(sTypeFichier) Then
                If bEnlever Then bAjouterMenuContextuel = True : Exit Function
                Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sTypeFichier)
                    If sDescriptionTypeFichier.Length > 0 Then
                        rk.SetValue("", sDescriptionTypeFichier)
                    End If
                End Using
            End If

            Dim sCleDescriptionCmd$ = sTypeFichier & "\" & sDossierShell & "\" & sCmd

            If bEnlever Then

                If bEnleverTypeFichier Then
                    ' Si c'est un type de fichier créé à l'occasion
                    '  il faut aussi le supprimer (mais seulement dans ce cas)
                    If bCleRegistreCRExiste(sTypeFichier) Then
                        Registry.ClassesRoot.DeleteSubKeyTree(sTypeFichier)
                        If bPrompt Then _
                            MsgBox("Le type de fichier [" & sTypeFichier & "]" & vbLf &
                                "a été enlevé avec succès dans la base de registre",
                            MsgBoxStyle.Information, m_sTitreMsg)
                    Else
                        If bPrompt Then _
                            MsgBox("Le type de fichier [" & sTypeFichier & "]" & vbLf &
                                "est introuvable dans la base de registre",
                            MsgBoxStyle.Information, m_sTitreMsg)
                    End If
                Else

                    If bCleRegistreCRExiste(sCleDescriptionCmd) Then
                        Registry.ClassesRoot.DeleteSubKeyTree(sCleDescriptionCmd)
                        If bPrompt Then _
                            MsgBox("Le menu contextuel [" & sDescriptionCmd & "]" & vbLf &
                                "a été enlevé avec succès dans la base de registre pour les fichiers du type :" & vbLf &
                                "[" & sTypeFichier & "]",
                            MsgBoxStyle.Information, m_sTitreMsg)
                    Else
                        If bPrompt Then _
                            MsgBox("Le menu contextuel [" & sDescriptionCmd & "]" & vbLf &
                                "est introuvable dans la base de registre pour les fichiers du type :" & vbLf &
                                "[" & sTypeFichier & "]",
                            MsgBoxStyle.Information, m_sTitreMsg)
                    End If

                End If
                bAjouterMenuContextuel = True
                Exit Function
            End If

            Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sCleDescriptionCmd)
                rk.SetValue("", sDescriptionCmd)
            End Using 'rk.Close()

            Dim sCleCmd$ = sTypeFichier & "\" & sDossierShell & "\" & sCmd & "\" & sDossierCmd
            Using rk As RegistryKey = Registry.ClassesRoot.CreateSubKey(sCleCmd)
                ' Ajouter automatiquement des guillemets " si le chemin contient au moins un espace
                If sCheminExe.IndexOf(" ") > -1 Then _
                    sCheminExe = """" & sCheminExe & """"
                rk.SetValue("", sCheminExe & " " & sCmdDef)
            End Using 'rk.Close()

            If bPrompt Then _
                MsgBox("Le menu contextuel [" & sDescriptionCmd & "]" & vbLf &
                    "a été ajouté avec succès dans la base de registre pour les fichiers du type :" & vbLf &
                    "[" & sTypeFichier & "]", MsgBoxStyle.Information, m_sTitreMsg)

            Return True

        Catch ex As Exception
            AfficherMsgErreur2(ex, "bAjouterMenuContextuel",
                "Cause possible : L'application doit être lancée en tant qu'admin. pour cette opération.")
            Return False
        End Try

    End Function

    Public Function bCleRegistreCRExiste(sCle$,
        Optional sSousCle$ = "") As Boolean

        ' Vérifier si une clé ClassesRoot existe dans la base de registre
        ' Note : la sous-clé est ici un "sous-dossier" (et non un "fichier")

        Try
            ' Si la clé n'existe pas, on passe dans le Catch
            Using rkCRCle As RegistryKey = Registry.ClassesRoot.OpenSubKey(
                sCle & "\" & sSousCle)

                ' Liste des sous-clés (sous forme de "sous-dossier")
                'Dim asListeSousClesCR$() = rkCRCle.GetSubKeyNames

                If IsNothing(rkCRCle) Then Return False
            End Using ' rkCRCle.Close() est automatiquement appelé
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function bCleRegistreCRExiste(sCle$, sSousCle$,
        ByRef sValSousCle$) As Boolean

        ' Vérifier si une clé ClassesRoot existe dans la base de registre
        '  et si elle est trouvée, alors lire la valeur de la sous-clé
        ' Renvoyer True si la valeur de la sous-clé a pu être lue
        ' Note : la sous-clé est ici un "fichier" (et non un "sous-dossier")

        sValSousCle = ""
        Try
            ' Si la clé n'existe pas, on passe dans le Catch
            Using rkCRCle As RegistryKey = Registry.ClassesRoot.OpenSubKey(sCle)
                If IsNothing(rkCRCle) Then Return False
                ' Pour lire la valeur par défaut d'un "dossier", laisser ""
                Dim oVal As Object = rkCRCle.GetValue(sSousCle)
                ' Si la sous-clé n'existe pas, oVal reste à Nothing 
                '  (aucune exception n'est générée)
                If IsNothing(oVal) Then Return False
                Dim sValSousCle0$ = CStr(oVal)
                ' Il faut aussi tester ce cas obligatoirement
                If IsNothing(sValSousCle0) Then Return False
                sValSousCle = sValSousCle0
            End Using ' rkCRCle.Close() est automatiquement appelé
            Return True
        Catch
            Return False
        End Try

    End Function

    Public Function bCleRegistreLMExiste(sCle$,
        Optional sSousCle$ = "",
        Optional ByRef sValSousCle$ = "",
        Optional sNouvValSousCle$ = "") As Boolean

        ' Vérifier si une clé/sous-clé LocalMachine existe dans la base de registre
        sValSousCle = ""
        Try
            Dim bEcriture As Boolean = False
            If sNouvValSousCle.Length > 0 Then bEcriture = True
            ' Si la clé n'existe pas, on passe dans le Catch
            Using rkLMCle As RegistryKey = Registry.LocalMachine.OpenSubKey(sCle,
                writable:=bEcriture)
                ' Lecture de la valeur de la sous-clé (sous forme de "fichier")
                Dim oVal As Object = rkLMCle.GetValue(sSousCle)

                ' Liste des sous-clés (sous forme de "sous-dossier")
                'Dim asListeSousClesLM$() = rkLMCle.GetSubKeyNames

                ' Si la sous-clé n'existe pas, oVal reste à Nothing 
                '  (aucune exception n'est générée)
                If IsNothing(oVal) Then Return False
                Dim sValSousCle0$ = CStr(oVal)
                ' Il faut aussi tester ce cas obligatoirement
                If IsNothing(sValSousCle0) Then Return False
                sValSousCle = sValSousCle0
                If bEcriture Then
                    oVal = CInt(sNouvValSousCle)
                    rkLMCle.SetValue(sSousCle, oVal, RegistryValueKind.DWord)
                End If
            End Using ' rkLMCle.Close() est automatiquement appelé
            Return True ' On peut lire cette clé, donc elle existe
        Catch
            Return False
        End Try

    End Function

    Public Function bCleRegistreCUExiste(sCle$,
        Optional sSousCle$ = "",
        Optional ByRef sValSousCle$ = "") As Boolean

        ' Vérifier si une clé/sous-clé CurrentUser existe dans la base de registre
        '  et si oui renvoyer la valeur de la sous-clé
        sValSousCle = ""
        Try
            ' Si la clé n'existe pas, on passe dans le Catch
            Using rkCUCle As RegistryKey = Registry.CurrentUser.OpenSubKey(sCle)
                Dim oVal As Object = rkCUCle.GetValue(sSousCle)
                ' Si la sous-clé n'existe pas, oVal reste à Nothing 
                '  (aucune exception n'est générée)
                If IsNothing(oVal) Then Return False
                Dim sValSousCle0$ = CStr(oVal)
                ' Il faut aussi tester ce cas obligatoirement
                If IsNothing(sValSousCle0) Then Return False
                sValSousCle = sValSousCle0
            End Using ' rkCUCle.Close() est automatiquement appelé
            Return True ' On peut lire cette clé, donc elle existe
        Catch
            Return False
        End Try

    End Function

    Public Function asListeSousClesCU(sCle$) As String()

        ' Renvoyer la liste des sous-clés de type CurrentUser
        asListeSousClesCU = Nothing
        Try
            ' Si la clé n'existe pas, on passe dans le Catch
            Using rkCUCle As RegistryKey = Registry.CurrentUser.OpenSubKey(sCle)
                If IsNothing(rkCUCle) Then Exit Function
                asListeSousClesCU = rkCUCle.GetSubKeyNames
            End Using ' rkCUCle.Close() est automatiquement appelé
        Catch
        End Try

    End Function

End Module