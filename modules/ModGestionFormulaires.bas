Attribute VB_Name = "ModGestionFormulaires"
'******************************************************************************
'***    DataSource Copyright                                                                       ***
'******************************************************************************
'***    MODULE:                                                                                          ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***   Constante Pour L'Etat D'Un Formulaire                                               ***
'******************************************************************************
Public Const ETATCREATION = 0
Public Const ETATFORMULAIRE = 1
Public Const ETATFEUILLEDONNER = 2
Public Const ETATFERMER = 3

'******************************************************************************
'***    Constante Pour Le Mode De Saisie                                                    ***
'******************************************************************************
Public Const MODENULL = 0
Public Const MODEAJOUT = 1
Public Const MODEMODIF = 2
Public Const MODESUPPRIME = 3
Public Const MODEVALIDE = 4
Public Const MODEDEPLACEMENT = 5

'******************************************************************************
'***    Constante Pour Le Rafraichissement                                                 ***
'******************************************************************************
Public Const RAFRAICHIRVISIBLE = 0
Public Const RAFRAICHIRCHOIX = 1
Public Const RAFRAICHIRDEPLACEMENT = 2

'******************************************************************************
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Les Filtres                                                                   ***
'******************************************************************************
Public GStrSocietesFiltre As String
Public GStrContactsFiltre As String
Public GStrCadeauxAchatsFiltre As String
Public GStrChaudieresRemonteesFiltre As String

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function DonneFiltre(ByVal StrFiltre As String) As String
   Dim RsFiltre As DAO.Recordset

   On Error GoTo Err_DonneFiltre

   DonneFiltre = vbNullString

   StrFiltre = Mid(StrFiltre, InStr(1, StrFiltre, "In") + 4, Len(StrFiltre) - InStr(1, StrFiltre, "In") - 4)

   Set RsFiltre = CurrentDb.OpenRecordset(StrFiltre)

   Select Case RsFiltre.EOF
      Case True

      Case False

         Do Until RsFiltre.EOF = True

            DonneFiltre = DonneFiltre & RsFiltre.Fields(0) & ";"

            RsFiltre.MoveNext

         Loop

            DonneFiltre = Mid(DonneFiltre, 1, Len(DonneFiltre) - 1)

   End Select

   RsFiltre.Close

Exit_DonneFiltre:

   Set RsFiltre = Nothing

   Exit Function

Err_DonneFiltre:

   DonneFiltre = vbNullString

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltre"

   Resume Exit_DonneFiltre
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function DonneFiltreElement(ByVal StrFiltre As String, ByRef StrFiltreTableRequeteNom As String, ByRef StrFiltreChampSource As String, ByRef StrFiltreChampDestination As String, ByRef StrFiltreWhere As String) As Boolean

   On Error GoTo Err_DonneFiltreElement

   DonneFiltreElement = True

   Select Case InStr(InStr(1, StrFiltre, "FROM") + 5, StrFiltre, " ")
      Case 0

         StrFiltreTableRequeteNom = Mid(StrFiltre, InStr(1, StrFiltre, "FROM") + 5, InStr(InStr(1, StrFiltre, "FROM") + 5, StrFiltre, ";") - InStr(1, StrFiltre, "FROM") - 5)

      Case Else

         StrFiltreTableRequeteNom = Mid(StrFiltre, InStr(1, StrFiltre, "FROM") + 5, InStr(InStr(1, StrFiltre, "FROM") + 5, StrFiltre, " ") - InStr(1, StrFiltre, "FROM") - 5)

   End Select

   StrFiltreChampSource = Left(StrFiltre, InStr(1, StrFiltre, " ") - 1)

   StrFiltreChampDestination = Mid(StrFiltre, InStr(1, StrFiltre, ".") + 1, InStr(InStr(1, StrFiltre, "."), StrFiltre, " FROM") - InStr(1, StrFiltre, ".") - 1)

   Select Case InStr(1, StrFiltre, "WHERE")
      Case 0

         StrFiltreWhere = vbNullString

      Case Else

         StrFiltreWhere = Mid(Left(StrFiltre, InStrRev(StrFiltre, ";") - 1), InStr(1, StrFiltre, "WHERE") + 6)

   End Select

Exit_DonneFiltreElement:

   Exit Function

Err_DonneFiltreElement:

   DonneFiltreElement = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltreElement"

   Resume Exit_DonneFiltreElement
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ChangeBoutonMode(ByRef FrmForm As Access.Form, ByVal IntMode As Integer) As Boolean

   On Error GoTo Err_ChangeBoutonMode

   ChangeBoutonMode = True

   FrmForm.IntModeOld = FrmForm.IntMode

   FrmForm.IntMode = IntMode

   Select Case IntMode
      Case MODEAJOUT

         FrmForm.AllowAdditions = True

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 65280

      Case MODEMODIF

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = True

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 16711680

      Case MODESUPPRIME

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = True

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 255

      Case MODEDEPLACEMENT

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = True

         FrmForm.CmdAjouter.Enabled = True

         FrmForm.CmdModifier.Enabled = True

         FrmForm.CmdSupprimer.Enabled = True

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = True

         FrmForm.CmdAfficherTous.Enabled = True

         FrmForm.RecButton.BackStyle = vbEmpty

         FrmForm.RecButton.BackColor = 255

      Case MODEVALIDE

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 65535

   End Select

Exit_ChangeBoutonMode:

   Exit Function

Err_ChangeBoutonMode:

   ChangeBoutonMode = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ChangeBoutonMode"

   Resume Exit_ChangeBoutonMode
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function FermeFormulaire(Optional ByVal StrFormName As String) As Boolean
   Dim FrmForm As Access.Form

   On Error GoTo Err_FermeFormulaire

   For Each FrmForm In Forms

      Select Case FrmForm.Name
         Case StrFormName

         Case Else

            DoCmd.Close acForm, FrmForm.Name, acSavePrompt

            FermeFormulaire = True

      End Select

   Next

Exit_FermeFormulaire:

   Exit Function

Err_FermeFormulaire:

   FermeFormulaire = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "FermeFormulaire"

   Resume Exit_FermeFormulaire
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function DonneEtatFormulaire(ByVal StrFormName As String, ByRef intFormEtat As Integer) As Boolean
   Dim FrmDataBase As Access.Forms
   Dim FrmForm As Access.Form

   On Error GoTo Exit_DonneEtatFormulaire

   Set FrmDataBase = Forms

   DonneEtatFormulaire = False

   For Each FrmForm In FrmDataBase

      Select Case FrmForm.Name
         Case StrFormName

            intFormEtat = FrmForm.CurrentView

            DonneEtatFormulaire = True

            GoTo Exit_DonneEtatFormulaire

         Case Else

      End Select
   Next

   intFormEtat = ETATFERMER

   DonneEtatFormulaire = False

Exit_DonneEtatFormulaire:

   Exit Function

Err_DonneEtatFormulaire:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneEtatFormulaire"

   Resume Exit_DonneEtatFormulaire
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreMenuPrincipale() As Boolean

   On Error GoTo Err_OuvreMenuPrincipale

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         DoCmd.OpenForm "FrmMenuPrincipale", acNormal, , , acFormEdit, acWindowNormal

      Case "CALL"

         DoCmd.OpenForm "FrmMenuCall", acNormal, , , acFormEdit, acWindowNormal

   End Select

Exit_OuvreMenuPrincipale:

   Exit Function

Err_OuvreMenuPrincipale:

   OuvreMenuPrincipale = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreMenuPrincipale"

   Resume Exit_OuvreMenuPrincipale
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreRecherches(Optional ByVal bytTypeOuvre As Byte) As Boolean

   On Error GoTo Err_OuvreRecherches

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         FermeFormulaire "FrmMenuPrincipale"

      Case "CALL"

         FermeFormulaire "FrmMenuCall"

   End Select

   DoCmd.OpenForm "FrmRecherche", acNormal, , , acFormEdit

   Select Case bytTypeOuvre
      Case 0

      Case Else

         Forms!FrmRecherche!TbsRecherche.Value = bytTypeOuvre - 1

   End Select

Exit_OuvreRecherches:

   Exit Function

Err_OuvreRecherches:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreRecherches"

   Resume Exit_OuvreRecherches
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreRecherchesMulti() As Boolean

   On Error GoTo Err_OuvreRecherchesMulti

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         FermeFormulaire "FrmMenuPrincipale"

      Case "CALL"

         FermeFormulaire "FrmMenuCall"

   End Select

   DoCmd.OpenForm "FrmRechercheMulti", acNormal, , , acFormEdit

Exit_OuvreRecherchesMulti:

   Exit Function

Err_OuvreRecherchesMulti:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreRecherchesMulti"

   Resume Exit_OuvreRecherchesMulti
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreCadeaux(ByVal blnTypeOuvre As Boolean, Optional ByVal Num) As Boolean
   Dim LngCadAchNum As Long
   Dim LngCadAchSocNum As Long
   Dim StrSQLFiltre As String
   Dim IntEtat As Integer

   On Error GoTo Err_OuvreCadeaux

   OuvreCadeaux = True

   Select Case IsMissing(Num)
      Case True

      Case False

         Select Case VarType(Num)
            Case vbInteger, vbLong

               Select Case blnTypeOuvre
                  Case True

                     LngCadAchSocNum = Num

                  Case False

                     LngCadAchNum = Num

               End Select
            Case Else

         End Select
   End Select

   Select Case LngCadAchSocNum
      Case Is > vbEmpty

         StrSQLFiltre = "CadAchSocNum=" & LngCadAchSocNum

         DonneEtatFormulaire "FrmSaisieCadeauxAchats", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieCadeauxAchats.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreCadeaux

               End Select

               On Error GoTo Err_OuvreCadeaux

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieCadeauxAchats", acNormal, , StrSQLFiltre, acFormEdit

         End Select

         Forms!FrmSaisieCadeauxAchats.LstCadAchSocNum.DefaultValue = LngCadAchSocNum

         GoTo Exit_OuvreCadeaux

      Case Else

   End Select

   Select Case LngCadAchNum
      Case Is > vbEmpty

         StrSQLFiltre = "CadAchNum=" & LngCadAchNum

         DonneEtatFormulaire "FrmSaisieCadeauxAchats", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieCadeauxAchats.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreCadeaux

               End Select

               On Error GoTo Err_OuvreCadeaux

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieCadeauxAchats", acNormal, , StrSQLFiltre, acFormEdit

         End Select
      Case Else

   End Select

Exit_OuvreCadeaux:

   Exit Function

Err_OuvreCadeaux:

   OuvreCadeaux = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreCadeaux"

   Resume Exit_OuvreCadeaux
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreChaudieres(ByVal blnTypeOuvre As Boolean, Optional ByVal Num) As Boolean
   Dim LngChdRemNum As Long
   Dim LngChdRemSocNum As Long
   Dim StrSQLFiltre As String
   Dim IntEtat As Integer

   On Error GoTo Err_OuvreChaudieres

   OuvreChaudieres = True

   Select Case IsMissing(Num)
      Case True

      Case False

         Select Case VarType(Num)
            Case vbInteger, vbLong

               Select Case blnTypeOuvre
                  Case True

                     LngChdRemSocNum = Num

                  Case False

                     LngChdRemNum = Num

               End Select
            Case Else

         End Select
   End Select

   Select Case LngChdRemSocNum
      Case Is > vbEmpty

          StrSQLFiltre = "ChdRemSocNum=" & LngChdRemSocNum

          DonneEtatFormulaire "FrmSaisieChaudieresRemontees", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieChaudieresRemontees.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreChaudieres

               End Select

               On Error GoTo Err_OuvreChaudieres

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieChaudieresRemontees", acNormal, , StrSQLFiltre, acFormEdit

         End Select

         Forms!FrmSaisieChaudieresRemontees.LstChdRemSocNum.DefaultValue = LngChdRemSocNum

         GoTo Exit_OuvreChaudieres

      Case Else

   End Select

   Select Case LngChdRemNum
      Case Is > vbEmpty

         StrSQLFiltre = "ChdRemNum=" & LngChdRemNum

         DonneEtatFormulaire "FrmSaisieChaudieresRemontees", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieChaudieresRemontees.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreChaudieres

               End Select

               On Error GoTo Err_OuvreChaudieres

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieChaudieresRemontees", acNormal, , StrSQLFiltre, acFormEdit

         End Select
      Case Else

   End Select

Exit_OuvreChaudieres:

   Exit Function

Err_OuvreChaudieres:

   OuvreChaudieres = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreChaudieres"

   Resume Exit_OuvreChaudieres
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreBonusAnimation(ByVal blnTypeOuvre As Boolean, Optional ByVal Num) As Boolean
   Dim LngMvtNum As Long
   Dim LngMvtSocNum As Long
   Dim StrSQLFiltre As String
   Dim IntEtat As Integer

   On Error GoTo Err_OuvreBonusAnimation

   OuvreBonusAnimation = True

   Select Case IsMissing(Num)
      Case True

      Case False

         Select Case VarType(Num)
            Case vbInteger, vbLong

               Select Case blnTypeOuvre
                  Case True

                     LngMvtSocNum = Num

                  Case False

                     LngMvtNum = Num

               End Select
            Case Else

         End Select
   End Select

   Select Case LngMvtSocNum
      Case Is > vbEmpty

          StrSQLFiltre = "MvtSocNum=" & LngMvtSocNum

          DonneEtatFormulaire "FrmSaisieBonusAnimation", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieBonusAnimation.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreBonusAnimation

               End Select

               On Error GoTo Err_OuvreBonusAnimation

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieBonusAnimation", acNormal, , StrSQLFiltre, acFormEdit

         End Select

         Forms!FrmSaisieBonusAnimation.LstMvtSocNum.DefaultValue = LngMvtSocNum

         GoTo Exit_OuvreBonusAnimation

      Case Else

   End Select

   Select Case LngMvtNum
      Case Is > vbEmpty

         StrSQLFiltre = "MvtNum=" & LngMvtNum

         DonneEtatFormulaire "FrmSaisieBonusAnimation", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieBonusAnimation.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreBonusAnimation

               End Select

               On Error GoTo Err_OuvreBonusAnimation

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieBonusAnimation", acNormal, , StrSQLFiltre, acFormEdit

         End Select
      Case Else

   End Select

Exit_OuvreBonusAnimation:

   Exit Function

Err_OuvreBonusAnimation:

   OuvreBonusAnimation = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreBonusAnimation"

   Resume Exit_OuvreBonusAnimation
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreRegularisationCheques(ByVal blnTypeOuvre As Boolean, Optional ByVal Num) As Boolean
   Dim LngMvtNum As Long
   Dim LngMvtSocNum As Long
   Dim StrSQLFiltre As String
   Dim IntEtat As Integer

   On Error GoTo Err_OuvreRegularisationCheques

   OuvreRegularisationCheques = True

   Select Case IsMissing(Num)
      Case True

      Case False

         Select Case VarType(Num)
            Case vbInteger, vbLong

               Select Case blnTypeOuvre
                  Case True

                     LngMvtSocNum = Num

                  Case False

                     LngMvtNum = Num

               End Select
            Case Else

         End Select
   End Select

   Select Case LngMvtSocNum
      Case Is > vbEmpty

          StrSQLFiltre = "MvtSocNum=" & LngMvtSocNum

          DonneEtatFormulaire "FrmSaisieRegularisationCheques", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieRegularisationCheques.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreRegularisationCheques

               End Select

               On Error GoTo Err_OuvreRegularisationCheques

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieRegularisationCheques", acNormal, , StrSQLFiltre, acFormEdit

         End Select

         Forms!FrmSaisieRegularisationCheques.LstMvtSocNum.DefaultValue = LngMvtSocNum

         GoTo Exit_OuvreRegularisationCheques

      Case Else

   End Select

   Select Case LngMvtNum
      Case Is > vbEmpty

         StrSQLFiltre = "MvtNum=" & LngMvtNum

         DonneEtatFormulaire "FrmSaisieRegularisationCheques", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieRegularisationCheques.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLFiltre

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_OuvreRegularisationCheques

               End Select

               On Error GoTo Err_OuvreRegularisationCheques

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmSaisieRegularisationCheques", acNormal, , StrSQLFiltre, acFormEdit

         End Select
      Case Else

   End Select

Exit_OuvreRegularisationCheques:

   Exit Function

Err_OuvreRegularisationCheques:

   OuvreRegularisationCheques = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreRegularisationCheques"

   Resume Exit_OuvreRegularisationCheques
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteCadeauxArticles(Optional ByVal StrReference As String) As Boolean

   On Error GoTo Err_AjouteCadeauxArticles

   AjouteCadeauxArticles = True

  Select Case MsgBox("Voulez Vous Créer Une Nouvelle Reference ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DoCmd.Close acForm, "FrmSaisieCadeauxArticles", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCadeauxArticles", acNormal, , , acFormAdd

         Select Case StrReference
            Case vbNullString

            Case Else

               Forms!FrmSaisieCadeauxArticles.TxtCadArtRef = StrReference

         End Select

      Case vbNo

   End Select

Exit_AjouteCadeauxArticles:

   Exit Function

Err_AjouteCadeauxArticles:

   AjouteCadeauxArticles = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteCadeauxArticles"

   Resume Exit_AjouteCadeauxArticles
End Function

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Sub AjouteChaudieresArticles(Optional ByVal StrReference As String)

   On Error GoTo Err_AjouteChaudieresArticles

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Reference ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DoCmd.Close acForm, "FrmSaisieChaudieresArticles", acSavePrompt

         DoCmd.OpenForm "FrmSaisieChaudieresArticles", acNormal, , , acFormAdd

         Select Case StrReference
            Case vbNullString

            Case Else

               Forms!FrmSaisieChaudieresArticles.TxtChdArtRef = StrReference

         End Select

      Case vbNo

   End Select

Exit_AjouteChaudieresArticles:

   Exit Sub

Err_AjouteChaudieresArticles:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteChaudieresArticles"

   Resume Exit_AjouteChaudieresArticles
End Sub

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteCivilites(Optional ByVal StrCivilite As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteCivilites

   AjouteCivilites = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Civilités ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstCivilites", IntEtat

         Select Case IntEtat
           Case ETATFORMULAIRE

               Forms!FrmLstCivilites.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "CivCode=" & StrCivilite

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_AjouteCivilites

               End Select

               GoTo Err_AjouteCivilites

               DoCmd.RunCommand acCmdRecordsGoToNew

               DoCmd.RunCommand acCmdWindowCascade

           Case Else

              DoCmd.OpenForm "FrmLstCivilites", acNormal, , , acFormAdd

         End Select

         Select Case StrCivilite
            Case vbNullString

            Case Else

               Forms!FrmLstCivilites.TxtCivCode = StrCivilite

         End Select

      Case vbNo

   End Select

Exit_AjouteCivilites:

   Exit Function

Err_AjouteCivilites:

   AjouteCivilites = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteCivilites"

   Resume Exit_AjouteCivilites
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteVilles(Optional ByVal strCodePostal As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteVilles

   AjouteVilles = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Villes ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstVilles", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmLstVilles.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "VilCodePostal=" & strCodePostal

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_AjouteVilles

               End Select

               On Error GoTo Err_AjouteVilles

               DoCmd.RunCommand acCmdRecordsGoToNew

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmLstVilles", acNormal, , , acFormAdd

         End Select

         Select Case strCodePostal
            Case vbNullString

            Case Else

               Forms!FrmLstVilles.TxtVilCodePostal = strCodePostal

         End Select

      Case vbNo

   End Select

Exit_AjouteVilles:

   Exit Function

Err_AjouteVilles:

   AjouteVilles = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteVilles"

   Resume Exit_AjouteVilles
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteFonctions(Optional ByVal strFonction As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteFonctions

   AjouteFonctions = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Fonctions ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstContactsFonctions", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmLstContactsFonctions.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "ConFonCode=" & strFonction

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_AjouteFonctions

               End Select

               On Error GoTo Err_AjouteFonctions

               DoCmd.RunCommand acCmdRecordsGoToNew

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmLstContactsFonctions", acNormal, , , acFormAdd

         End Select

         Select Case strFonction
            Case vbNullString

            Case Else

               Forms!FrmLstContactsFonctions.TxtConFonCode = strFonction

         End Select
      Case vbNo

   End Select

Exit_AjouteFonctions:

   Exit Function

Err_AjouteFonctions:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteFonctions"

   Resume Exit_AjouteFonctions
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteSocietes(Optional ByVal StrRaison As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteSocietes

   AjouteSocietes = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Sociétés ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
           Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=0"

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_AjouteSocietes

               End Select

               On Error GoTo Err_AjouteSocietes

               DoCmd.RunCommand acCmdWindowCascade

           Case Else

              DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=0", acFormEdit, acWindowNormal

       End Select

         Select Case StrRaison
            Case vbNullString

            Case Else

               Forms!FrmFicheSocietes.TxtSocRaison.DefaultValue = """ & StrRaison & """

         End Select

      Case vbNo

   End Select

Exit_AjouteSocietes:

   Exit Function

Err_AjouteSocietes:

   AjouteSocietes = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteSocietes"

   Resume Exit_AjouteSocietes
End Function
