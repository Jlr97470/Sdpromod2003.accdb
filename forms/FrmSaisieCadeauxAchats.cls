Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    DataSource Copyright                                                                       ***
'******************************************************************************
'***    FORM:                                                                                              ***
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
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable De Controle De Mode De Saisie                                            ***
'******************************************************************************
Public IntMode As Integer
Public IntModeOld As Integer

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable De Controle De Mode De Rafraichissement                           ***
'******************************************************************************
Private TxtCadAchSocNum As String
Private TxtCadAchCadArtNum As String

'******************************************************************************
'***    Declaration De Procedure Priver                                                       ***
'******************************************************************************

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
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

              On Error Resume Next

               Set CtlActive = Me.ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = LstCadAchCadArtNum

              End Select

              On Error GoTo Err_Rafraichir

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

              Select Case TxtCadAchSocNum
                  Case vbNullString

                     LstCadAchSocNum.Requery

                  Case Else

                     LstCadAchSocNum.Requery

                     LstCadAchSocNum.Text = TxtCadAchSocNum

                     TxtCadAchSocNum = vbNullString

              End Select

              Select Case TxtCadAchCadArtNum
                  Case vbNullString

                     LstCadAchCadArtNum.Requery

                  Case Else

                     LstCadAchCadArtNum.Requery

                     LstCadAchCadArtNum.Text = TxtCadAchCadArtNum

                     TxtCadAchCadArtNum = vbNullString

              End Select

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

              CtlActive.SetFocus

              Set CtlActive = Nothing

            Case Else

         End Select

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

'******************************************************************************
'***    Declaration D'Evenement Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   DoCmd.RunCommand acCmdWindowCascade

   ChangeBoutonMode Me, MODEDEPLACEMENT

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Open"

   Resume Exit_Form_Open
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Err_Form_Unload

   Select Case IntMode
      Case MODEDEPLACEMENT

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

   End Select

Exit_Form_Unload:

   Exit Sub

Err_Form_Unload:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Unload"

   Resume Exit_Form_Unload
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)

   On Error GoTo Err_Form_ApplyFilter

   Select Case IntMode
      Case MODEDEPLACEMENT

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

   End Select

Exit_Form_ApplyFilter:

   Exit Sub

Err_Form_ApplyFilter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_ApplyFilter"

   Resume Exit_Form_ApplyFilter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)
   Dim LngSoldePoints As Long
   Dim LngPoints As Long

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsNull(CadAchSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Societe Code STA Est Invalide !", , Me.Caption

         LstCadAchSocNum.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsDate(CadAchDateSaisie)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Date Saisie Est Invalide !", , Me.Caption

         TxtCadAchDateSaisie.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsDate(CadAchHeureSaisie)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Heure Saisie Est Invalide !", , Me.Caption

         TxtCadAchHeureSaisie.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsDate(CadAchDate)
      Case True

         Select Case CadAchDateSaisie > CadAchDate
            Case False

            Case True

               Cancel = True

               MsgBox "Le Champ Cadeaux Achats Date Est Invalide !", , Me.Caption

               TxtCadAchDate.SetFocus

               GoTo Exit_Form_BeforeUpdate
         End Select
      Case False

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Date Est Invalide !", , Me.Caption

         TxtCadAchDate.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsNull(CadAchCadArtNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Articles Reference Est Invalide !", , Me.Caption

         LstCadAchCadArtNum.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsNull(CadAchQuantite)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Quantite Est Invalide !", , Me.Caption

         TxtCadAchQuantite.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsNull(CadAchPoints)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Cadeaux Achats Points Est Invalide !", , Me.Caption

         TxtCadAchPoints.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case LstCadAchEtat
      Case "Commandé"

         Select Case IntModeOld
            Case MODEAJOUT

            Case MODEMODIF

               LngPoints = Nz(DLookup("CadAchPoints", "TBLCADEAUXACHATS", "CadAchNum=" & CadAchNum), vbEmpty)

         End Select

         LngSoldePoints = Nz(DLookup("Solde", "SelSocietesSoldes", "MvtSocNum=" & CadAchSocNum), vbEmpty) + Nz(DLookup("PointsRea2003", "TBLSOCIETESRESULTATSBONUS", "SocNum=" & CadAchSocNum), vbEmpty) + Nz(DLookup("PointsRea2004", "TBLSOCIETESRESULTATSBONUS", "SocNum=" & CadAchSocNum), vbEmpty)

         Select Case CadAchPoints <= (LngSoldePoints + LngPoints)
            Case True

            Case False

               Select Case MsgBox("La Societées " & LstCadAchSocNum.Column(2) & " Ne Possédent Pas Assez De Points : " & _
                  LngSoldePoints & " Points !" & vbCrLf & "Voulez-Vous Enregistrer La Commande ?", vbQuestion + vbYesNo, Me.Caption)

                  Case vbYes

                  Case vbNo

                     Cancel = True

                     LstCadAchSocNum.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

         End Select

      Case Else

   End Select

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_BeforeUpdate"

   Resume Exit_Form_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

   LstCadAchSocNum.SetFocus

   Select Case IntModeOld
      Case MODEAJOUT

         AjouteMouvementsCadeaux , CadAchNum

      Case MODEMODIF

         SupprimeMouvementsCadeaux , CadAchNum

         AjouteMouvementsCadeaux CadAchSocNum

   End Select

   ChangeBoutonMode Me, MODEDEPLACEMENT

   LstCadAchCadArtNum.SetFocus

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_AfterUpdate:

   Exit Sub

Err_Form_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_AfterUpdate"

   Resume Exit_Form_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Current()

   On Error GoTo Exit_Form_Current

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Current"

   Resume Exit_Form_Current
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Activate()

   On Error GoTo Exit_Form_Activate

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_Activate:

   Exit Sub

Err_Form_Activate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Activate"

   Resume Exit_Form_Activate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouter_Click()
   Dim TxtDateSaisie As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         LstCadAchCadArtNum.SetFocus

         LstCadAchSocNum.DefaultValue = LstCadAchSocNum

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateSaisie = TxtCadAchDateSaisie

   TxtCadAchDateSaisie = "01/01/2003"

   TxtCadAchDateSaisie = TxtDateSaisie

   LstCadAchCadArtNum.SetFocus

Exit_CmdAjouter_Click:

   Exit Sub

Err_CmdAjouter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouter_Click"

   Resume Exit_CmdAjouter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdModifier_Click()
   Dim TxtDateSaisie As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstCadAchSocNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateSaisie = TxtCadAchDateSaisie

         TxtCadAchDateSaisie = "01/01/2003"

         TxtCadAchDateSaisie = TxtDateSaisie

         LstCadAchCadArtNum.SetFocus

   End Select

Exit_CmdModifier_Click:

   Exit Sub

Err_CmdModifier_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdModifier_Click"

   Resume Exit_CmdModifier_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSupprimer_Click()
   Dim LngCadAchSocNum As Long
   Dim LngCadAchNum As Long

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstCadAchSocNum.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         LngCadAchSocNum = CadAchSocNum

         LngCadAchNum = CadAchNum

         SupprimeMouvementsCadeaux , LngCadAchNum

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Select Case Err.Number
                  Case 2501

                  Case Else

                     MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSupprimer_Click"

               End Select

               Err.Clear

         End Select

         On Error GoTo Err_CmdSupprimer_Click

         AjouteMouvementsCadeaux LngCadAchSocNum

         ChangeBoutonMode Me, MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

            Case Else

               LstCadAchCadArtNum.SetFocus

         End Select

   End Select

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSupprimer_Click"

   Resume Exit_CmdSupprimer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   LstCadAchSocNum.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

   LstCadAchCadArtNum.SetFocus

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Select Case Err.Number
            Case 2501, 2759

            Case Else

               MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdValider_Click"

         End Select

         Err.Clear

         ChangeBoutonMode Me, IntModeOld

   End Select

   On Error GoTo Err_CmdValider_Click

Exit_CmdValider_Click:

   Exit Sub

Err_CmdValider_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdValider_Click"

   Resume Exit_CmdValider_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAnnuler_Click()

   On Error GoTo Err_CmdAnnuler_Click

   LstCadAchSocNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstCadAchCadArtNum.SetFocus

   End Select

Exit_CmdAnnuler_Click:

   Exit Sub
Err_CmdAnnuler_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAnnuler_Click"

   Resume Exit_CmdAnnuler_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.Close acForm, Me.Name, acSavePrompt

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

   End Select

   On Error GoTo Err_CmdFermer_Click

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherSocietes_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   Select Case IsNull(CadAchSocNum)
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & CadAchSocNum

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_CmdAfficherSocietes_Click

               End Select

               On Error GoTo Err_CmdAfficherSocietes_Click

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & CadAchSocNum, acFormEdit

         End Select

   End Select

Exit_CmdAfficherSocietes_Click:

   Exit Sub

Err_CmdAfficherSocietes_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherSocietes_Click"

   Resume Exit_CmdAfficherSocietes_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdExporteExcel_Click()
   Dim StrSQLRequete As String
   Dim RsCadeauxAchats As DAO.Recordset

   On Error GoTo Err_CmdExporteExcel

   Select Case Me.RecordsetClone.RecordCount = 0
      Case True

      Case False

         StrSQLRequete = Me.RecordSource

         Select Case Me.FilterOn
            Case True

               Select Case InStr(1, StrSQLRequete, "WHERE")
                  Case 0

                     StrSQLRequete = Left(StrSQLRequete, InStrRev(StrSQLRequete, ";") - 1) & " WHERE " & IIf(Left(Me.Filter, 1) = Chr(34) And Right(Me.Filter, 1) = Chr(34), Mid(Me.Filter, 2, Len(Me.Filter) - 2), Me.Filter) & ";"

                  Case Else

                     StrSQLRequete = Left(StrSQLRequete, InStr(1, StrSQLRequete, "WHERE") + 5) & " " & IIf(Left(Me.Filter, 1) = Chr(34) And Right(Me.Filter, 1) = Chr(34), Mid(Me.Filter, 2, Len(Me.Filter) - 2), Me.Filter) & Mid(StrSQLRequete, InStr(1, StrSQLRequete, "WHERE") + 5)

               End Select

            Case False

         End Select

         On Error Resume Next

         Set RsCadeauxAchats = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case Err.Number
            Case 0

               RsCadeauxAchats.Close

               ExporteExcelRequete (StrSQLRequete)

            Case Else

               Err.Clear

               MsgBox "Impossible D'Exporter " & IIf(Me.FilterOn, "Avec Le Filtre En Cours", vbNullString) & " !!!"

         End Select

   End Select

Exit_CmdExporteExcel:

   Set RsCadeauxAchats = Nothing

   Exit Sub

Err_CmdExporteExcel:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteExcel"

   Resume Exit_CmdExporteExcel
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.PrintOut acSelection

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterCadeaux_Click()

   On Error GoTo Err_CmdAjouterCadeaux_Click

   AjouteCadeauxArticles

Exit_CmdAjouterCadeaux_Click:

   Exit Sub

Err_CmdAjouterCadeaux_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterCadeaux_Click"

   Resume Exit_CmdAjouterCadeaux_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherTous_Click()

   On Error Resume Next

   Select Case Me.Filter
      Case vbNullString

      Case Else

         Me.Filter = vbNullString

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

         End Select

         Me.Requery

   End Select

Exit_CmdAfficherTous_Click:

   Exit Sub

Err_CmdAfficherTous_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherTous_Click"

   Resume Exit_CmdAfficherTous_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdRechercher_Click()

   On Error GoTo Exit_CmdRechercher_Click

   OuvreRecherches 3

Exit_CmdRechercher_Click:

   Exit Sub

Err_CmdRechercher_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdRechercher_Click"

   Resume Exit_CmdRechercher_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdRechercherMulti_Click()

   On Error GoTo Err_CmdRechercherMulti_Click

   OuvreRecherchesMulti

Exit_CmdRechercherMulti_Click:

   Exit Sub

Err_CmdRechercherMulti_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdRechercherMulti_Click"

   Resume Exit_CmdRechercherMulti_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchCadArtNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstCadAchCadArtNum_NotInList

   AjouteCadeauxArticles (NewData)

   Response = acDataErrContinue

   TxtCadAchCadArtNum = LstCadAchCadArtNum.Text

   LstCadAchCadArtNum.Undo

Exit_LstCadAchCadArtNum_NotInList:

   Exit Sub

Err_LstCadAchCadArtNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchCadArtNum_NotInList"

   Resume Exit_LstCadAchCadArtNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtCadAchQuantite_AfterUpdate()

   On Error GoTo Err_TxtCadAchQuantite_AfterUpdate

   TxtCadAchPoints = LstCadAchCadArtNum.Column(4) * TxtCadAchQuantite

Exit_TxtCadAchQuantite_AfterUpdate:

   Exit Sub

Err_TxtCadAchQuantite_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtCadAchQuantite_AfterUpdate"

   Resume Exit_TxtCadAchQuantite_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchCadArtNum_AfterUpdate()

   On Error GoTo Err_LstCadAchCadArtNum_AfterUpdate

   TxtCadAchCadArtNum = vbNullString

   TxtCadAchPoints = LstCadAchCadArtNum.Column(4) * TxtCadAchQuantite

Exit_LstCadAchCadArtNum_AfterUpdate:

   Exit Sub

Err_LstCadAchCadArtNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchCadArtNum_AfterUpdate"

   Resume Exit_LstCadAchCadArtNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchSocNum_GotFocus()

   On Error GoTo Err_LstCadAchSocNum_GotFocus

 '  LstCadAchSocNum.Dropdown

Exit_LstCadAchSocNum_GotFocus:

   Exit Sub

Err_LstCadAchSocNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchSocNum_GotFocus"

   Resume Exit_LstCadAchSocNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchCadArtNum_GotFocus()

   On Error GoTo Err_LstCadAchCadArtNum_GotFocus

 '  LstCadAchCadArtNum.Dropdown

Exit_LstCadAchCadArtNum_GotFocus:

   Exit Sub

Err_LstCadAchCadArtNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchCadArtNum_GotFocus"

   Resume Exit_LstCadAchCadArtNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchEtat_GotFocus()

   On Error GoTo Err_LstCadAchEtat_GotFocus

'  LstCadAchEtat.Dropdown

Exit_LstCadAchEtat_GotFocus:

   Exit Sub

Err_LstCadAchEtat_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchEtat_GotFocus"

   Resume Exit_LstCadAchEtat_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadAchCadArtNum_LostFocus()

   On Error GoTo Err_LstCadAchCadArtNum_LostFocus

   TxtCadAchCadArtNum = vbNullString

Exit_LstCadAchCadArtNum_LostFocus:

   Exit Sub

Err_LstCadAchCadArtNum_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadAchCadArtNum_LostFocus"

   Resume Exit_LstCadAchCadArtNum_LostFocus
End Sub
