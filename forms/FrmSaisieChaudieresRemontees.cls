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
'***   Declaration De Variable Public                                                           ***
'******************************************************************************

'******************************************************************************
'***   Variable De Controle De Mode De Saisie                                             ***
'******************************************************************************
Public IntMode As Integer
Public IntModeOld As Integer

'******************************************************************************
'***   Declaration De Variable Public                                                           ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Le Mode Rafraichissement                                           ***
'******************************************************************************
Private TxtChdRemChdArtRef As String
Private TxtChdRemSocNum As String

'******************************************************************************
'***    Declaration De Procedure Priver                                                        ***
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
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Set CtlActive = TxtChdRemSerie

               End Select

               On Error GoTo Err_Rafraichir

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               Select Case TxtChdRemSocNum
                  Case vbNullString

                     LstChdRemSocNum.Requery

                  Case Else

                     LstChdRemSocNum.Requery

                     LstChdRemSocNum.Text = TxtChdRemSocNum

                     TxtChdRemSocNum = vbNullString

               End Select

               Select Case TxtChdRemChdArtRef
                  Case vbNullString

                     LstChdRemChdArtRef.Requery

                  Case Else

                     LstChdRemChdArtRef.Requery

                     LstChdRemChdArtRef.Text = TxtChdRemChdArtRef

                     TxtChdRemChdArtRef = vbNullString

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
   Dim VarChdRemNum As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(ChdRemSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Code STA Est Invalide !", , Me.Caption

         LstChdRemSocNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(ChdRemDate)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Date De Saisie Est Invalide !", , Me.Caption

         TxtChdRemDateSaisie.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(ChdRemHeure)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Heure De Saisie Est Invalide !", , Me.Caption

         TxtChdRemHeureSaisie.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

  Select Case IsDate(ChdRemSD)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Date D'Arrivée Saunier Duval Est Invalide !", , Me.Caption

         TxtChdRemSD.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(ChdRemServ)
      Case True

         Select Case ChdRemServ < ChdRemSD
            Case True

            Case False

               Cancel = True

               MsgBox "Le Champ Date Mise En Service Est Incorrecte !", , Me.Caption

               TxtChdRemSD.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select
      Case False

         Cancel = True

         MsgBox "Le Champ Date De Mise En Service Est Invalide !", , Me.Caption

         TxtChdRemServ.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(ChdRemChdArtRef)
      Case True

         Cancel = True

         MsgBox "Le Champ Reférence Est Obligatoire !", , Me.Caption

         LstChdRemChdArtRef.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

   End Select

   Select Case IsNull(ChdRemSerie)
      Case True

         Cancel = True

         MsgBox "Le Champ Numero De Série Est Obligatoire !", , Me.Caption

         TxtChdRemSerie.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case False

         Select Case Not (Len(TxtChdRemSerie) = 8 Or Len(TxtChdRemSerie) = 10)
            Case True

               Cancel = True

               MsgBox "Le Numero De Série Doit Etre Sur 8 Ou 10 Caractere !", , Me.Caption

               TxtChdRemSerie.SetFocus

               GoTo Exit_Form_BeforeUpdate

            Case False

                VarChdRemNum = DLookup("ChdRemNum", "TBLCHAUDIERESREMONTEES", "ChdRemSerie='" & ChdRemSerie & "' AND ChdRemNum<>" & ChdRemNum)

                Select Case IsNull(VarChdRemNum)
                   Case True

                   Case False

                      Cancel = True

                      MsgBox "Le Numero De Série Est Deja Utiliser Pour La Remontée N° " & VarChdRemNum & " !", , Me.Caption

                      TxtChdRemSerie.SetFocus

                      GoTo Exit_Form_BeforeUpdate

                End Select

                VarChdRemNum = DLookup("RCNum", "CDP2002CHAUDIERESREMONTEES", "RCSerie='" & ChdRemSerie & "'  AND Year(RCDate)<>2003")

                Select Case IsNull(VarChdRemNum)
                   Case True

                   Case False

                      Cancel = True

                      MsgBox "Le Numero De Série Est Deja Utiliser Pour La Remontée N° " & VarChdRemNum & " Dans Coeur De Pro 2002!", , Me.Caption

                      TxtChdRemSerie.SetFocus

                      GoTo Exit_Form_BeforeUpdate

                End Select

         End Select
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

   TxtChdRemSerie.SetFocus

   Select Case IntModeOld
      Case MODEAJOUT

         AjouteMouvementsChaudieres , ChdRemNum

      Case MODEMODIF

         SupprimeMouvementsChaudieres , ChdRemNum

         AjouteMouvementsChaudieres ChdRemSocNum

   End Select

   ChangeBoutonMode Me, MODEDEPLACEMENT

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
Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

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
Private Sub Form_Current()

   On Error GoTo Err_Form_Current

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
Private Sub CmdAjouter_Click()
   Dim TxtDateSaisie As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus
      Case Else

         LstChdRemSocNum.SetFocus

         LstChdRemSocNum.DefaultValue = LstChdRemSocNum

         TxtChdRemSD.DefaultValue = "'" & TxtChdRemSD & "'"

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateSaisie = TxtChdRemDateSaisie

   TxtChdRemDateSaisie = "01/01/2003"

   TxtChdRemDateSaisie = TxtDateSaisie

   TxtChdRemSerie.SetFocus

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

         LstChdRemSocNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateSaisie = TxtChdRemDateSaisie

         TxtChdRemDateSaisie = "01/01/2003"

         TxtChdRemDateSaisie = TxtDateSaisie

         TxtChdRemSerie.SetFocus

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
   Dim LngChdRemSocNum As Long
   Dim LngChdRemNum As Long

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstChdRemSocNum.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         LngChdRemSocNum = ChdRemSocNum

         LngChdRemNum = ChdRemNum

         SupprimeMouvementsChaudieres , LngChdRemNum

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

         AjouteMouvementsChaudieres LngChdRemSocNum

         ChangeBoutonMode Me, MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

            Case Else

               TxtChdRemSerie.SetFocus

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

   LstChdRemSocNum.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

   TxtChdRemSerie.SetFocus

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

   LstChdRemSocNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtChdRemSerie.SetFocus

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
Private Sub CmdAfficherSocietes_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   Select Case IsNull(ChdRemSocNum)
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & ChdRemSocNum

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

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & ChdRemSocNum, acFormEdit

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
Private Sub CmdExporteExcel_Click()
   Dim StrSQLRequete As String
   Dim RsChaudieresRemontees As DAO.Recordset

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

         Set RsChaudieresRemontees = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case Err.Number
            Case 0

               RsChaudieresRemontees.Close

               ExporteExcelRequete (StrSQLRequete)

            Case Else

               Err.Clear

               MsgBox "Impossible D'Exporter " & IIf(Me.FilterOn, "Avec Le Filtre En Cours", vbNullString) & " !!!"

         End Select

   End Select

Exit_CmdExporteExcel:

   Set RsChaudieresRemontees = Nothing

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
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterChaudieres_Click()

   On Error GoTo Err_CmdAjouterChaudieres_Click

   AjouteChaudieresArticles

Exit_CmdAjouterChaudieres_Click:

   Exit Sub

Err_CmdAjouterChaudieres_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterChaudieres_Click"

   Resume Exit_CmdAjouterChaudieres_Click
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

   OuvreRecherches 4

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
Private Sub LstChdRemChdArtRef_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstChdRemChdArtRef_NotInList

   AjouteChaudieresArticles (NewData)

   Response = acDataErrContinue

   TxtChdRemChdArtRef = LstChdRemChdArtRef.Text

   LstChdRemChdArtRef.Undo

Exit_LstChdRemChdArtRef_NotInList:

   Exit Sub

Err_LstChdRemChdArtRef_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChdRemChdArtRef_NotInList"

   Resume Exit_LstChdRemChdArtRef_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtChdRemServ_AfterUpdate()
   Dim BytSemNum As Byte

   On Error GoTo Err_TxtChdRemServ_AfterUpdate

   BytSemNum = DLookup("[SemNum]", "[TBLSEMESTRES]", "CDate('" & TxtChdRemServ & "') Between [SemDebut] AND [SemFin]")

   TxtChdRemSemNum = BytSemNum

Exit_TxtChdRemServ_AfterUpdate:

   Exit Sub

Err_TxtChdRemServ_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtChdRemServ_AfterUpdate"

   Resume Exit_TxtChdRemServ_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtChdRemSerie_AfterUpdate()
   Dim RsSocietesChaudieres As DAO.Recordset
   Dim VarChdRemNum As Variant
   Dim StrSocietesChaudieres As String

   On Error GoTo Err_TxtChdRemSerie_AfterUpdate

   Select Case IntMode
      Case MODEMODIF, MODEAJOUT

         Select Case Not (Len(TxtChdRemSerie) = 8 Or Len(TxtChdRemSerie) = 10)
            Case True

               MsgBox "Le Numero De Série Doit Etre Sur 8 Ou 10 Caractere !", , Me.Caption

               GoTo Exit_TxtChdRemSerie_AfterUpdate

            Case False

                StrSocietesChaudieres = "SELECT TBLCHAUDIERESREMONTEES.*, TBLSOCIETES.* FROM TBLSOCIETES INNER JOIN " & _
                     " TBLCHAUDIERESREMONTEES ON TBLSOCIETES.SocNum = TBLCHAUDIERESREMONTEES.ChdRemSocNum " & _
                     " WHERE ChdRemSerie='" & ChdRemSerie & "' AND ChdRemNum<>" & ChdRemNum & ";"

                Set RsSocietesChaudieres = CurrentDb.OpenRecordset(StrSocietesChaudieres)

                Select Case RsSocietesChaudieres.EOF
                   Case True

                   Case False

                     Select Case MsgBox("Le Numero De Série Est Deja Utiliser Pour La Remontée N° " & RsSocietesChaudieres!ChdRemNum & " !" & vbCrLf & _
                                                     "Code STA : " & RsSocietesChaudieres!SocSTA & vbCrLf & _
                                                     "Societes : " & RsSocietesChaudieres!SocRaison & vbCrLf & _
                                                     "Remplacement : " & RsSocietesChaudieres!ChdRemRemp & vbCrLf & _
                                                     "Voulez Vous Allez Sur La Fiche Remontees De Chaudieres Correspondante ?", vbYesNo, Me.Caption)

                        Case vbYes

                           CmdAnnuler_Click

                           DoCmd.ApplyFilter , "ChdRemNum=" & RsSocietesChaudieres!ChdRemNum

                        Case vbNo

                     End Select

                      GoTo Exit_TxtChdRemSerie_AfterUpdate

                End Select

                VarChdRemNum = DLookup("RCNum", "CDP2002CHAUDIERESREMONTEES", "RCSerie='" & ChdRemSerie & "' AND Year(RCDate)<>2003")

                Select Case IsNull(VarChdRemNum)
                   Case True

                   Case False

                      MsgBox "Le Numero De Série Est Deja Utiliser Pour La Remontée N° " & VarChdRemNum & " Dans Coeur De Pro 2002!", , Me.Caption

                      GoTo Exit_TxtChdRemSerie_AfterUpdate

                End Select

         End Select

         Case Else

   End Select

Exit_TxtChdRemSerie_AfterUpdate:

   Exit Sub

Err_TxtChdRemSerie_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtChdRemSerie_AfterUpdate"

   Resume Exit_TxtChdRemSerie_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstChdRemChdArtRef_LostFocus()

   On Error GoTo Err_LstChdRemChdArtRef_LostFocus

   TxtChdRemChdArtRef = vbNullString

Exit_LstChdRemChdArtRef_LostFocus:

   Exit Sub

Err_LstChdRemChdArtRef_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChdRemChdArtRef_LostFocus"

   Resume Exit_LstChdRemChdArtRef_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstChdRemSocNum_GotFocus()

   On Error GoTo Err_LstChdRemSocNum_GotFocus

'  LstChdRemSocNum.Dropdown

Exit_LstChdRemSocNum_GotFocus:

   Exit Sub

Err_LstChdRemSocNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChdRemSocNum_GotFocus"

   Resume Exit_LstChdRemSocNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstChdRemChdArtRef_GotFocus()

   On Error GoTo Err_LstChdRemChdArtRef_GotFocus

'   LstChdRemChdArtRef.Dropdown

Exit_LstChdRemChdArtRef_GotFocus:

   Exit Sub

Err_LstChdRemChdArtRef_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChdRemChdArtRef_GotFocus"

   Resume Exit_LstChdRemChdArtRef_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstChdRemChdArtRef_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_LstChdRemChdArtRef_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_LstChdRemChdArtRef_KeyPress:

    Exit Sub

Err_LstChdRemChdArtRef_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChdRemChdArtRef_KeyPress"

    Resume Exit_LstChdRemChdArtRef_KeyPress
End Sub
