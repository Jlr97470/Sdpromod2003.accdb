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
'***    Variable De Controle De Mode Saisie                                                 ***
'******************************************************************************
Public IntMode As Integer
Public IntModeOld As Integer
Public BlnModif As Boolean

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
                  Case vbEmpty

                  Case Else

                      Err.Clear

                      Set CtlActive = TxtSocQstTypID

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               CmbSocQstTypIDTypNum.Requery

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRDEPLACEMENT

               Select Case IsNull(SocQstTypIDTypNum)
                  Case False

                     Select Case SocQstTypIDTypNum
                        Case vbBoolean

                           TxtSocQstTypIDMin.Enabled = False

                           TxtSocQstTypIDMax.Enabled = False

                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                           TxtSocQstTypIDMin.Enabled = True

                           TxtSocQstTypIDMax.Enabled = True

                     End Select

                  Case True

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

   On Error Resume Next

   Select Case IsObject(Me.Parent)
      Case True

      Case False

         Cancel = True

   End Select

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

         Cancel = True

   End Select

   On Error GoTo Err_Form_Open

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
Private Sub Form_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocQstTypID)
      Case False

         TxtSocQstTypID = UCase(TxtSocQstTypID)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Est Obligatoire", , Me.Caption

         TxtSocQstTypID.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocQstTypIDNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire", , Me.Caption

         CmbSocQstTypIDTypNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocQstTypIDLibelle)
      Case False

         TxtSocQstTypIDLibelle = UCase(TxtSocQstTypIDLibelle)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Libelle Est Obligatoire", , Me.Caption

         TxtSocQstTypIDLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocQstTypIDMin)
      Case True

      Case False

         Select Case SocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

               On Error Resume Next

               Select Case SocQstTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(SocQstTypIDMin)

                  Case vbInteger

                     VarValeur = CInt(SocQstTypIDMin)

                  Case vbLong

                     VarValeur = CLng(SocQstTypIDMin)

                  Case vbSingle

                     VarValeur = CSng(SocQstTypIDMin)

                  Case vbDouble

                     VarValeur = CDbl(SocQstTypIDMin)

                  Case vbCurrency

                     VarValeur = CCur(SocQstTypIDMin)

                  Case vbDate

                     VarValeur = CDate(SocQstTypIDMin)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Questionnaires Types ID Min Est Invalide", , Me.Caption

                     TxtSocQstTypIDMin.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               SocQstTypIDMin = Null

         End Select

   End Select

   Select Case IsNull(SocQstTypIDMax)
      Case True

      Case False

         Select Case SocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               On Error Resume Next

               Select Case SocQstTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(SocQstTypIDMax)

                  Case vbInteger

                     VarValeur = CInt(SocQstTypIDMax)

                  Case vbLong

                     VarValeur = CLng(SocQstTypIDMax)

                  Case vbSingle

                     VarValeur = CSng(SocQstTypIDMax)

                  Case vbDouble

                     VarValeur = CDbl(SocQstTypIDMax)

                  Case vbCurrency

                     VarValeur = CCur(SocQstTypIDMax)

                  Case vbDate

                     VarValeur = CDate(SocQstTypIDMax)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Questionnaires Types ID Max Est Invalide", , Me.Caption

                     TxtSocQstTypIDMax.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               SocQstTypIDMax = Null

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

   CreeAnalyseSocietesQuestionnaires SocQstTypIDSocQstTypNum

   TxtSocQstTypID.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtSocQstTypID.SetFocus

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
   Dim BolIDObligatoire As Boolean

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.Parent!SocQstTypNbResultat
      Case Is > Me.RecordsetClone.RecordCount

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

               CmdImprimer.SetFocus

            Case Else

               TxtSocQstTypID.SetFocus

         End Select

         ChangeBoutonMode Me, MODEAJOUT

         DoCmd.RunCommand acCmdRecordsGoToNew

         BolIDObligatoire = CbxSocQstTypIDObligatoire

         CbxSocQstTypIDObligatoire = True

         CbxSocQstTypIDObligatoire = BolIDObligatoire

         TxtSocQstTypID.SetFocus

      Case Else

         MsgBox "Attention Le Nombre De Résultat Attendu Est Atteint !!!", vbCritical, Me.Caption

   End Select

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
   Dim BolIDObligatoire As Boolean

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtSocQstTypID.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         BolIDObligatoire = CbxSocQstTypIDObligatoire

         CbxSocQstTypIDObligatoire = True

         CbxSocQstTypIDObligatoire = BolIDObligatoire

         TxtSocQstTypID.SetFocus

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

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtSocQstTypID.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

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

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                TxtSocQstTypID.SetFocus

          End Select

         On Error GoTo Err_CmdSupprimer_Click

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

   TxtSocQstTypID.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

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

   TxtSocQstTypID.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtSocQstTypID.SetFocus

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
Private Sub CmbSocQstTypIDTypNum_AfterUpdate()

   On Error GoTo Err_CmbSocQstTypIDTypNum_AfterUpdate

   Select Case IsNull(SocQstTypIDTypNum)
      Case False

         Select Case SocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

               TxtSocQstTypIDMin.Enabled = True

               TxtSocQstTypIDMax.Enabled = True

            Case vbBoolean

               TxtSocQstTypIDMin.Enabled = False

               TxtSocQstTypIDMax.Enabled = False

               TxtSocQstTypIDMin = Null

               TxtSocQstTypIDMax = Null

         End Select

      Case True

         MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

         CmbSocQstTypIDTypNum.SetFocus

   End Select

Exit_CmbSocQstTypIDTypNum_AfterUpdate:

    Exit Sub

Err_CmbSocQstTypIDTypNum_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstTypIDTypNum_AfterUpdate"

    Resume Exit_CmbSocQstTypIDTypNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypIDMin_Exit(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_TxtSocQstTypIDMin_Exit

   Select Case IsNull(SocQstTypIDMin)
      Case True

      Case False

         Select Case IsNull(SocQstTypIDTypNum)
            Case False

               Select Case SocQstTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case SocQstTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(SocQstTypIDMin)

                        Case vbInteger

                           VarValeur = CInt(SocQstTypIDMin)

                        Case vbLong

                           VarValeur = CLng(SocQstTypIDMin)

                        Case vbSingle

                           VarValeur = CSng(SocQstTypIDMin)

                        Case vbDouble

                           VarValeur = CDbl(SocQstTypIDMin)

                        Case vbCurrency

                           VarValeur = CCur(SocQstTypIDMin)

                        Case vbDate

                           VarValeur = CDate(SocQstTypIDMin)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Questionnaires Types ID Min Est Invalide !", , Me.Caption

                           Cancel = True

                     End Select

                     On Error GoTo Err_TxtSocQstTypIDMin_Exit

                  Case vbBoolean

                     SocQstTypIDMin = Null

               End Select

            Case True

               MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

               CmbSocQstTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtSocQstTypIDMin_Exit:

    Exit Sub

Err_TxtSocQstTypIDMin_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypIDMin_Exit"

    Resume Exit_TxtSocQstTypIDMin_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypIDMax_Exit(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_TxtSocQstTypIDMax_Exit

   Select Case IsNull(SocQstTypIDMax)
      Case True

      Case False

         Select Case IsNull(SocQstTypIDTypNum)
            Case False

               Select Case SocQstTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case SocQstTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(SocQstTypIDMax)

                        Case vbInteger

                           VarValeur = CInt(SocQstTypIDMax)

                        Case vbLong

                           VarValeur = CLng(SocQstTypIDMax)

                        Case vbSingle

                           VarValeur = CSng(SocQstTypIDMax)

                        Case vbDouble

                           VarValeur = CDbl(SocQstTypIDMax)

                        Case vbCurrency

                           VarValeur = CCur(SocQstTypIDMax)

                        Case vbDate

                           VarValeur = CDate(SocQstTypIDMax)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Questionnaires Types ID Max Est Invalide !", , Me.Caption

                           TxtSocQstTypIDMax.SetFocus

                     End Select

                     On Error GoTo Err_TxtSocQstTypIDMax_Exit

                  Case vbBoolean

                     SocQstTypIDMax = Null

               End Select

            Case True

               MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

               CmbSocQstTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtSocQstTypIDMax_Exit:

    Exit Sub

Err_TxtSocQstTypIDMax_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypIDMax_Exit"

    Resume Exit_TxtSocQstTypIDMax_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypID_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtSocQstTypID_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtSocQstTypID_KeyPress:

    Exit Sub

Err_TxtSocQstTypID_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypID_KeyPress"

    Resume Exit_TxtSocQstTypID_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypIDLibelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtSocQstTypIDLibelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtSocQstTypIDLibelle_KeyPress:

    Exit Sub

Err_TxtSocQstTypIDLibelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypIDLibelle_KeyPress"

    Resume Exit_TxtSocQstTypIDLibelle_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypIDMin_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtSocQstTypIDMin_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtSocQstTypIDMin_KeyPress:

    Exit Sub

Err_TxtSocQstTypIDMin_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypIDMin_KeyPress"

    Resume Exit_TxtSocQstTypIDMin_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocQstTypIDMax_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_TxtSocQstTypIDMax_KeyPress

   Select Case SocQstTypIDTypNum
      Case vbString

      Case Else

         KeyAscii = Asc(UCase((Chr$(KeyAscii))))

   End Select

Exit_TxtSocQstTypIDMax_KeyPress:

    Exit Sub

Err_TxtSocQstTypIDMax_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocQstTypIDMax_KeyPress"

    Resume Exit_TxtSocQstTypIDMax_KeyPress
End Sub
