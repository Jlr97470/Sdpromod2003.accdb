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

'******************************************************************************
'***    Declaration De Variable Privée                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Du Numero De Questionnaires En Cours                               ***
'******************************************************************************

'******************************************************************************
'***    Declaration De Procedure Privée                                                       ***
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

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = CadSocCallQst1

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               Select Case CtlActive.Name
                  Case "CmdAfficherCall1", "CmdAfficherCall2"

                     CmdImprimer.SetFocus

                     Set CtlActive = CmdImprimer

                  Case Else

               End Select

               CmdAfficherCall1.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL1", "SocCallSocNum=" & SocNum)), True, False)

               CmdAfficherCall2.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL2", "SocCallSocNum=" & SocNum)), True, False)

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

               Select Case IsNull(SocCallNum)
                  Case True

                  Case False

                     TxtSocCallQst3.Enabled = (CadSocCallQst2 = 1)

                     TxtSocCallQst3.Locked = Not (CadSocCallQst2 = 1)

                     TxtSocCallQst4.Enabled = (CadSocCallQst2 = 1)

                     TxtSocCallQst4.Locked = Not (CadSocCallQst2 = 1)

                     LstSocCallQst5.Enabled = (CadSocCallQst2 = 1)

                     LstSocCallQst5.Locked = Not (CadSocCallQst2 = 1)

                     TxtSocCallQst6.Enabled = (CadSocCallQst2 = 1)

                     TxtSocCallQst6.Locked = Not (CadSocCallQst2 = 1)

                     TxtSocCallQst7.Enabled = (CadSocCallQst2 = 1)

                     TxtSocCallQst7.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst8.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst8.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst9.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst9.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst10.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst10.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst11.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst11.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst12.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst12.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst13.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst13.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst14.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst14.Locked = Not (CadSocCallQst2 = 1)

                     TxtSocCallQst15.Enabled = (CadSocCallQst2 = 1)

                     TxtSocCallQst15.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst16.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst16.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst17.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst17.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst18.Enabled = (CadSocCallQst2 = 2)

                     CadSocCallQst18.Locked = Not (CadSocCallQst2 = 2)

                     TxtSocCallQst19.Enabled = (CadSocCallQst2 = 2)

                     TxtSocCallQst19.Locked = Not (CadSocCallQst2 = 2)

                     LstSocCallQst20.Enabled = (CadSocCallQst2 = 2)

                     LstSocCallQst20.Locked = Not (CadSocCallQst2 = 2)

                     TxtSocCallQst21.Enabled = (CadSocCallQst2 = 2)

                     TxtSocCallQst21.Locked = Not (CadSocCallQst2 = 2)

                     Select Case CadSocCallQst2
                        Case 0

                        Case 1

                           CadSocCallQst18 = 0

                           TxtSocCallQst19 = Null

                           LstSocCallQst20 = 0

                           TxtSocCallQst21 = Null

                        Case 2

                           TxtSocCallQst3 = Null

                           TxtSocCallQst4 = Null

                           LstSocCallQst5 = 0

                           TxtSocCallQst6 = Null

                           TxtSocCallQst7 = Null

                           CadSocCallQst8 = 0

                           CadSocCallQst9 = 0

                           CadSocCallQst10 = 0

                           CadSocCallQst11 = 0

                           CadSocCallQst12 = 0

                           CadSocCallQst13 = 0

                           CadSocCallQst14 = 0

                           TxtSocCallQst15 = Null

                           CadSocCallQst16 = 0

                           CadSocCallQst17 = 0

                     End Select

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
'***    Declaration D'Evenement Privée                                                        ***
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

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         Me.MenuBar = "Menu Coeur De Pro 2003"

         Me.Toolbar = "Outil Coeur De Pro 2003"

         CmdAfficherContacts.Visible = True

         CmdAfficherSocietes.Visible = True

      Case Else

         Me.MenuBar = "Menu Call Coeur De Pro 2003"

         Me.Toolbar = "Outil Call Coeur De Pro 2003"

         CmdAfficherContacts.Visible = False

         CmdAfficherSocietes.Visible = False

   End Select

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
Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case CadSocCallQst1
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst1").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst1.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst2
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst2").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst2.SetFocus

         GoTo Exit_Form_BeforeUpdate

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
Private Sub Form_Current()

   On Error GoTo Err_Form_Current

   Rafraichir (RAFRAICHIRDEPLACEMENT)

   SFrmSaisieCallContacts.Form.AllowEdits = True

   SFrmSaisieCallSocietes.Form.AllowEdits = True

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Select Case Err.Number
            Case 2501, 2759

            Case Else

               MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdFermer_Click"

         End Select

         Err.Clear

         GoTo Exit_CmdFermer_Click

   End Select

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
Private Sub CmdRechercher_Click()

   On Error GoTo Err_CmdRechercher_Click

   OuvreRecherches 2

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
Private Sub CmdAfficherContacts_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherContacts_Click

   Select Case IsNull(DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum))
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheContacts", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheContacts.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "ConNum=" & DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum)

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_CmdAfficherContacts_Click

               End Select

               On Error GoTo Err_CmdAfficherContacts_Click

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum), acFormEdit

         End Select

   End Select

Exit_CmdAfficherContacts_Click:

   Exit Sub

Err_CmdAfficherContacts_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherContacts_Click"

   Resume Exit_CmdAfficherContacts_Click
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

   Select Case IsNull(SocNum)
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & SocNum

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

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & SocNum, acFormEdit

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
Private Sub CmdMinimize_Click()

   On Error GoTo Err_CmdMinimize_Click

   DoCmd.Minimize

Exit_CmdMinimize_Click:

   Exit Sub

Err_CmdMinimize_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdMinimize_Click"

   Resume Exit_CmdMinimize_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherCall1_Click()

   On Error GoTo Err_CmdAfficherCall1_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall1", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall1", acNormal, , "SocNum=" & SocNum, , acWindowNormal

   End Select

Exit_CmdAfficherCall1_Click:

   Exit Sub

Err_CmdAfficherCall1_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCall1_Click"

   Resume Exit_CmdAfficherCall1_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherCall2_Click()

   On Error GoTo Err_CmdAfficherCall2_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall2", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall2", acNormal, , "SocNum=" & SocNum, , acWindowNormal

   End Select

Exit_CmdAfficherCall2_Click:

   Exit Sub

Err_CmdAfficherCall2_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCall2_Click"

   Resume Exit_CmdAfficherCall2_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst2_AfterUpdate()

   On Error GoTo Err_CadSocCallQst2_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst2
      Case 0

      Case 1

         TxtSocCallQst3.SetFocus

      Case 2

         CadSocCallQst18.SetFocus

   End Select

Exit_CadSocCallQst2_AfterUpdate:

   Exit Sub

Err_CadSocCallQst2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst2_AfterUpdate"

   Resume Exit_CadSocCallQst2_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocCallQst5_AfterUpdate()

   On Error GoTo Err_LstSocCallQst5_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case LstSocCallQst5
      Case 0, 1, 3, 4

            EtiSocCallQst6.Visible = False

            TxtSocCallQst6.Visible = False

            EtiSocCallQst7.Visible = False

            TxtSocCallQst7.Visible = False

      Case 2

            EtiSocCallQst6.Visible = True

            TxtSocCallQst6.Visible = True

            EtiSocCallQst7.Visible = False

            TxtSocCallQst7.Visible = False

      Case 5

            EtiSocCallQst6.Visible = False

            TxtSocCallQst6.Visible = False

            EtiSocCallQst7.Visible = True

            TxtSocCallQst7.Visible = True

   End Select

Exit_LstSocCallQst5_AfterUpdate:

   Exit Sub

Err_LstSocCallQst5_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocCallQst5_AfterUpdate"

   Resume Exit_LstSocCallQst5_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocCallQst20_AfterUpdate()

   On Error GoTo Err_LstSocCallQst20_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case LstSocCallQst20
      Case 0, 1, 2, 3

            EtiSocCallQst21.Visible = False

            TxtSocCallQst21.Visible = False

      Case 4

            EtiSocCallQst21.Visible = True

            TxtSocCallQst21.Visible = True

   End Select

Exit_LstSocCallQst20_AfterUpdate:

   Exit Sub

Err_LstSocCallQst20_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocCallQst20_AfterUpdate"

   Resume Exit_LstSocCallQst20_AfterUpdate
End Sub
