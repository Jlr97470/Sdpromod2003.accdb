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
                  Case "CmdAfficherCall2", "CmdAfficherCall3"

                     CmdImprimer.SetFocus

                     Set CtlActive = CmdImprimer

                  Case Else

               End Select

               CmdAfficherCall2.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL2", "SocCallSocNum=" & SocNum)), True, False)

               CmdAfficherCall3.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL3", "SocCallSocNum=" & SocNum)), True, False)

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

   Select Case CadSocCallQst3
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst3").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst3.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst5
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst5").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst5.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst7
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst7").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst7.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst9
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst9").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst9.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst11
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst11").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst11.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst13
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst13").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst13.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst15
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst15").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst15.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst17
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst17").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst17.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst19
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst19").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst19.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst21
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst21").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst21.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst23
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst23").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst23.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst25
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst25").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst25.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst27
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst27").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst27.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst29
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst29").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst29.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst31
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst31").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst31.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst33
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst33").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst33.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst35
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst35").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst35.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst37
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst37").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst37.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst39
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst39").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst39.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst41
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst41").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst41.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst43
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst43").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst43.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst45
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst45").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst45.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst49
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst49").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst49.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst50
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst50").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst50.SetFocus

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
Private Sub CmdAfficherCall3_Click()

   On Error GoTo Err_CmdAfficherCall3_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall3", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall3", acNormal, , "SocNum=" & SocNum, , acWindowNormal

   End Select

Exit_CmdAfficherCall3_Click:

   Exit Sub

Err_CmdAfficherCall3_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCall3_Click"

   Resume Exit_CmdAfficherCall3_Click
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
