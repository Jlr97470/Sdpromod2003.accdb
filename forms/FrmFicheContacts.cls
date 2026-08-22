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
'***    Variable De Controle Du Mode De Saisie                                            ***
'******************************************************************************
Public IntMode As Integer
Public IntModeOld As Integer
Public BlnModif As Boolean
Public BlnChangeAdresse As Boolean
Public BlnChangeTelephone As Boolean

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Varaible Du Mode De Rafraichissement                                              ***
'******************************************************************************
Private TxtConSocNum As String
Private TxtConCivCode As String
Private TxtConFonCode As String
Private TxtConVilNum As String
Private TxtConPerVilNum As String

'******************************************************************************
'***    PROCEDURE :                                                                                  ***
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
   Dim StrSQLSocietesDebut As String
   Dim StrSQLSocietesFin As String
   Dim StrSQLFonctionsDebut As String
   Dim StrSQLFonctionsFin As String
   Dim StrSQLCiv As String
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   StrSQLSocietesDebut = "SELECT TBLSOCIETES.SocNum, TBLSOCIETES.SocRaison,TBLSOCIETES.SocSTA, " & _
                                       "TBLSOCIETES.SocAdh, TBLSOCIETES.SocTel1, TBLSOCIETES.SocTel2, TBLSOCIETES.SocTypCode FROM TBLSOCIETES "

   StrSQLSocietesFin = " ORDER BY TBLSOCIETES.SocRaison; "

   StrSQLFonctionsDebut = "SELECT TBLCONTACTSFONCTIONS.ConFonCode, TBLCONTACTSFONCTIONS.ConFonLibelle FROM TBLCONTACTSFONCTIONS "

   StrSQLFonctionsFin = "ORDER BY TBLCONTACTSFONCTIONS.ConFonCode;"

   StrSQLCiv = "SELECT TBLCIVILITES.CivCode, TBLCIVILITES.CivLibelle FROM TBLCIVILITES ORDER BY TBLCIVILITES.CivCode;"

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

                      Set CtlActive = Me.TxtConNum

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT, RAFRAICHIRCHOIX

               Select Case LstConFonCode.Column(0)
                  Case "SDDR", "ATC"

                     StrSQLSocietesDebut = StrSQLSocietesDebut & " WHERE TBLSOCIETES.SocTypCode='DR' "

                  Case "TA"

                     StrSQLSocietesDebut = StrSQLSocietesDebut & " WHERE TBLSOCIETES.SocTypCode='STA' "

                  Case Else

               End Select

               Select Case LstConSocNum.Column(6)
                  Case "SD"

                     TxtConSocSTA.Visible = False

                     TxtConSocAdh.Visible = False

                  Case "DR"

                     StrSQLFonctionsDebut = StrSQLFonctionsDebut & " WHERE TBLCONTACTSFONCTIONS.ConFonCode<>'TA' "

                     TxtConSocSTA.Visible = False

                     TxtConSocAdh.Visible = False

                  Case "STA"

                     StrSQLFonctionsDebut = StrSQLFonctionsDebut & " WHERE TBLCONTACTSFONCTIONS.ConFonCode='TA' "

                     TxtConSocSTA.Visible = True

                     TxtConSocAdh.Visible = True

                  Case Else

                     TxtConSocSTA.Visible = False

                     TxtConSocAdh.Visible = False

               End Select

               LstConSocNum.RowSource = StrSQLSocietesDebut & StrSQLSocietesFin

               LstConFonCode.RowSource = StrSQLFonctionsDebut & StrSQLFonctionsFin

               Select Case TxtConSocNum
                  Case vbNullString

                     LstConSocNum.Requery

                  Case Else

                     LstConSocNum.Requery

                     LstConSocNum.Text = TxtConSocNum

                     TxtConSocNum = vbNullString

               End Select

               Select Case TxtConFonCode
                  Case vbNullString

                     LstConFonCode.Requery

                  Case Else

                     LstConFonCode.Requery

                     LstConFonCode.Text = TxtConFonCode

                     TxtConFonCode = vbNullString

               End Select
            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               LstConCivCode.RowSource = StrSQLCiv

               Select Case TxtConCivCode
                  Case vbNullString

                     LstConCivCode.Requery

                  Case Else

                     LstConCivCode.Requery

                     LstConCivCode.Text = TxtConCivCode

                     TxtConCivCode = vbNullString

               End Select

            Case Else

         End Select

         Select Case OgtFicheContactsDetails
            Case 0

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

                     Select Case TxtConVilNum
                        Case vbNullString

                           LstConVilNum.Requery

                        Case True

                           LstConVilNum.Requery

                           LstConVilNum.Text = TxtConVilNum

                           TxtConVilNum = vbNullString

                     End Select
                  Case Else

               End Select
            Case 1

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

                     Select Case TxtConPerVilNum
                        Case vbNullString

                           LstConPerVilNum.Requery

                        Case True

                           LstConPerVilNum.Requery

                           LstConPerVilNum.Text = TxtConPerVilNum

                           TxtConPerVilNum = vbNullString

                     End Select
                  Case Else

               End Select

            Case 2

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                  Case Else

               End Select
            Case 3

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstOrganigramme.Requery

                  Case Else

               End Select
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

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         Me.MenuBar = "Menu Coeur De Pro 2003"

         Me.Toolbar = "Outil Coeur De Pro 2003"

      Case Else

         Me.MenuBar = "Menu Call Coeur De Pro 2003"

         Me.Toolbar = "Outil Call Coeur De Pro 2003"

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
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   OgtFicheContactsDetails_Change

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Load"

   Resume Exit_Form_Load
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

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(ConDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Contacts Date De Création Est Invalide !", , Me.Caption

         TxtConDateCre.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   ConDateMaj = Date

   Select Case IsNull(ConSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Contacts Sociétes Est Obligatoire !", , Me.Caption

         LstConSocNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(ConFonCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Contacts Fonctions Est Obligatoire !", , Me.Caption

         LstConFonCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(ConCivCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Contacts Civilitées Est Obligatoire !", , Me.Caption

         LstConCivCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(ConNom)
      Case False

         TxtConNom = UCase(TxtConNom)
      Case True

         Cancel = True

         MsgBox "Le Champ Contacts Nom Est Obligatoire !", , Me.Caption

         TxtConNom.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   TxtConPrenom = UCase(TxtConPrenom)

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
   Dim StrSQLContacts As String
   Dim StrSQLSocietes As String

   On Error GoTo Err_Form_AfterUpdate

   Select Case IntModeOld
      Case MODEMODIF

         Select Case BlnChangeAdresse
            Case True

               Select Case MsgBox("Voulez Vous Mettre A Jour L'Addresse Professionnel De La Societe ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLSocietes = "UPDATE TBLSOCIETES SET SocAdresse='" & ConAdresse & "' , SocVilNum=" & ConVilNum & _
                                                " WHERE SocNum=" & ConSocNum & ";"

                     CurrentDb.Execute StrSQLSocietes

                  Case vbNo

               End Select

               Select Case MsgBox("Voulez Vous Mettre A Jour L'Addresse Professionnel Des Autres Contacts ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLContacts = "UPDATE TBLCONTACTS SET ConAdresse='" & ConAdresse & "' , ConVilNum=" & ConVilNum & _
                                                " WHERE ConSocNum=" & ConSocNum & ";"

                     CurrentDb.Execute StrSQLContacts

                  Case vbNo

               End Select

               BlnChangeAdresse = False

            Case False

         End Select

         Select Case BlnChangeTelephone
            Case True

               Select Case MsgBox("Voulez Vous Mettre A Jour Les Telephones Professionnel De La Societe ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLSocietes = "UPDATE TBLSOCIETES SET SocTel1Lib='" & ConTelBLib & "' , SocTel1='" & ConTelB & _
                                                "' , SocTel2Lib='" & ConTelDLib & "' , SocTel2='" & ConTelD & "' , SocFax='" & ConFax & _
                                                "' WHERE SocNum=" & ConSocNum & ";"

                     CurrentDb.Execute StrSQLSocietes

                  Case vbNo

               End Select

               Select Case MsgBox("Voulez Vous Mettre A Jour Les Telephones Professionnel Des Autres Contacts ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLContacts = "UPDATE TBLCONTACTS SET ConTelBLib='" & ConTelBLib & "' , ConTelB='" & ConTelB & _
                                                "' , ConTelDLib='" & ConTelDLib & "' , ConTelD='" & ConTelD & "' , ConFax='" & ConFax & _
                                                "' WHERE ConSocNum=" & ConSocNum & ";"

                     CurrentDb.Execute StrSQLContacts

                  Case vbNo

               End Select

               BlnChangeTelephone = False

            Case False

         End Select

      Case Else

         BlnChangeAdresse = False

         BlnChangeTelephone = False

   End Select

   TxtConNum.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtConNum.SetFocus

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
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         TxtConNum.SetFocus

         LstConSocNum.DefaultValue = IIf(IsNull(LstConSocNum.Value), vbNullString, LstConSocNum.Value)

         TxtConAdresse.DefaultValue = IIf(IsNull(TxtConAdresse.Value), vbNullString, """" & TxtConAdresse.Value & """")

         LstConVilNum.DefaultValue = IIf(IsNull(LstConVilNum.Value), vbNullString, LstConVilNum.Value)

         TxtConTelBLib.DefaultValue = IIf(IsNull(TxtConTelBLib.Value), vbNullString, """" & TxtConTelBLib.Value & """")

         TxtConTelB.DefaultValue = IIf(IsNull(TxtConTelB.Value), vbNullString, """" & TxtConTelB.Value & """")

         TxtConTelDLib.DefaultValue = IIf(IsNull(TxtConTelDLib.Value), vbNullString, """" & TxtConTelDLib.Value & """")

         TxtConTelD.DefaultValue = IIf(IsNull(TxtConTelD.Value), vbNullString, """" & TxtConTelD.Value & """")

         TxtConFax.DefaultValue = IIf(IsNull(TxtConFax.Value), vbNullString, """" & TxtConFax.Value & """")

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtConDateCre

   TxtConDateCre = "01/01/2003"

   TxtConDateCre = TxtDateCreer

   TxtConNum.SetFocus

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
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtConNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtConDateCre

         TxtConDateCre = "01/01/2003"

         TxtConDateCre = TxtDateCreer

         TxtConNum.SetFocus

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

         TxtConNum.SetFocus

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

         On Error GoTo Err_CmdSupprimer_Click

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                TxtConNum.SetFocus

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

   TxtConNum.SetFocus

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

   TxtConNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtConNum.SetFocus

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
Private Sub CmdAfficherSocietes_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   Select Case IsNull(ConSocNum)
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & ConSocNum

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

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & ConSocNum, acFormEdit

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
   Dim RsContacts As DAO.Recordset

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

         Set RsContacts = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case Err.Number
            Case 0

               RsContacts.Close

               ExporteExcelRequete (StrSQLRequete)

            Case Else

               Err.Clear

               MsgBox "Impossible D'Exporter " & IIf(Me.FilterOn, "Avec Le Filtre En Cours", vbNullString) & " !!!"

         End Select

   End Select

Exit_CmdExporteExcel:

   Set RsContacts = Nothing

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
Private Sub CmdAjouterVilles_Click()

   On Error GoTo Err_CmdAjouterVilles_Click

   AjouteVilles

Exit_CmdAjouterVilles_Click:

   Exit Sub

Err_CmdAjouterVilles_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterVilles_Click"

   Resume Exit_CmdAjouterVilles_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterVillesPer_Click()

   On Error GoTo Err_CmdAjouterVillesPer_Click

   AjouteVilles

Exit_CmdAjouterVillesPer_Click:

   Exit Sub

Err_CmdAjouterVillesPer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterVillesPer_Click"

   Resume Exit_CmdAjouterVillesPer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterSocietes_Click()

   On Error GoTo Err_CmdAjouterSocietes_Click

   AjouteSocietes

Exit_CmdAjouterSocietes_Click:

   Exit Sub

Err_CmdAjouterSocietes_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterSocietes_Click"

   Resume Exit_CmdAjouterSocietes_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterFonctions_Click()

   On Error GoTo Err_CmdAjouterFonctions_Click

   AjouteFonctions

Exit_CmdAjouterFonctions_Click:

   Exit Sub

Err_CmdAjouterFonctions_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterFonctions_Click"

   Resume Exit_CmdAjouterFonctions_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterCivilites_Click()

   On Error GoTo Err_CmdAjouterCivilites_Click

   AjouteCivilites

Exit_CmdAjouterCivilites_Click:

   Exit Sub

Err_CmdAjouterCivilites_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterCivilites_Click"

   Resume Exit_CmdAjouterCivilites_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstOrganigramme_DblClick(Cancel As Integer)
   Dim RsContacts As DAO.Recordset
   Dim LngConNum As Long

    On Error GoTo Err_LstOrganigramme_DblClick

    Select Case IntMode
      Case MODEDEPLACEMENT

         Select Case IsNull(LstOrganigramme.Column(4))
            Case True

            Case False

               LngConNum = LstOrganigramme.Column(4)

               Me.Filter = vbNullString

               Set RsContacts = Me.RecordsetClone

               RsContacts.FindFirst "ConNum=" & LngConNum

               Me.Bookmark = RsContacts.Bookmark

               RsContacts.Close

         End Select

      Case Else

   End Select

Exit_LstOrganigramme_DblClick:

   Set RsContacts = Nothing

    Exit Sub

Err_LstOrganigramme_DblClick:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstOrganigramme_DblClick"

    Resume Exit_LstOrganigramme_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstOrganigramme_Enter()

   On Error GoTo Err_LstOrganigramme_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_LstOrganigramme_Enter:

    Exit Sub

Err_LstOrganigramme_Enter:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstOrganigramme_Enter"

    Resume Exit_LstOrganigramme_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstOrganigramme_Exit(Cancel As Integer)

    On Error GoTo Err_LstOrganigramme_Exit

    Me.AllowEdits = BlnModif

Exit_LstOrganigramme_Exit:

    Exit Sub

Err_LstOrganigramme_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstOrganigramme_Exit"

    Resume Exit_LstOrganigramme_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub OgtFicheContactsDetails_Change()
   Dim StrSQLConVilNum As String
   Dim StrSQLConPerVilNum As String
   Dim StrSQLOrganigrame As String

   On Error GoTo Err_OgtFicheContactsDetails_Change

   StrSQLConVilNum = "SELECT SelVillesDepartementsCode.VilNum, SelVillesDepartementsCode.VilCodePostal, SelVillesDepartementsCode.VilNom, TBLDEPARTEMENTS.DepPays FROM SelVillesDepartementsCode LEFT JOIN TBLDEPARTEMENTS ON SelVillesDepartementsCode.VilDepCode = TBLDEPARTEMENTS.DepCode ORDER BY SelVillesDepartementsCode.VilCodePostal;"

   StrSQLConPerVilNum = "SELECT SelVillesDepartementsCode.VilNum, SelVillesDepartementsCode.VilCodePostal, SelVillesDepartementsCode.VilNom, TBLDEPARTEMENTS.DepPays FROM SelVillesDepartementsCode LEFT JOIN TBLDEPARTEMENTS ON SelVillesDepartementsCode.VilDepCode = TBLDEPARTEMENTS.DepCode ORDER BY SelVillesDepartementsCode.VilCodePostal;"

   StrSQLOrganigrame = "SELECT SelContactsFonctions.* " & _
                                  "FROM SelContactsFonctions WHERE SelContactsFonctions.ConNum<>[TxtConNum] AND SelContactsFonctions.ConSocNum=[LstConSocNum];"

   Select Case OgtFicheContactsDetails
      Case 0

         LstConVilNum.RowSource = StrSQLConVilNum

         LstConPerVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

      Case 1

         LstConVilNum.RowSource = vbNullString

         LstConPerVilNum.RowSource = StrSQLConPerVilNum

         LstOrganigramme.RowSource = vbNullString

      Case 2

         LstConVilNum.RowSource = vbNullString

         LstConPerVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

      Case 3

         LstConVilNum.RowSource = vbNullString

         LstConPerVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = StrSQLOrganigrame

   End Select

   Rafraichir (RAFRAICHIRCHOIX)

Exit_OgtFicheContactsDetails_Change:

   Exit Sub

Err_OgtFicheContactsDetails_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OgtFicheContactsDetails_Change"

   Resume Exit_OgtFicheContactsDetails_Change
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConCivCode_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstConCivCode_NotInList

   AjouteCivilites (NewData)

   Response = acDataErrContinue

   TxtConCivCode = LstConCivCode.Text

   LstConCivCode.Undo

Exit_LstConCivCode_NotInList:

   Exit Sub

Err_LstConCivCode_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConCivCode_NotInList"

   Resume Exit_LstConCivCode_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConFonCode_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstConFonCode_NotInList

   AjouteFonctions (NewData)

   Response = acDataErrContinue

   TxtConFonCode = LstConFonCode.Text

   LstConFonCode.Undo

Exit_LstConFonCode_NotInList:

   Exit Sub

Err_LstConFonCode_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConFonCode_NotInList"

   Resume Exit_LstConFonCode_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConSocNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstConSocNum_NotInList

   AjouteSocietes (NewData)

   Response = acDataErrContinue

   TxtConSocNum = LstConSocNum.Text

   LstConSocNum.Undo

Exit_LstConSocNum_NotInList:

   Exit Sub

Err_LstConSocNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConSocNum_NotInList"

   Resume Exit_LstConSocNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConVilNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstConVilNum_NotInList

   AjouteVilles (NewData)

   Response = acDataErrContinue

   TxtConVilNum = LstConVilNum.Text

   LstConVilNum.Undo

Exit_LstConVilNum_NotInList:

   Exit Sub

Err_LstConVilNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConVilNum_NotInList"

   Resume Exit_LstConVilNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConPerVilNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstConPerVilNum_NotInList

   AjouteVilles (NewData)

   Response = acDataErrContinue

   TxtConPerVilNum = LstConPerVilNum.Text

   LstConPerVilNum.Undo

Exit_LstConPerVilNum_NotInList:

   Exit Sub

Err_LstConPerVilNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConPerVilNum_NotInList"

   Resume Exit_LstConPerVilNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConSocNum_AfterUpdate()
   Dim RsSocietes As DAO.Recordset
   Dim StrSQLSocietes As String

   On Error GoTo Err_LstConSocNum_AfterUpdate

   Select Case MsgBox("Voulais Vous Copier Les Coordonnees Professionel De L'Entreprise ?", vbYesNo, Me.Caption)
      Case vbYes

         StrSQLSocietes = "SELECT * FROM TBLSOCIETES WHERE SocNum=" & ConSocNum & ";"

         Set RsSocietes = CurrentDb.OpenRecordset(StrSQLSocietes)

         Select Case RsSocietes.EOF
            Case True

            Case False

               ConAdresse = RsSocietes!SocAdresse

               ConVilNum = RsSocietes!SocVilNum

               ConTelBLib = RsSocietes!SocTel1Lib

               ConTelB = RsSocietes!SocTel1

               ConTelDLib = RsSocietes!SocTel2Lib

               ConTelD = RsSocietes!SocTel2

               ConFax = RsSocietes!SocFax

         End Select

         RsSocietes.Close

      Case vbNo

   End Select

Exit_LstConSocNum_AfterUpdate:

   Set RsSocietes = Nothing

   Exit Sub

Err_LstConSocNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConSocNum_AfterUpdate"

   Resume Exit_LstConSocNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConFonCode_AfterUpdate()

   On Error GoTo Err_LstConFonCode_AfterUpdate

Exit_LstConFonCode_AfterUpdate:

   Exit Sub

Err_LstConFonCode_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConFonCode_AfterUpdate"

   Resume Exit_LstConFonCode_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConAdresse_AfterUpdate()

   On Error GoTo Err_TxtConAdresse_AfterUpdate

   BlnChangeAdresse = True

Exit_TxtConAdresse_AfterUpdate:

   Exit Sub

Err_TxtConAdresse_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConAdresse_AfterUpdate"

   Resume Exit_TxtConAdresse_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConVilNum_AfterUpdate()

   On Error GoTo Err_LstConVilNum_AfterUpdate

   BlnChangeAdresse = True

Exit_LstConVilNum_AfterUpdate:

   Exit Sub

Err_LstConVilNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConVilNum_AfterUpdate"

   Resume Exit_LstConVilNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConTelBLib_AfterUpdate()

   On Error GoTo Err_TxtConTelBLib_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtConTelBLib_AfterUpdate:

   Exit Sub

Err_TxtConTelBLib_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConTelBLib_AfterUpdate"

   Resume Exit_TxtConTelBLib_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConTelB_AfterUpdate()

   On Error GoTo Err_TxtConTelB_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtConTelB_AfterUpdate:

   Exit Sub

Err_TxtConTelB_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConTelB_AfterUpdate"

   Resume Exit_TxtConTelB_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConTelDLib_AfterUpdate()

   On Error GoTo Err_TxtConTelDLib_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtConTelDLib_AfterUpdate:

   Exit Sub

Err_TxtConTelDLib_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConTelDLib_AfterUpdate"

   Resume Exit_TxtConTelDLib_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConTelD_AfterUpdate()

   On Error GoTo Err_TxtConTelD_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtConTelD_AfterUpdate:

   Exit Sub

Err_TxtConTelD_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConTelD_AfterUpdate"

   Resume Exit_TxtConTelD_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConFax_AfterUpdate()

   On Error GoTo Err_TxtConFax_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtConFax_AfterUpdate:

   Exit Sub

Err_TxtConFax_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConFax_AfterUpdate"

   Resume Exit_TxtConFax_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConFonCode_GotFocus()

   On Error GoTo Err_LstConFonCode_GotFocus

   LstConFonCode.Dropdown

Exit_LstConFonCode_GotFocus:

   Exit Sub

Err_LstConFonCode_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConFonCode_GotFocus"

   Resume Exit_LstConFonCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConSocNum_GotFocus()

   On Error GoTo Err_LstConSocNum_GotFocus

   LstConSocNum.Dropdown

Exit_LstConSocNum_GotFocus:

   Exit Sub

Err_LstConSocNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConSocNum_GotFocus"

   Resume Exit_LstConSocNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConCivCode_GotFocus()

   On Error GoTo Err_LstConCivCode_GotFocus

   LstConCivCode.Dropdown

Exit_LstConCivCode_GotFocus:

   Exit Sub

Err_LstConCivCode_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConCivCode_GotFocus"

   Resume Exit_LstConCivCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConVilNum_GotFocus()

   On Error GoTo Err_LstConVilNum_GotFocus

   LstConVilNum.Dropdown

Exit_LstConVilNum_GotFocus:

   Exit Sub

Err_LstConVilNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConVilNum_GotFocus"

   Resume Exit_LstConVilNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConPerVilNum_GotFocus()

   On Error GoTo Err_LstConPerVilNum_GotFocus

   LstConPerVilNum.Dropdown

Exit_LstConPerVilNum_GotFocus:

   Exit Sub

Err_LstConPerVilNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConPerVilNum_GotFocus"

   Resume Exit_LstConPerVilNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConSocNum_LostFocus()

   On Error GoTo Err_LstConSocNum_LostFocus

   TxtConSocNum = vbNullString

   Rafraichir (RAFRAICHIRCHOIX)

Exit_LstConSocNum_LostFocus:

   Exit Sub

Err_LstConSocNum_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConSocNum_LostFocus"

   Resume Exit_LstConSocNum_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConFonCode_LostFocus()

   On Error GoTo Err_LstConFonCode_LostFocus

   TxtConFonCode = vbNullString

   Rafraichir (RAFRAICHIRCHOIX)

Exit_LstConFonCode_LostFocus:

   Exit Sub

Err_LstConFonCode_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConFonCode_LostFocus"

   Resume Exit_LstConFonCode_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConCivCode_LostFocus()

   On Error GoTo Err_LstConCivCode_LostFocus

   TxtConCivCode = vbNullString

Exit_LstConCivCode_LostFocus:

   Exit Sub

Err_LstConCivCode_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConCivCode_LostFocus"

   Resume Exit_LstConCivCode_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConVilNum_LostFocus()

   On Error GoTo Err_LstConVilNum_LostFocus

   TxtConCivCode = vbNullString

Exit_LstConVilNum_LostFocus:

   Exit Sub

Err_LstConVilNum_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConVilNum_LostFocus"

   Resume Exit_LstConVilNum_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstConPerVilNum_LostFocus()

   On Error GoTo Err_LstConPerVilNum_LostFocus

   TxtConPerVilNum = vbNullString

Exit_LstConPerVilNum_LostFocus:

   Exit Sub

Err_LstConPerVilNum_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstConPerVilNum_LostFocus"

   Resume Exit_LstConPerVilNum_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConNom_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtConNom_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtConNom_KeyPress:

    Exit Sub

Err_TxtConNom_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConNom_KeyPress"

    Resume Exit_TxtConNom_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtConPrenom_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtConPrenom_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtConPrenom_KeyPress:

    Exit Sub

Err_TxtConPrenom_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtConPrenom_KeyPress"

    Resume Exit_TxtConPrenom_KeyPress
End Sub
