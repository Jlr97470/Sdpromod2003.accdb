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
Public BlnModif As Boolean

'******************************************************************************
'***    Declaration De Variable Privée                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Du Numero De Questionnaires En Cours                               ***
'******************************************************************************
Private LngSocQstNum As Long
Private FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats As Access.Form

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
   Dim StrSQLSocQstSocNum As String
   Dim StrSQLSocQstSocQstTypCode As String
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   StrSQLSocQstSocQstTypCode = "SELECT SocQstTypNum, SocQstTypCode, SocQstTypLibelle, SocQstTypDateDebut, SocQstTypDateFin, SocQstTypNbResultat " & _
      " FROM SelSocietesQuestionnairesTypes ;"

   StrSQLSocQstSocNum = "SELECT SocNum, SocSTA, SocRaison, SocAdresse, SocVilCodePostal, SocVilNom, SocTel1, SocTel2, SocFax, ConNum,  ConCivCode, ConNom, ConPrenom, ConTelPort, ConEMail, SocWeb FROM SelSocietesSTAVillesContactsTAInscrit WHERE (((SelSocietesSTAVillesContactsTAInscrit.SocAdh)=True));"

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

                      Set CtlActive = CmbSocQstSocQstTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               CmbSocQstSocQstTypCode.RowSource = StrSQLSocQstSocQstTypCode

               CmbSocQstSocQstTypCode.Requery

               CmbSocQstSocNum.RowSource = StrSQLSocQstSocNum

               CmbSocQstSocNum.Requery

               SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Requery

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

   LngSocQstNum = Nz(SocQstNum)

   On Error Resume Next

   Select Case TypeName(FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats) = "Nothing" And TypeName(SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form) = "Form_SFrmSaisieQuestionnairesTypesQuestionnairesResultats"
      Case True

         Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form

      Case False

    End Select

    Select Case Err.Number
      Case 0

      Case Else

         Err.Clear

   End Select

   On Error GoTo Err_Form_Load

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
   Dim RsQuestionnairesValeur As DAO.Recordset

   On Error GoTo Err_Form_Unload

   Select Case IntMode
      Case MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case 0

               Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = Nothing

            Case Else

               Set RsQuestionnairesValeur = FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.RecordsetClone

               RsQuestionnairesValeur.FindFirst "SocQstTypIDObligatoire=True AND IsNull([SocQstResValeur])=True"

               Select Case RsQuestionnairesValeur.NoMatch
                  Case True

                     Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = Nothing

                  Case False

                     Cancel = True

                     MsgBox RsQuestionnairesValeur!SocQstTypID & " - " & RsQuestionnairesValeur!SocQstTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.Bookmark = RsQuestionnairesValeur.Bookmark

                     SFrmSaisieQuestionnairesTypesQuestionnairesResultats.SetFocus

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats!CmbSocQstResValeur.SetFocus

               End Select

               RsQuestionnairesValeur.Close

         End Select

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

   End Select

Exit_Form_Unload:

   Set RsQuestionnairesValeur = Nothing

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
   Dim RsQuestionnairesValeur As DAO.Recordset

   On Error GoTo Err_Form_ApplyFilter

   Select Case IntMode
      Case MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case 0

               Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = Nothing

            Case Else

               Set RsQuestionnairesValeur = FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.RecordsetClone

               RsQuestionnairesValeur.FindFirst "PerQstTypIDObligatoire=True AND IsNull([PerQstResValeur])=True"

               Select Case RsQuestionnairesValeur.NoMatch
                  Case True

                        Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = Nothing

                  Case False

                     Cancel = True

                     MsgBox RsQuestionnairesValeur!PerQstTypID & " - " & RsQuestionnairesValeur!PerQstTypIDLibelle & vbCrLf & "n'a pas ete repondue !", , Me.Caption

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.Bookmark = RsQuestionnairesValeur.Bookmark

                     SFrmSaisieQuestionnairesTypesQuestionnairesResultats.SetFocus

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats!CmbPerQstResValeur.SetFocus

               End Select

               RsQuestionnairesValeur.Close

         End Select

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

   End Select

Exit_Form_ApplyFilter:

   Set RsQuestionnairesValeur = Nothing

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

   Select Case IsNull(SocQstSocQstTypNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Code Est Obligatoire !", , Me.Caption

         CmbSocQstSocQstTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocQstSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Personnel Numero Est Obligatoire !", , Me.Caption

         CmbSocQstSocNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

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
   Dim StrSQLQuestionnaires As String

   On Error GoTo Err_Form_AfterUpdate

   LngSocQstNum = SocQstNum

   StrSQLQuestionnaires = "DELETE TBLSOCIETESQUESTIONNAIRESRESULTATS.* " & _
                                       " FROM TBLSOCIETESQUESTIONNAIRESRESULTATS LEFT JOIN " & _
                                       " SelSocietesQuestionnairesTypesID ON (TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocNum = " & _
                                       " SelSocietesQuestionnairesTypesID.SocQstSocNum) AND (TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocQstTypIDNum = " & _
                                       " SelSocietesQuestionnairesTypesID.SocQstTypIDNum) WHERE (((SelSocietesQuestionnairesTypesID.SocQstSocNum) Is Null));"

   CurrentDb.Execute StrSQLQuestionnaires

   StrSQLQuestionnaires = "INSERT INTO TBLSOCIETESQUESTIONNAIRESRESULTATS ( SocQstResSocNum, " & _
                                        "  SocQstResSocQstTypIDNum ) SELECT SelSocietesQuestionnairesTypesID.SocQstSocNum, " & _
                                        " SelSocietesQuestionnairesTypesID.SocQstTypIDNum FROM SelSocietesQuestionnairesTypesID " & _
                                        " WHERE SelSocietesQuestionnairesTypesID.SocQstSocNum=" & SocQstSocNum & _
                                        " AND SelSocietesQuestionnairesTypesID.SocQstTypIDSocQstTypNum=" & SocQstSocQstTypNum & ";"

   CurrentDb.Execute StrSQLQuestionnaires

   CmbSocQstSocQstTypCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   CmbSocQstSocQstTypCode.SetFocus

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
   Dim RsQuestionnaires As DAO.Recordset
   Dim RsQuestionnairesValeur As DAO.Recordset
   Dim RsFormQuestionnairesValeur As DAO.Recordset
   Dim StrQuestionnairesValeur As String

   On Error GoTo Err_Form_Current

   On Error Resume Next

   Select Case TypeName(FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats) = "Nothing" And TypeName(SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form) = "Form_SFrmSaisieQuestionnairesTypesQuestionnairesResultats"
      Case True

         Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form

      Case False

    End Select

    Select Case Err.Number
      Case 0

      Case Else

         Err.Clear

   End Select

   On Error GoTo Err_Form_Current

   Select Case IntMode
      Case MODEDEPLACEMENT

         StrQuestionnairesValeur = "SELECT SocQstNum, SocQstTypID, SocQstTypIDObligatoire, SocQstTypIDLibelle, SocQstResValeur " & _
               " FROM TBLSOCIETESQUESTIONNAIRESTYPES INNER JOIN (TBLSOCIETESQUESTIONNAIRESTYPESID INNER JOIN (TBLSOCIETESQUESTIONNAIRESRESULTATS INNER JOIN TBLSOCIETESQUESTIONNAIRES ON TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocNum = TBLSOCIETESQUESTIONNAIRES.SocQstSocNum) ON TBLSOCIETESQUESTIONNAIRESTYPESID.SocQstTypIDNum = TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocQstTypIDNum) ON (TBLSOCIETESQUESTIONNAIRESTYPES.SocQstTypNum = TBLSOCIETESQUESTIONNAIRES.SocQstSocQstTypNum) AND (TBLSOCIETESQUESTIONNAIRESTYPES.SocQstTypNum = TBLSOCIETESQUESTIONNAIRESTYPESID.SocQstTypIDSocQstTypNum) " & _
               " WHERE SocQstNum=" & LngSocQstNum & ";"

         Set RsQuestionnairesValeur = CurrentDb.OpenRecordset(StrQuestionnairesValeur)

         RsQuestionnairesValeur.FindFirst "SocQstTypIDObligatoire=True AND IsNull([SocQstResValeur])=True"

         Select Case RsQuestionnairesValeur.NoMatch
            Case True

               LngSocQstNum = SocQstNum

            Case False

               Select Case LngSocQstNum
                  Case SocQstNum

                  Case Else

                     Set RsQuestionnaires = Me.RecordsetClone

                     RsQuestionnaires.FindFirst "SocQstNum=" & LngSocQstNum

                     Me.Bookmark = RsQuestionnaires.Bookmark

                     MsgBox RsQuestionnairesValeur!SocQstTypID & " - " & RsQuestionnairesValeur!SocQstTypIDLibelle & vbCrLf & "n'a pas ete repondue !", , Me.Caption

                     Set RsFormQuestionnairesValeur = FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.RecordsetClone

                     RsFormQuestionnairesValeur.FindFirst "SocQstTypID=" & RsQuestionnairesValeur!SocQstTypID

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.Bookmark = RsFormQuestionnairesValeur.Bookmark

                     SFrmSaisieQuestionnairesTypesQuestionnairesResultats.SetFocus

                     FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats!CmbSocQstResValeur.SetFocus

                     RsQuestionnaires.Close

                     RsFormQuestionnairesValeur.Close

               End Select

         End Select

         RsQuestionnairesValeur.Close

      Case Else

   End Select

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Set RsQuestionnaires = Nothing

   Set RsQuestionnairesValeur = Nothing

   Set RsFormQuestionnairesValeur = Nothing

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

   On Error Resume Next

   Select Case TypeName(FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats) = "Nothing" And TypeName(SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form) = "Form_SFrmSaisieQuestionnairesTypesQuestionnairesResultats"
      Case True

         Set FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats = SFrmSaisieQuestionnairesTypesQuestionnairesResultats.Form

      Case False

    End Select

    Select Case Err.Number
      Case 0

      Case Else

         Err.Clear

   End Select

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
Private Sub CmdAjouter_Click()
   Dim TxtDateSaisie As String
   Dim RsQuestionnairesValeur As DAO.Recordset

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         CmbSocQstSocQstTypCode.SetFocus

         CmbSocQstSocQstTypCode.DefaultValue = "'" & CmbSocQstSocQstTypCode.Value & "'"

         Set RsQuestionnairesValeur = FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.RecordsetClone

         RsQuestionnairesValeur.FindFirst "SocQstTypIDObligatoire=True AND IsNull([SocQstResValeur])=True"

         Select Case RsQuestionnairesValeur.NoMatch
            Case True

            Case False

               MsgBox RsQuestionnairesValeur!SocQstTypID & " - " & RsQuestionnairesValeur!SocQstTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

               FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.Bookmark = RsQuestionnairesValeur.Bookmark

               SFrmSaisieQuestionnairesTypesQuestionnairesResultats.SetFocus

               GoTo Exit_CmdAjouter_Click

         End Select

         RsQuestionnairesValeur.Close

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateSaisie = TxtSocQstDateSaisie

   TxtSocQstDateSaisie = "01/01/2003"

   TxtSocQstDateSaisie = TxtDateSaisie

   CmbSocQstSocQstTypCode.SetFocus

Exit_CmdAjouter_Click:

   Set RsQuestionnairesValeur = Nothing

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

         CmbSocQstSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateSaisie = TxtSocQstDateSaisie

         TxtSocQstDateSaisie = "01/01/2003"

         TxtSocQstDateSaisie = TxtDateSaisie

         CmbSocQstSocQstTypCode.SetFocus

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
   Dim StrSQLQuestionnaires As String

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbSocQstSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

               StrSQLQuestionnaires = "DELETE TBLSOCIETESQUESTIONNAIRESRESULTATS.* " & _
                                                   " FROM TBLSOCIETESQUESTIONNAIRESRESULTATS LEFT JOIN " & _
                                                   " SelSocietesQuestionnairesTypesID ON (TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocNum = " & _
                                                   " SelSocietesQuestionnairesTypesID.SocQstSocNum) AND (TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocQstTypIDNum = " & _
                                                   " SelSocietesQuestionnairesTypesID.SocQstTypIDNum) WHERE (((SelSocietesQuestionnairesTypesID.SocQstSocNum) Is Null));"

               CurrentDb.Execute StrSQLQuestionnaires

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

                CmbSocQstSocQstTypCode.SetFocus

          End Select

         On Error GoTo Err_CmdSupprimer_Click

   End Select

   LngSocQstNum = Nz(SocQstNum)

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

   CmbSocQstSocQstTypCode.SetFocus

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

   CmbSocQstSocQstTypCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbSocQstSocQstTypCode.SetFocus

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

   OuvreRecherches

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

   Select Case IsNull(CmbSocQstSocNum.Column(9)) Or CmbSocQstSocNum.Column(9) = vbNullString
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheContacts", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheContacts.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "ConNum=" & CmbSocQstSocNum.Column(9)

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

               DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & CmbSocQstSocNum.Column(9), acFormEdit

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

   Select Case IsNull(CmbSocQstSocNum.Column(0))
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & CmbSocQstSocNum.Column(0)

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

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & CmbSocQstSocNum.Column(0), acFormEdit

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
Private Sub SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter()

   On Error GoTo Err_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter

   FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.AllowEdits = True

Exit_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter:

    Exit Sub

Err_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter"

    Resume Exit_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit(Cancel As Integer)

   On Error GoTo Err_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit

   FrmSFrmSaisieQuestionnairesTypesQuestionnairesResultats.AllowEdits = False

Exit_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit:

    Exit Sub

Err_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit"

    Resume Exit_SFrmSaisieQuestionnairesTypesQuestionnairesResultats_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocQstSocQstTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbSocQstSocQstTypCode_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbSocQstSocQstTypCode_KeyPress:

    Exit Sub

Err_CmbSocQstSocQstTypCode_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstSocQstTypCode_KeyPress"

    Resume Exit_CmbSocQstSocQstTypCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocQstSocNum_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbSocQstSocNum_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbSocQstSocNum_KeyPress:

    Exit Sub

Err_CmbSocQstSocNum_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstSocNum_KeyPress"

    Resume Exit_CmbSocQstSocNum_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocQstSocQstTypCode_GotFocus()

    On Error GoTo Err_CmbSocQstSocQstTypCode_GotFocus

  ' CmbSocQstSocQstTypCode.Dropdown

Exit_CmbSocQstSocQstTypCode_GotFocus:

    Exit Sub

Err_CmbSocQstSocQstTypCode_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstSocQstTypCode_GotFocus"

    Resume Exit_CmbSocQstSocQstTypCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocQstSocNum_GotFocus()

    On Error GoTo Err_CmbSocQstSocNum_GotFocus

    CmbSocQstSocNum.Dropdown

Exit_CmbSocQstSocNum_GotFocus:

    Exit Sub

Err_CmbSocQstSocNum_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstSocNum_GotFocus"

    Resume Exit_CmbSocQstSocNum_GotFocus
End Sub
