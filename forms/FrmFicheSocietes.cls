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
'***     Declaration De Variable Public                                                         ***
'******************************************************************************
Public IntMode As Integer
Public IntModeOld As Integer
Public BlnModif As Boolean
Public BlnChangeAdresse As Boolean
Public BlnChangeTelephone As Boolean

'******************************************************************************
'***     Declaration De Variable Priver                                                         ***
'******************************************************************************
Private TxtSocTypCode As String
Private TxtSocSocNum As String
Private TxtSocConNum As String
Private TxtSocVilNum As String

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
   Dim StrSQLSocietesTypes As String
   Dim StrSQLSocietesFilialeDebut As String
   Dim StrSQLSocietesFilialeFin As String
   Dim StrSQLChargerAffaireDebut As String
   Dim StrSQLChargerAffaireFin As String
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   StrSQLSocietesTypes = "SELECT SocTypCode, SocTypLibelle FROM TBLSOCIETESTYPES ORDER BY SocTypCode;"

   StrSQLSocietesFilialeDebut = "SELECT SocNum, SocRaison, SocSTA, " & _
                                              "SocAdh, SocTel1, SocTel2, SocTypCode FROM TBLSOCIETES "

   StrSQLSocietesFilialeFin = " ORDER BY SocRaison;"

   StrSQLChargerAffaireDebut = "SELECT ConNum, ConNom, ConPrenom, ConCivCode, ConInscrit, " & _
                                                "ConFonCode, SocRaison, SocNum FROM TBLSOCIETES INNER JOIN TBLCONTACTS ON " & _
                                                "TBLSOCIETES.SocNum = TBLCONTACTS.ConSocNum "

   StrSQLChargerAffaireFin = " ORDER BY ConNom;"

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

                      Set CtlActive = Me.TxtSocNum

               End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               LstSocTypCode.RowSource = StrSQLSocietesTypes

               Select Case TxtSocTypCode
                  Case vbNullString

                     LstSocTypCode.Requery

                  Case Else

                     LstSocTypCode.Requery

                     LstSocTypCode.Text = TxtSocTypCode

                     TxtSocTypCode = vbNullString

               End Select

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT, RAFRAICHIRCHOIX

                  Select Case LstSocTypCode.Column(0)
                     Case "SD", "DR"

                        LstSocConNum.RowSource = vbNullString

                        Select Case IsNull(SocConNum)
                           Case True

                           Case False

                              SocConNum = Null

                        End Select

                        Select Case IsNull(SocSTA)
                           Case True

                           Case False

                              SocSTA = Null

                        End Select

                        Select Case IsNull(SocSTAOld)
                           Case True

                           Case False

                              SocSTAOld = Null

                        End Select

                        Select Case LstSocTypCode.Column(0)
                           Case "SD"

                              LstSocSocNum.RowSource = vbNullString

                              Select Case IsNull(SocSocNum)
                                  Case True

                                  Case False

                                     SocSocNum = Null

                               End Select

                           Case "DR"

                              LstSocSocNum.RowSource = StrSQLSocietesFilialeDebut & " WHERE SocTypCode='SD' " & StrSQLSocietesFilialeFin

                        End Select

                        Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
                           Case "GESTION"

                           Case Else

                        End Select

                        Select Case PgeSocietesFiliale.Visible = (LstSocTypCode.Column(0) = "DR")
                           Case True

                           Case False

                              PgeSocietesFiliale.Visible = (LstSocTypCode.Column(0) = "DR")

                        End Select

                        Select Case PgeSocietesChargeAffaire.Visible = False
                           Case True

                           Case False

                              PgeSocietesChargeAffaire.Visible = False

                              PgeResultats.Visible = False

                              PgeCadeaux.Visible = False

                              PgeChaudieres.Visible = False

                              PgeMouvement.Visible = False

                              CmdCadeaux.Visible = False

                              CmdChaudieres.Visible = False

                              CmdBonusAnimation.Visible = False

                              CmdRegularisationCheques.Visible = False

                              TxtSocSTA.Visible = False

                              TxtSocSTAOld.Visible = False

                              BslSocAdh.Visible = False

                        End Select

                     Case "STA"

                        LstSocSocNum.RowSource = vbNullString

                        LstSocConNum.RowSource = StrSQLChargerAffaireDebut & " WHERE ConFonCode='ATC' " & StrSQLChargerAffaireFin

                        Select Case PgeSocietesFiliale.Visible = False
                           Case True

                           Case False

                              PgeSocietesFiliale.Visible = False

                              PgeSocietesChargeAffaire.Visible = True

                              TxtSocSTA.Visible = True

                              TxtSocSTAOld.Visible = True

                              BslSocAdh.Visible = True

                        End Select

                        Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
                           Case "GESTION"

                              Select Case PgeResultats.Visible = True
                                 Case True

                                 Case False

                                    PgeResultats.Visible = True

                                    PgeCadeaux.Visible = True

                                    PgeChaudieres.Visible = True

                                    PgeMouvement.Visible = True

                                    PgeCadeaux.Visible = True

                                    PgeChaudieres.Visible = True

                                    PgeMouvement.Visible = True

                                    CmdCadeaux.Visible = True

                                    CmdChaudieres.Visible = True

                                    CmdBonusAnimation.Visible = True

                                    CmdRegularisationCheques.Visible = True

                              End Select

                           Case Else

                              Select Case PgeResultats.Visible = False
                                 Case True

                                 Case False

                                    PgeResultats.Visible = False

                                    PgeCadeaux.Visible = False

                                    PgeChaudieres.Visible = False

                                    PgeMouvement.Visible = False

                                    PgeCadeaux.Visible = False

                                    PgeChaudieres.Visible = False

                                    PgeMouvement.Visible = False

                                    CmdCadeaux.Visible = False

                                    CmdChaudieres.Visible = False

                                    CmdBonusAnimation.Visible = False

                                    CmdRegularisationCheques.Visible = False

                           End Select

                        End Select

                  End Select

               Case Else

            End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

               Select Case OgtSocietesFiliale
                  Case 0

                     Select Case TxtSocSocNum
                        Case vbNullString

                           LstSocSocNum.Requery

                        Case Else

                           LstSocSocNum.Requery

                           LstSocSocNum.Text = TxtSocSocNum

                           TxtSocSocNum = vbNullString

                     End Select

                  Case 1

                     Select Case TxtSocConNum
                        Case vbNullString

                           LstSocConNum.Requery

                        Case Else

                           LstSocConNum.Requery

                           LstSocConNum.Text = TxtSocConNum

                           TxtSocConNum = vbNullString

                     End Select

                     LstSocAdhSocNum.Requery

               End Select
            Case Else

         End Select

         Select Case OgtFicheSocietesDetails
            Case 0

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

                     Select Case TxtSocVilNum
                        Case vbNullString

                           LstSocVilNum.Requery

                        Case Else

                           LstSocVilNum.Requery

                           LstSocVilNum.Text = TxtSocVilNum

                           TxtSocVilNum = vbNullString

                     End Select

                  Case Else

               End Select

            Case 1

            Case 2

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     Select Case IsNull(SocNum)
                        Case True

                        Case False

                           Select Case IntMode
                              Case MODEDEPLACEMENT

                                 tvwClassement.SelectedItem = tvwClassement.Nodes("SocNum" & SocNum)

                              Case Else

                           End Select

                     End Select

                  Case Else

               End Select

            Case 3

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstOrganigramme.Requery

                  Case Else

               End Select

            Case 4

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstMouvements.Requery

                  Case Else

               End Select

            Case 5

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstChaudieres.Requery

                  Case Else

               End Select

            Case 6

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstCadeaux.Requery

                  Case Else

               End Select

            Case 7

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

                     SFrmFicheSocietesResultatsChaudieresSemestres.Requery

                     SFrmFicheSocietesResultatsSoldesExercices.Requery

                     SFrmFicheSocietesResultatsISOExercices.Requery

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
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub CreeClassement()
   Dim rsSocietesFiliales As DAO.Recordset
   Dim ndsClassement As MSComctlLib.Nodes
   Dim ndeClassement As MSComctlLib.Node
   Dim lngSocNumNiv0 As Long
   Dim lngSocNumNiv1 As Long
   Dim lngSocNumNiv2 As Long

   On Error GoTo Err_CreeClassement

   Set rsSocietesFiliales = CurrentDb.OpenRecordset("SelSocietesFiliales")

   Set ndsClassement = tvwClassement.Nodes

   ndsClassement.Clear

   tvwClassement.Style = tvwTreelinesPlusMinusText

   lngSocNumNiv0 = vbEmpty

   lngSocNumNiv1 = vbEmpty

   lngSocNumNiv2 = vbEmpty

   Do While Not rsSocietesFiliales.EOF

      Select Case rsSocietesFiliales!SocNumNiv0
         Case lngSocNumNiv0

            Select Case rsSocietesFiliales!SocNumNiv1
               Case lngSocNumNiv1

                  Select Case rsSocietesFiliales!SocNumNiv2
                     Case lngSocNumNiv2

                        rsSocietesFiliales.MoveNext

                     Case Else

                        lngSocNumNiv2 = rsSocietesFiliales!SocNumNiv2

                        Set ndeClassement = ndsClassement.Add("SocNum" & lngSocNumNiv1, tvwChild, "SocNum" & lngSocNumNiv2, rsSocietesFiliales!SocTypCodeNiv2 & "-" & rsSocietesFiliales!SocRaisonNiv2 & "-" & RemplaceChr(rsSocietesFiliales!SocAdresseNiv2, vbCrLf, " - "))

                  End Select

               Case Else

                     lngSocNumNiv1 = rsSocietesFiliales!SocNumNiv1

                     Set ndeClassement = ndsClassement.Add("SocNum" & lngSocNumNiv0, tvwChild, "SocNum" & lngSocNumNiv1, rsSocietesFiliales!SocTypCodeNiv1 & "-" & rsSocietesFiliales!SocRaisonNiv1)

            End Select

         Case Else

            lngSocNumNiv0 = rsSocietesFiliales!SocNumNiv0

            Set ndeClassement = ndsClassement.Add(, , "SocNum" & lngSocNumNiv0, rsSocietesFiliales!SocTypCodeNiv0 & "-" & rsSocietesFiliales!SocRaisonNiv0)

      End Select

   Loop

Exit_CreeClassement:

   rsSocietesFiliales.Close

   Set rsSocietesFiliales = Nothing

   Set ndsClassement = Nothing

   Exit Sub

Err_CreeClassement:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeClassement"

   Resume Exit_CreeClassement
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

   CreeClassement

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

   Select Case LstSocTypCode.Column(0)
      Case "SD", "DR"

      Case "STA"

         Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
            Case "GESTION"

               OgtFicheSocietesDetails = 7

            Case Else

         End Select

   End Select

   OgtFicheSocietesDetails_Change

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

   Select Case IsDate(SocDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Sociétés Date De Creation Est Invalide !", , Me.Caption

         TxtSocDateCre.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   SocDateMaj = Date

   Select Case IsNull(SocRaison)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Raison Sociale Est Obligatoire !", , Me.Caption

         TxtSocRaison.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(LstSocTypCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Types Est Obligatoire !", , Me.Caption

         LstSocTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case LstSocTypCode.Column(0)
      Case "SD"

      Case "DR"

         Select Case IsNull(LstSocSocNum)
            Case False

            Case True

               Cancel = True

               MsgBox "Le Champ Sociétés Filiales Est Obligatoire !", , Me.Caption

               LstSocSocNum.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select

      Case "STA"

         Select Case IsNull(LstSocConNum)
            Case False

            Case True

               Cancel = True

               MsgBox "Le Champ Sociétés Chargées D'Affaire Est Obligatoire !", , Me.Caption

               LstSocConNum.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select

   End Select

   Select Case IsNull(SocAdresse)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Adresse Est Obligatoire !", , Me.Caption

         TxtSocAdresse.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocVilNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Ville Est Obligatoire !", , Me.Caption

         LstSocVilNum.SetFocus

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
   Dim BlnAdhEtat As Variant
   Dim DatAdh As Date
   Dim StrSQLSocietesAdherents As String
   Dim StrSQLContacts As String

   On Error GoTo Err_Form_AfterUpdate

   StrSQLSocietesAdherents = "DELETE TBLSOCIETESADHERENTS.* FROM TBLSOCIETESADHERENTS WHERE SocAdhSocNum=" & SocNum & " AND SocAdhDate=CDate('" & Date & "');"

   BlnAdhEtat = DLookup("SocAdhEtat", "SelSocietesAdherentsExercices", "SocAdhSocNum=" & SocNum & " AND ExeExercice=" & Year(Date))

   Select Case IsNull(BlnAdhEtat) Or BlnAdhEtat <> SocAdh
      Case True

         StrSQLSocietesAdherents = "INSERT INTO TBLSOCIETESADHERENTS (SocAdhSocNum,SocAdhExeNum,SocAdhDate,SocAdhEtat) VALUES (" & SocNum & "," & DLookup("ExeNum", "TBLEXERCICES", "ExeExercice=" & Year(Date)) & ",Date() ," & SocAdh & ") ;"

         CurrentDb.Execute StrSQLSocietesAdherents

      Case False

   End Select

   Select Case IntModeOld
      Case MODEMODIF

         Select Case BlnChangeAdresse
            Case True

               Select Case MsgBox("Voulez Vous Mettre A Jour L'Addresse Professionnel Des Contacts ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLContacts = "UPDATE TBLCONTACTS SET ConAdresse='" & SocAdresse & "' , ConVilNum=" & SocVilNum & _
                                                " WHERE ConSocNum=" & SocNum & ";"

                     CurrentDb.Execute StrSQLContacts

                  Case vbNo

               End Select

               BlnChangeAdresse = False

            Case False

         End Select

         Select Case BlnChangeTelephone
            Case True

               Select Case MsgBox("Voulez Vous Mettre A Jour Les Telephones Professionnel Des Contacts ?", vbQuestion + vbYesNo, Me.Caption)
                  Case vbYes

                     StrSQLContacts = "UPDATE TBLCONTACTS SET ConTelBLib='" & SocTel1Lib & "' , ConTelB='" & SocTel1 & _
                                                "' , ConTelDLib='" & SocTel2Lib & "' , ConTelD='" & SocTel2 & "' , ConFax='" & SocFax & _
                                                "' WHERE ConSocNum=" & SocNum & ";"

                     CurrentDb.Execute StrSQLContacts

                  Case vbNo

               End Select

               BlnChangeTelephone = False

            Case False

         End Select

      Case MODEAJOUT

         BlnChangeAdresse = False

         BlnChangeTelephone = False

   End Select

   CreeClassement

   TxtSocNum.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtSocNum.SetFocus

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
Private Sub LstSocVilNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_LstSocVilNum_NotInList

   AjouteVilles (NewData)

   Response = acDataErrContinue

   TxtSocVilNum = LstSocVilNum.Text

   LstSocVilNum.Undo

Exit_LstSocVilNum_NotInList:

   Exit Sub

Err_LstSocVilNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocVilNum_NotInList"

   Resume Exit_LstSocVilNum_NotInList
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

         TxtSocNum.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtSocDateCre

   TxtSocDateCre = "01/01/2003"

   TxtSocDateCre = TxtDateCreer

   TxtSocNum.SetFocus

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

         TxtSocNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtSocDateCre

         TxtSocDateCre = "01/01/2003"

         TxtSocDateCre = TxtDateCreer

         TxtSocNum.SetFocus

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

         TxtSocNum.SetFocus

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

               TxtSocNum.SetFocus

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

   TxtSocNum.SetFocus

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

   TxtSocNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtSocNum.SetFocus

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

   DoCmd.Close acForm, Me.Name, acSavePrompt

   On Error Resume Next

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

   On Error GoTo Exit_CmdRechercher_Click

   OuvreRecherches 1

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
Private Sub CmdAfficherContacts_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherContacts_Click

   DonneEtatFormulaire "FrmFicheContacts", IntEtat

   Select Case IntEtat
      Case ETATFORMULAIRE

         Forms!FrmFicheContacts.SetFocus

         On Error Resume Next

         DoCmd.ApplyFilter , "ConSocNum=" & Nz(SocNum, 0)

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

         DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConSocNum=" & Nz(SocNum, 0), acFormEdit

   End Select

   Forms!FrmFicheContacts.LstConSocNum.DefaultValue = SocNum

   Forms!FrmFicheContacts.LstConSocNum.DefaultValue = IIf(IsNull(SocNum), vbNullString, SocNum)

   Forms!FrmFicheContacts.TxtConAdresse.DefaultValue = IIf(IsNull(SocAdresse.Value), vbNullString, """" & SocAdresse & """")

   Forms!FrmFicheContacts.LstConVilNum.DefaultValue = IIf(IsNull(SocVilNum), vbNullString, SocVilNum)

   Forms!FrmFicheContacts.TxtConFax.DefaultValue = IIf(IsNull(SocFax), vbNullString, """" & SocFax & """")

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
Private Sub CmdChaudieres_Click()

   On Error GoTo Err_CmdChaudieres_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         OuvreChaudieres True, CLng(SocNum)

   End Select

Exit_CmdChaudieres_Click:

   Exit Sub

Err_CmdChaudieres_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdChaudieres_Click"

   Resume Exit_CmdChaudieres_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdCadeaux_Click()

   On Error GoTo Err_CmdCadeaux_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         OuvreCadeaux True, CLng(SocNum)

   End Select

Exit_CmdCadeaux_Click:

   Exit Sub

Err_CmdCadeaux_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdCadeaux_Click"

   Resume Exit_CmdCadeaux_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdBonusAnimation_Click()

   On Error GoTo Err_CmdBonusAnimation_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         OuvreBonusAnimation True, CLng(SocNum)

   End Select

Exit_CmdBonusAnimation_Click:

   Exit Sub

Err_CmdBonusAnimation_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdBonusAnimation_Click"

   Resume Exit_CmdBonusAnimation_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdRegularisationCheques_Click()

   On Error GoTo Err_CmdRegularisationCheques_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         OuvreRegularisationCheques True, CLng(SocNum)

   End Select

Exit_CmdRegularisationCheques_Click:

   Exit Sub

Err_CmdRegularisationCheques_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdRegularisationCheques_Click"

   Resume Exit_CmdRegularisationCheques_Click
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
Private Sub CmdExporteExcelOnglet_Click()
   Dim StrSQLRequete As String

   On Error GoTo Err_CmdExporteExcelOnglet

   Select Case OgtFicheSocietesDetails
      Case 4

         StrSQLRequete = RemplaceChr(LstMouvements.RowSource, "[TxtSocNum]", TxtSocNum)

      Case 5

         StrSQLRequete = RemplaceChr(RemplaceChr(LstChaudieres.RowSource, "[TxtSocNum]", TxtSocNum), "[LstChaudieresExercices]", LstChaudieresExercices)

      Case 6

         StrSQLRequete = RemplaceChr(LstCadeaux.RowSource, "[TxtSocNum]", TxtSocNum)

   End Select

   Select Case StrSQLRequete
      Case vbNullString

      Case Else

         ExporteExcelRequete (StrSQLRequete)

   End Select

Exit_CmdExporteExcelOnglet:

   Exit Sub

Err_CmdExporteExcelOnglet:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteExcelOnglet"

   Resume Exit_CmdExporteExcelOnglet
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
   Dim RsSocietes As DAO.Recordset

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

         Set RsSocietes = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case Err.Number
            Case 0

               RsSocietes.Close

               ExporteExcelRequete (StrSQLRequete)

            Case Else

               Err.Clear

               MsgBox "Impossible D'Exporter " & IIf(Me.FilterOn, "Avec Le Filtre En Cours", vbNullString) & " !!!"

         End Select

   End Select

Exit_CmdExporteExcel:

   Set RsSocietes = Nothing

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
Private Sub LstMouvements_DblClick(Cancel As Integer)

   On Error GoTo Err_LstMouvements_DblClick

   Select Case IsNull(LstMouvements.Column(2))
      Case True

      Case False

         Select Case LstMouvements.Column(2)
            Case MVTNATURESOLDE

            Case MVTNATURECADEAUX

               OuvreCadeaux False, CLng(LstMouvements.Column(9))

            Case MVTNATURECHAUDIERES

               OuvreChaudieres False, CLng(LstMouvements.Column(9))

            Case MVTNATUREBONUSANIMATION

         End Select

   End Select

Exit_LstMouvements_DblClick:

   Exit Sub

Err_LstMouvements_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMouvements_DblClick"

   Resume Exit_LstMouvements_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadeaux_DblClick(Cancel As Integer)

   On Error GoTo Err_LstCadeaux_DblClick

   Select Case IsNull(LstCadeaux.Column(1))
      Case True

      Case False

         OuvreCadeaux False, CLng(LstCadeaux.Column(1))

   End Select

Exit_LstCadeaux_DblClick:

   Exit Sub

Err_LstCadeaux_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadeaux_DblClick"

   Resume Exit_LstCadeaux_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieres_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChaudieres_DblClick

   Select Case IsNull(LstChaudieres.Column(1))
      Case True

      Case False

         OuvreChaudieres False, CLng(LstChaudieres.Column(1))

   End Select

Exit_LstChaudieres_DblClick:

   Exit Sub

Err_LstChaudieres_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieres_DblClick"

   Resume Exit_LstChaudieres_DblClick
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
   Dim IntEtat As Integer

   On Error GoTo Err_LstOrganigramme_DblClick

   Select Case IsNull(LstOrganigramme.Column(4))
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheContacts", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheContacts.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "ConNum=" & LstOrganigramme.Column(4)

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_LstOrganigramme_DblClick

               End Select

               On Error GoTo Err_LstOrganigramme_DblClick

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & LstOrganigramme.Column(4), acFormEdit

         End Select
   End Select

Exit_LstOrganigramme_DblClick:

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
Private Sub tvwClassement_Click()
   Dim RsSocietes As DAO.Recordset
   Dim StrSocNum As String

   On Error GoTo Err_tvwClassement_Click

   Select Case IntMode
      Case MODEDEPLACEMENT

      Case Else

         GoTo Exit_tvwClassement_Click

   End Select

   StrSocNum = Mid(Me.tvwClassement.SelectedItem.Key, 7)

   Me.Filter = vbNullString

   Me.Requery

   Set RsSocietes = Me.RecordsetClone

   RsSocietes.FindFirst "SocNum=" & StrSocNum

   Select Case RsSocietes.NoMatch
      Case True

      Case False

         Me.Undo

         Me.Bookmark = RsSocietes.Bookmark

   End Select

   RsSocietes.Close

   Set RsSocietes = Nothing

Exit_tvwClassement_Click:

    Exit Sub

Err_tvwClassement_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "tvwClassement_Click"

    Resume Exit_tvwClassement_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadeaux_Enter()

   On Error GoTo Err_LstCadeaux_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_LstCadeaux_Enter:

   Exit Sub

Err_LstCadeaux_Enter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadeaux_Enter"

   Resume Exit_LstCadeaux_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub tvwClassement_Enter()

   On Error GoTo Err_tvwClassement_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_tvwClassement_Enter:

   Exit Sub

Err_tvwClassement_Enter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "tvwClassement_Enter"

   Resume Exit_tvwClassement_Enter
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
Private Sub LstMouvements_Enter()

   On Error GoTo Err_LstMouvements_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_LstMouvements_Enter:

   Exit Sub

Err_LstMouvements_Enter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMouvements_Enter"

   Resume Exit_LstMouvements_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieres_Enter()

   On Error GoTo Err_LstChaudieres_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_LstChaudieres_Enter:

   Exit Sub

Err_LstChaudieres_Enter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieres_Enter"

   Resume Exit_LstChaudieres_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieresExercices_Enter()

   On Error GoTo Err_LstChaudieresExercices_Enter

   BlnModif = Me.AllowEdits

   Me.AllowEdits = True

Exit_LstChaudieresExercices_Enter:

   Exit Sub

Err_LstChaudieresExercices_Enter:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieresExercices_Enter"

   Resume Exit_LstChaudieresExercices_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstCadeaux_Exit(Cancel As Integer)

   On Error GoTo Err_LstCadeaux_Exit

   Me.AllowEdits = BlnModif

Exit_LstCadeaux_Exit:

   Exit Sub

Err_LstCadeaux_Exit:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstCadeaux_Exit"

   Resume Exit_LstCadeaux_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstMouvements_Exit(Cancel As Integer)

   On Error GoTo Err_LstMouvements_Exit

   Me.AllowEdits = BlnModif

Exit_LstMouvements_Exit:

   Exit Sub

Err_LstMouvements_Exit:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMouvements_Exit"

   Resume Exit_LstMouvements_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieres_Exit(Cancel As Integer)

   On Error GoTo Err_LstChaudieres_Exit

   Me.AllowEdits = BlnModif

Exit_LstChaudieres_Exit:

   Exit Sub

Err_LstChaudieres_Exit:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieres_Exit"

   Resume Exit_LstChaudieres_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieresExercices_Exit(Cancel As Integer)

   On Error GoTo Err_LstChaudieresExercices_Exit

   Me.AllowEdits = BlnModif

Exit_LstChaudieresExercices_Exit:

   Exit Sub

Err_LstChaudieresExercices_Exit:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieresExercices_Exit"

   Resume Exit_LstChaudieresExercices_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub tvwClassement_Exit(Cancel As Integer)

   On Error GoTo Err_tvwClassement_Exit

   Me.AllowEdits = BlnModif

Exit_tvwClassement_Exit:

   Exit Sub

Err_tvwClassement_Exit:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "tvwClassement_Exit"

   Resume Exit_tvwClassement_Exit
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
Private Sub LstSocConNum_AfterUpdate()

   On Error GoTo Err_LstSocConNum_AfterUpdate

   SocSocNum = LstSocConNum.Column(7)

Exit_LstSocConNum_AfterUpdate:

   Exit Sub

Err_LstSocConNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocConNum_AfterUpdate"

   Resume Exit_LstSocConNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocTypCode_AfterUpdate()

   On Error GoTo Err_LstSocTypCode_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

Exit_LstSocTypCode_AfterUpdate:

   Exit Sub

Err_LstSocTypCode_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocTypCode_AfterUpdate"

   Resume Exit_LstSocTypCode_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocVilNum_AfterUpdate()

   On Error GoTo Err_LstSocVilNum_AfterUpdate

   BlnChangeAdresse = True

Exit_LstSocVilNum_AfterUpdate:

   Exit Sub

Err_LstSocVilNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocVilNum_AfterUpdate"

   Resume Exit_LstSocVilNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocAdresse_AfterUpdate()

   On Error GoTo Err_TxtSocAdresse_AfterUpdate

   BlnChangeAdresse = True

Exit_TxtSocAdresse_AfterUpdate:

   Exit Sub

Err_TxtSocAdresse_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocAdresse_AfterUpdate"

   Resume Exit_TxtSocAdresse_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocTel1Lib_AfterUpdate()

   On Error GoTo Err_TxtSocTel1Lib_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtSocTel1Lib_AfterUpdate:

   Exit Sub

Err_TxtSocTel1Lib_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocTel1Lib_AfterUpdate"

   Resume Exit_TxtSocTel1Lib_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocTel1_AfterUpdate()

   On Error GoTo Err_TxtSocTel1_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtSocTel1_AfterUpdate:

   Exit Sub

Err_TxtSocTel1_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocTel1_AfterUpdate"

   Resume Exit_TxtSocTel1_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocTel2Lib_AfterUpdate()

   On Error GoTo Err_TxtSocTel2Lib_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtSocTel2Lib_AfterUpdate:

   Exit Sub

Err_TxtSocTel2Lib_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocTel2Lib_AfterUpdate"

   Resume Exit_TxtSocTel2Lib_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocTel2_AfterUpdate()

   On Error GoTo Err_TxtSocTel2_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtSocTel2_AfterUpdate:

   Exit Sub

Err_TxtSocTel2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocTel2_AfterUpdate"

   Resume Exit_TxtSocTel2_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocFax_AfterUpdate()

   On Error GoTo Err_TxtSocFax_AfterUpdate

   BlnChangeTelephone = True

Exit_TxtSocFax_AfterUpdate:

   Exit Sub

Err_TxtSocFax_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocFax_AfterUpdate"

   Resume Exit_TxtSocFax_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChaudieresExercices_AfterUpdate()

   On Error GoTo Err_LstChaudieresExercices_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

Exit_LstChaudieresExercices_AfterUpdate:

   Exit Sub

Err_LstChaudieresExercices_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChaudieresExercices_AfterUpdate"

   Resume Exit_LstChaudieresExercices_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub OgtFicheSocietesDetails_Change()
   Dim StrSQLSocVilNum As String
   Dim StrSQLOrganigramme As String
   Dim StrSQLMouvements As String
   Dim StrSQLChaudieres As String
   Dim StrSQLChaudieresEnCours As String
   Dim StrSQLCadeaux As String
   Dim StrSFrmFicheSocietesResultatsChaudieresSemestres As String
   Dim StrSFrmFicheSocietesResultatsSoldesExercices As String
   Dim StrSFrmFicheSocietesResultatsISOExercices As String

   On Error GoTo Err_OgtFicheSocietesDetails_Change

   StrSQLSocVilNum = "SELECT VilNum, VilCodePostal, VilNom, DepPays FROM SelVillesDepartementsCode LEFT JOIN TBLDEPARTEMENTS ON SelVillesDepartementsCode.VilDepCode = TBLDEPARTEMENTS.DepCode ORDER BY VilCodePostal;"

   StrSQLOrganigramme = "SELECT SelContactsFonctions.* " & _
                                     "FROM SelContactsFonctions WHERE ConSocNum=[TxtSocNum];"

   StrSQLMouvements = "SELECT MvtSocNum, MvtDate, MvtNatNum, MvtNatCode, " & _
                                  "ExeExercice, MvtIntitule, MvtDebit, " & _
                                  "MvtCredit, MvtQuantite, MvtDossier FROM SelMouvementsNaturesExercices WHERE " & _
                                  "SelMouvementsNaturesExercices.MvtSocNum=[TxtSocNum];"

   StrSQLChaudieres = "SELECT ChdRemSocNum, ChdRemNum, " & _
                                  "ChdRemChdArtRef, ChdArtLibelle, ChdRemSerie, " & _
                                  "SemLibelle, ChdRemServ, ChdRemDate " & _
                                  "FROM SelChaudieresArticlesSemestres WHERE ChdRemSocNum=[TxtSocNum] AND ([LstChaudieresExercices]=-1 Or [LstChaudieresExercices]=Year(ChdRemServ));"

   StrSQLCadeaux = "SELECT CadAchSocNum, CadAchNum, CadArtRef, CadArtDescription, " & _
                              "CadAchDate, CadAchDateSaisie, CadAchEtat, CadAchQuantite, " & _
                              "CadAchPoints FROM SelCadeauxArticles WHERE CadAchSocNum=[TxtSocNum];"

   StrSFrmFicheSocietesResultatsChaudieresSemestres = "SFrmFicheSocietesResultatsChaudieresSemestres"

   StrSFrmFicheSocietesResultatsSoldesExercices = "SFrmFicheSocietesResultatsSoldesExercices"

   StrSFrmFicheSocietesResultatsISOExercices = "SFrmFicheSocietesResultatsISOExercices"

   Select Case OgtFicheSocietesDetails
      Case 0

         LstSocVilNum.RowSource = StrSQLSocVilNum

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = False

      Case 1

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = False

      Case 2

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = False

      Case 3

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = StrSQLOrganigramme

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = False

      Case 4

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = StrSQLMouvements

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = True

      Case 5

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = StrSQLChaudieres

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = True

      Case 6

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = StrSQLCadeaux

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = vbNullString

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = vbNullString

         SFrmFicheSocietesResultatsISOExercices.SourceObject = vbNullString

         CmdExporteExcelOnglet.Visible = True

      Case 7

         LstSocVilNum.RowSource = vbNullString

         LstOrganigramme.RowSource = vbNullString

         LstMouvements.RowSource = vbNullString

         LstChaudieres.RowSource = vbNullString

         LstCadeaux.RowSource = vbNullString

         SFrmFicheSocietesResultatsChaudieresSemestres.SourceObject = StrSFrmFicheSocietesResultatsChaudieresSemestres

         SFrmFicheSocietesResultatsSoldesExercices.SourceObject = StrSFrmFicheSocietesResultatsSoldesExercices

         SFrmFicheSocietesResultatsISOExercices.SourceObject = StrSFrmFicheSocietesResultatsISOExercices

         CmdExporteExcelOnglet.Visible = False

   End Select

   TxtSocNum.SetFocus

   Rafraichir (RAFRAICHIRCHOIX)

Exit_OgtFicheSocietesDetails_Change:

   Exit Sub

Err_OgtFicheSocietesDetails_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OgtFicheSocietesDetails_Change"

   Resume Exit_OgtFicheSocietesDetails_Change
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub tvwClassement_BeforeLabelEdit(Cancel As Integer)

    On Error GoTo Err_tvwClassement_BeforeLabelEdit

    Cancel = True

Exit_tvwClassement_BeforeLabelEdit:

    Exit Sub

Err_tvwClassement_BeforeLabelEdit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "tvwClassement_BeforeLabelEdit"

    Resume Exit_tvwClassement_BeforeLabelEdit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocTypCode_GotFocus()

   On Error GoTo Err_LstSocTypCode_GotFocus

   LstSocTypCode.Dropdown

Exit_LstSocTypCode_GotFocus:

   Exit Sub

Err_LstSocTypCode_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocTypCode_GotFocus"

   Resume Exit_LstSocTypCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocSocNum_GotFocus()

   On Error GoTo Err_LstSocSocNum_GotFocus

   LstSocSocNum.Dropdown

Exit_LstSocSocNum_GotFocus:

   Exit Sub

Err_LstSocSocNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocSocNum_GotFocus"

   Resume Exit_LstSocSocNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocConNum_GotFocus()

   On Error GoTo Err_LstSocConNum_GotFocus

   LstSocConNum.Dropdown

Exit_LstSocConNum_GotFocus:

   Exit Sub

Err_LstSocConNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocConNum_GotFocus"

   Resume Exit_LstSocConNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocVilNum_GotFocus()

   On Error GoTo Err_LstSocVilNum_GotFocus

   LstSocVilNum.Dropdown

Exit_LstSocVilNum_GotFocus:

   Exit Sub

Err_LstSocVilNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocVilNum_GotFocus"

   Resume Exit_LstSocVilNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocVilNum_LostFocus()

   On Error GoTo Err_LstSocVilNum_LostFocus

   TxtSocVilNum = vbNullString

Exit_LstSocVilNum_LostFocus:

   Exit Sub

Err_LstSocVilNum_LostFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocVilNum_LostFocus"

   Resume Exit_LstSocVilNum_LostFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocRaison_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtSocRaison_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtSocRaison_KeyPress:

    Exit Sub

Err_TxtSocRaison_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocRaison_KeyPress"

    Resume Exit_TxtSocRaison_KeyPress
End Sub
