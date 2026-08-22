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
'***     Declaration De Variable Priver                                                         ***
'******************************************************************************
Private TxtSocVilNum As String

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

   SocDateMaj = Date

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
