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
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)

   On Error GoTo Err_Rafraichir

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

         LstSocietesAppels.Requery

         LstSocAppNom.Requery

      Case Else

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
   Dim RsFichier As DAO.Recordset

   On Error GoTo Err_Form_Open

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicType='FICCHM' AND FicCode LIKE '*=*' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         Me.HelpFile = RsFichier!FicValeur & "\" & RsFichier!FicCode2

   End Select

   RsFichier.Close

Exit_Form_Open:

    Set RsFichier = Nothing

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

   LstSocietesAppels.Selected(1) = True

   Me.Filter = "SocNum=" & Nz(LstSocietesAppels, 0)

   Me.FilterOn = True

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

   On Error GoTo Err_Form_Load

   TermineMessagerie

   Quit

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
Private Sub ImgMenuCall_Click()

    On Error GoTo Err_ImgMenuCall_Click

    DoCmd.OpenForm "FrmAbout", acNormal, , , acFormReadOnly, acDialog

Exit_ImgMenuCall_Click:

    Exit Sub

Err_ImgMenuCall_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "ImgMenuCall_Click"

    Resume Exit_ImgMenuCall_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   FermeFormulaire "FrmMenuCall"

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdQuitter_Click"

   Resume Exit_CmdQuitter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocietesAppels_DblClick(Cancel As Integer)

   On Error GoTo Err_LstSocietesAppels_DblClick

   Select Case IsNull(LstSocietesAppels.Column(0))
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall3", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall3", acNormal, , "SocNum=" & LstSocietesAppels.Column(0), , acWindowNormal

   End Select

Exit_LstSocietesAppels_DblClick:

   Exit Sub

Err_LstSocietesAppels_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocietesAppels_DblClick"

   Resume Exit_LstSocietesAppels_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstSocietesAppels_AfterUpdate()

   On Error GoTo Err_LstSocietesAppels_AfterUpdate

   Me.Filter = "SocNum=" & LstSocietesAppels

   Me.FilterOn = True

Exit_LstSocietesAppels_AfterUpdate:

   Exit Sub

Err_LstSocietesAppels_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstSocietesAppels_AfterUpdate"

   Resume Exit_LstSocietesAppels_AfterUpdate
End Sub
