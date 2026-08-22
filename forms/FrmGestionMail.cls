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
   Dim StrSQLEmailFichiersBody As String

   On Error GoTo Err_Rafraichir

   StrSQLEmailFichiersBody = "SELECT SelFichiersDetailler.FicCode2, SelFichiersDetailler.FicValeur FROM SelFichiersDetailler WHERE (((SelFichiersDetailler.FicValide)=True) AND ((SelFichiersDetailler.FicType2)='HTM')); "

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         On Error Resume Next

         Set CtlActive = Me.ActiveControl

         Select Case Err.Number
            Case 0

            Case Else

                Err.Clear

                Set CtlActive = TxtEMailExpediteur

        End Select

        On Error GoTo Err_Rafraichir

   End Select

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         CmbEMailChampDestination.Requery

         CmbEMailFichierBody.RowSource = StrSQLEmailFichiersBody

         CmbEMailFichierBody.Requery

      Case Else

   End Select

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         CtlActive.SetFocus

         Set CtlActive = Nothing

      Case Else

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
Private Sub ChangeEtatBouton()

  On Error GoTo Err_ChangeEtatBouton

   Select Case Not IsNull(TxtEMailExpediteur) And Not IsNull(CmbEMailChampDestination) And Not IsNull(TxtEMailSujet) And Not IsNull(CmbEMailFichierBody)
      Case True

         CmdExporteMailing.Enabled = True

      Case False

         CmdExporteMailing.Enabled = False

   End Select

Exit_ChangeEtatBouton:

   Exit Sub

Err_ChangeEtatBouton:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ChangeEtatBouton"

   Resume Exit_ChangeEtatBouton
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
   Dim RsRequete As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSQLRequete As String

   On Error GoTo Err_Form_Load

   Select Case IsNull(Me.OpenArgs)
      Case True

         DoCmd.Close acForm, Me.Name, acSavePrompt

      Case False

         StrSQLRequete = Me.OpenArgs

         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case RsRequete.EOF
            Case True

               MsgBox "Aucunne selection !", , Me.Caption

               DoCmd.Close acForm, Me.Name, acSavePrompt

               GoTo Exit_Form_Load

            Case False

               CmbEMailChampDestination.RowSource = vbNullString

               For Each FldChamp In RsRequete.Fields

                  On Error Resume Next

                  CmbEMailChampDestination.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        CmbEMailChampDestination.AddItem FldChamp.Name & ";"

                  End Select

                  On Error GoTo Err_Form_Load

               Next

         End Select

   End Select

   Rafraichir (RAFRAICHIRVISIBLE)

   ChangeEtatBouton

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdAjouterFichierAttacher_Click()
   Dim FsoSystemFichier As New Scripting.FileSystemObject
   Dim FilFichier As Scripting.File

   On Error GoTo Err_CmdAjouterFichierAttacher_Click

   CdgFichiers.DialogTitle = "AD 2003 - Indiquer L'Emplacement Du Fichier "

   CdgFichiers.ShowOpen

   Select Case FsoSystemFichier.FileExists(CdgFichiers.FileName)
      Case True

         Set FilFichier = FsoSystemFichier.GetFile(CdgFichiers.FileName)

         LstEMailFichiersAttacher.AddItem UCase(FilFichier.Name) & ";" & UCase(FilFichier.ParentFolder.Path)

      Case False

         MsgBox "Aucun Fichier N'a Ete Choisi !", vbInformation, Me.Caption

   End Select

Exit_CmdAjouterFichierAttacher_Click:

   Set FilFichier = Nothing

   Set FsoSystemFichier = Nothing

   Exit Sub

Err_CmdAjouterFichierAttacher_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterFichierAttacher_Click"

   Resume Exit_CmdAjouterFichierAttacher_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdExporteMailing_Click()
   Dim StrFichierAttacher() As String
   Dim IntNb As Integer

   On Error GoTo Err_CmdExporteMailing_Click

   For IntNb = 0 To LstEMailFichiersAttacher.ListCount - 1

      ReDim Preserve StrFichierAttacher(IntNb)

      StrFichierAttacher(IntNb) = LstEMailFichiersAttacher.Column(1, IntNb) & "\" & LstEMailFichiersAttacher.Column(0, IntNb)

   Next

   ExporteMailingRequete IIf(IsNull(TxtEMailExpediteur), vbNullString, TxtEMailExpediteur), IIf(IsNull(TxtEMailCopieCacher), vbNullString, TxtEMailCopieCacher), Me.OpenArgs, CmbEMailChampDestination, IIf(IsNull(TxtEMailSujet), vbNullString, TxtEMailSujet), CmbEMailFichierBody.Column(0), StrFichierAttacher, ChkEmailSimulation

Exit_CmdExporteMailing_Click:

   Exit Sub

Err_CmdExporteMailing_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteMailing_Click"

   Resume Exit_CmdExporteMailing_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstEMailFichiersAttacher_DblClick(Cancel As Integer)

   On Error GoTo Err_LstEMailFichiersAttacher_DblClick

   LstEMailFichiersAttacher.RemoveItem LstEMailFichiersAttacher.ListIndex

Exit_LstEMailFichiersAttacher_DblClick:

   Exit Sub

Err_LstEMailFichiersAttacher_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstEMailFichiersAttacher_DblClick"

   Resume Exit_LstEMailFichiersAttacher_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub TxtEMailExpediteur_AfterUpdate()

   On Error GoTo Err_TxtEMailExpediteur_AfterUpdate

   ChangeEtatBouton

Exit_TxtEMailExpediteur_AfterUpdate:

   Exit Sub

Err_TxtEMailExpediteur_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtEMailExpediteur_AfterUpdate"

   Resume Exit_TxtEMailExpediteur_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmbEMailChampDestination_AfterUpdate()

   On Error GoTo Err_CmbEMailChampDestination_AfterUpdate

   ChangeEtatBouton

Exit_CmbEMailChampDestination_AfterUpdate:

   Exit Sub

Err_CmbEMailChampDestination_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbEMailChampDestination_AfterUpdate"

   Resume Exit_CmbEMailChampDestination_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmbEMailFichierBody_AfterUpdate()

   On Error GoTo Err_CmbEMailFichierBody_AfterUpdate

   ChangeEtatBouton

Exit_CmbEMailFichierBody_AfterUpdate:

   Exit Sub

Err_CmbEMailFichierBody_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbEMailFichierBody_AfterUpdate"

   Resume Exit_CmbEMailFichierBody_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub TxtEMailSujet_AfterUpdate()

   On Error GoTo Err_TxtEMailSujet_AfterUpdate

   ChangeEtatBouton

Exit_TxtEMailSujet_AfterUpdate:

   Exit Sub

Err_TxtEMailSujet_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtEMailSujet_AfterUpdate"

   Resume Exit_TxtEMailSujet_AfterUpdate
End Sub
