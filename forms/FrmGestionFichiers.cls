Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ##MODULE_SUMMARY Module de gestion des importations et des expotations de Donne.
' ##PROJECT_TITLE DeltaInformatique (Frontale)
' ##PROJECT_SUMMARY Interface DeltaInformatique
Option Compare Database
Option Explicit

Private Const ModName = "FrmGestionFichiers"

' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)
   Dim StrSQLFichiersPreference As String
   Dim RsFichiersPreference As DAO.Recordset
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   StrSQLFichiersPreference = "SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe] UNION SELECT TBLMACHINES.MacNom AS FicCode1, '' AS FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacNom, '' HAVING TBLMACHINES.MacNom Not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES.MacUtilisateur AS FicCode1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacUtilisateur,'' HAVING TBLMACHINES.MacUtilisateur Not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES.MacDomaine AS FicCode1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacDomaine,'' HAVING TBLMACHINES.MacDomaine Not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];);"

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         On Error Resume Next

         Set CtlActive = Me.ActiveControl

         Select Case Err.Number
            Case 0

            Case Else

                Err.Clear

                Set CtlActive = CmbFichiersListe

        End Select

        On Error GoTo Err_Rafraichir

   End Select

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

         On Error Resume Next

         CmbFichiersPreference.RowSource = "SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe] ;"

         On Error GoTo Err_Rafraichir

         CmbFichiersListe.Requery

         CmbFichiersPreference.Requery

         LstFichiersManquant.Requery

         LstFichiersValide.Requery

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

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".Rafraichir"

   Resume Exit_Rafraichir
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub ChangeEtatBouton()
   Dim CtlActive As Access.Control

   On Error GoTo Err_ChangeEtatBouton

   On Error Resume Next

   Set CtlActive = Me.ActiveControl

   Select Case Err.Number
      Case 0

      Case Else

          Err.Clear

          Set CtlActive = CmbFichiersListe
  End Select

  On Error GoTo Err_ChangeEtatBouton

   Select Case Not IsNull(CmbFichiersListe) And Not IsNull(CmbFichiersPreference)
      Case True

         CmdValider.Enabled = True

         CmdSupprimer.Enabled = True

      Case False

         CmdValider.Enabled = False

         CmdSupprimer.Enabled = False

   End Select

   CtlActive.SetFocus

   Set CtlActive = Nothing

Exit_ChangeEtatBouton:

   Exit Sub

Err_ChangeEtatBouton:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".ChangeEtatBouton"

   Resume Exit_ChangeEtatBouton
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   Me.Caption = DLookup("ParValeur", "SelParametresDetailler", "ParType='PROG' AND ParCode='NOM'") & " - Gestion Fichiers - " & DLookup("ParValeur", "SelParametresDetailler", "ParType='PROG' AND ParCode='COPYRIGHT'") & "@" & Year(Date)

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description, , "Form_Open"

   Resume Err_Form_Open
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Rafraichir (RAFRAICHIRVISIBLE)

   ChangeEtatBouton

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, , "Form_Load"

   Resume Err_Form_Load
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_Activate:

    Exit Sub

Err_Form_Activate:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".Form_Activate"

   Resume Exit_Form_Activate
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub Form_Current()

   On Error GoTo Err_Form_Current

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

    Exit Sub

Err_Form_Current:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".Form_Current"

   Resume Exit_Form_Current
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmbFichiersListe_AfterUpdate()

   On Error GoTo Err_CmbFichiersListe_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   CmbFichiersPreference = Null

   TxtFichiersRepertoire = vbNullString

   ChangeEtatBouton

Exit_CmbFichiersListe_AfterUpdate:

    Exit Sub

Err_CmbFichiersListe_AfterUpdate:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmbFichiersListe_AfterUpdate"

   Resume Exit_CmbFichiersListe_AfterUpdate
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmbFichiersPreference_AfterUpdate()

   On Error GoTo Err_CmbFichiersPreference_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   TxtFichiersRepertoire = CmbFichiersPreference.Column(1)

   TxtExtension = UCase(Mid(CmbFichiersListe, InStrRev(CmbFichiersListe, ".") + 1))

   ChangeEtatBouton

Exit_CmbFichiersPreference_AfterUpdate:

    Exit Sub

Err_CmbFichiersPreference_AfterUpdate:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmbFichiersPreference_AfterUpdate"

   Resume Exit_CmbFichiersPreference_AfterUpdate
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmdRepertoireFichier_Click()
   Dim FsoSystemFichier As New Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim CdgFichiers As Office.FileDialog

   On Error GoTo Err_CmdRepertoireFichier_Click

   Set CdgFichiers = Application.FileDialog(msoFileDialogFilePicker)

   CdgFichiers.Title = "Indiquer L'Emplacement Du Fichier " & CmbFichiersListe

   CdgFichiers.Filters.Clear

   CdgFichiers.InitialFileName = IIf(IsNull(CmbFichiersListe), vbNullString, IIf(TxtFichiersRepertoire = vbNullString, "", TxtFichiersRepertoire & "\") & Nz(CmbFichiersListe))

   CdgFichiers.AllowMultiSelect = False

   CdgFichiers.Show

   If CdgFichiers.SelectedItems.Count > 0 Then

      Select Case FsoSystemFichier.FileExists(CdgFichiers.SelectedItems(1))
         Case True

            Set FilFichier = FsoSystemFichier.GetFile(CdgFichiers.SelectedItems(1))

            CmbFichiersListe = FilFichier.Name

            TxtFichiersRepertoire = FilFichier.ParentFolder.Path

            TxtExtension = UCase(Mid(CdgFichiers.SelectedItems(1), InStrRev(CdgFichiers.SelectedItems(1), ".") + 1))

         Case False

            MsgBox "Aucun Fichier N'a Ete Choisi", vbInformation, Me.Caption

      End Select

      Rafraichir (RAFRAICHIRCHOIX)

      Select Case CmbFichiersPreference.ListCount
         Case 0

            CmbFichiersPreference = "DEFAULT"

         Case Else

      End Select

      ChangeEtatBouton

   End If

Exit_CmdRepertoireFichier_Click:

   Set FilFichier = Nothing

   Set FsoSystemFichier = Nothing

   Exit Sub

Err_CmdRepertoireFichier_Click:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmdRepertoireFichier_Click"

   Resume Exit_CmdRepertoireFichier_Click
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmdValider_Click()
   Dim FsoSystemFichier As New Scripting.FileSystemObject

   On Error GoTo Err_CmdValider_Click

   Select Case FsoSystemFichier.FileExists(TxtFichiersRepertoire & "\" & CmbFichiersListe)
      Case True

           CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValide=False WHERE FicType='FIC" & TxtExtension & "' AND FicCode LIKE '*=" & CmbFichiersListe & "' ;"

           Select Case IsNull(DLookup("FicType", "TBLFICHIERS", "FicType='FIC" & TxtExtension & "' AND FicCode LIKE '" & CmbFichiersPreference & "=" & CmbFichiersListe & "'"))
              Case True

                 CurrentDb.Execute "INSERT INTO TBLFICHIERS (FicType,FicCode,FicValeur,FicValide) VALUES('FIC" & TxtExtension & "','" & CmbFichiersPreference & "=" & CmbFichiersListe & "','" & TxtFichiersRepertoire & "',True) ;"

              Case False

                 CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValeur='" & TxtFichiersRepertoire & "', FicValide=True WHERE FicType='FIC" & TxtExtension & "' AND FicCode='" & CmbFichiersPreference & "=" & CmbFichiersListe & "' ;"

           End Select

           Select Case TxtExtension
                Case "ACCDB"

                    ChangeTableInformationConnexion

                Case Else

            End Select

        Case False

            MsgBox "Chemin Invalide", , Me.Caption

   End Select

   Rafraichir (RAFRAICHIRCHOIX)

Exit_CmdValider_Click:

   Exit Sub

Err_CmdValider_Click:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmdValider_Click"

   Resume Exit_CmdValider_Click
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   CurrentDb.Execute "DELETE * FROM  TBLFICHIERS WHERE FicType='FIC" & TxtExtension & "' AND FicCode LIKE '" & CmbFichiersPreference & "=" & CmbFichiersListe & "' ;"

   Rafraichir (RAFRAICHIRCHOIX)

   CmbFichiersListe = Null

   CmbFichiersPreference = Null

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmdSupprimer_Click"

   Resume Exit_CmdSupprimer_Click
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmbFichiersListe_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_CmbFichiersListe_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_CmbFichiersListe_KeyPress:

   Exit Sub

Err_CmbFichiersListe_KeyPress:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmbFichiersListe_KeyPress"

   Resume Exit_CmbFichiersListe_KeyPress
End Sub
' ##SUMMARY Fonction qui retourne un Object me d'un fichier.
' ##RETURNS Retourne un Object Clsme du fichier.
' ##PARAM StrFichier - Valeur chaîne du chemin du fichier.
Private Sub CmbFichiersPreference_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_CmbFichiersPreference_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_CmbFichiersPreference_KeyPress:

   Exit Sub

Err_CmbFichiersPreference_KeyPress:

   Dim IntNumError As Long
   Dim StrError As String

   StrError = Err.Number & " " & Err.Description & vbCrLf

   For IntNumError = 0 To Errors.Count - 1

     StrError = StrError & Errors(IntNumError).Number & " " & Errors(IntNumError).Description & vbCrLf

   Next IntNumError

   MsgBox StrError, vbCritical, ModName & ".CmbFichiersPreference_KeyPress"

   Resume Exit_CmbFichiersPreference_KeyPress
End Sub
