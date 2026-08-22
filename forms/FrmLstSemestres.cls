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
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
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
                  Case 0

                  Case Else

                     Err.Clear

                     Set CtlActive = TxtSemNum

               End Select

               On Error GoTo Err_Rafraichir

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               LstSemExeNum.Requery

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

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Open"

   Resume Err_Form_Open
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Dim BytSemNum As Byte

    On Error GoTo Err_Form_BeforeUpdate

    Select Case Nz(SemExeNum, vbEmpty) > vbEmpty
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Semestre Exercice Est Invalide !", , Me.Caption

         LstSemExeNum.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case Nz(SemExeSemNum, vbEmpty) > vbEmpty
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Semestre Exercice Semestre Numero Est Invalide !", , Me.Caption

         TxtSemExeSemNum.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsDate(SemDebut)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Semestre Debut Est Invalide !", , Me.Caption

         TxtSemDebut.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsDate(SemFin)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Semestre Fin Est Invalide !", , Me.Caption

         TxtSemFin.SetFocus

         GoTo Exit_Form_BeforeUpdate
   End Select

   Select Case IsNull(SemLibelle)
      Case True

         Cancel = True

         MsgBox "Le Champ Semestre Libelle Est Obligatoire !", , Me.Caption

         TxtSemLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

    BytSemNum = Nz(DLookup("SemNum", "TBLSEMESTRES", "SemDebut<=CDate('" & TxtSemDebut & "') AND CDate('" & TxtSemDebut & "')<=SemFin"), vbEmpty)

    Select Case BytSemNum
        Case vbEmpty

        Case Else

            Select Case BytSemNum
                Case SemNum

                Case Else

                    Cancel = True

                    MsgBox "Le Debut du Semestre Est Contenue Dans Un Autre Semestre !", , "Semestre Debut"

                    GoTo Exit_Form_BeforeUpdate
            End Select
    End Select

    BytSemNum = Nz(DLookup("SemNum", "TBLSEMESTRES", "SemDebut<=CDate('" & TxtSemFin & "') AND CDate('" & TxtSemDebut & "')<=SemFin"), vbEmpty)

    Select Case BytSemNum
        Case vbEmpty

        Case Else

            Select Case BytSemNum
                Case SemNum

                Case Else

                    Cancel = True

                    MsgBox "La Fin du Semestre Est Contenue Dans Un Autre Semestre !", , "Semestre Fin"

                    GoTo Exit_Form_BeforeUpdate
            End Select
    End Select

    Select Case CStr(Year(TxtSemDebut))
        Case LstSemExeNum.Column(1)

        Case Else

            Cancel = True

            MsgBox "La Debut du Semestre Invalide Il N'Est Pas Sur L'Exercice " & LstSemExeNum.Column(1) & " !", , "Semestre Debut"

            GoTo Exit_Form_BeforeUpdate
    End Select

    Select Case CStr(Year(TxtSemFin))
        Case LstSemExeNum.Column(1)

        Case Else

            Cancel = True

            MsgBox "La Fin du Seméstre Invalide Il N'Est Pas Sur L'Exercice " & LstSemExeNum.Column(1) & " !", , "Semestre Fin"

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

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
Private Sub CmdRechercher_Click()

   On Error GoTo Exit_CmdRechercher_Click

   OuvreRecherches 3

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
Private Sub CmdExporteExcel_Click()
   Dim StrSQLRequete As String
   Dim RsSemestres As DAO.Recordset

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

         Set RsSemestres = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case Err.Number
            Case 0

               RsSemestres.Close

               ExporteExcelRequete (StrSQLRequete)

            Case Else

               Err.Clear

               MsgBox "Impossible D'Exporter " & IIf(Me.FilterOn, "Avec Le Filtre En Cours", vbNullString) & " !!!"

         End Select

   End Select

Exit_CmdExporteExcel:

   Set RsSemestres = Nothing

   Exit Sub

Err_CmdExporteExcel:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteExcel"

   Resume Exit_CmdExporteExcel
End Sub
