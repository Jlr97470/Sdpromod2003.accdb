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
'***  Declaration De Procedure Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Function DonneFiltreChaudieresRemontees() As String
   Dim StrChaudieresRemontees As String
   Dim VarItem As Variant

   On Error GoTo Err_DonneFiltreChaudieresRemontees

   DonneFiltreChaudieresRemontees = vbNullString

   Select Case LstLstChaudieresRemontees.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstChaudieresRemontees.ItemsSelected.Count
            Case 1

               StrChaudieresRemontees = "ChdRemNum=" & LstLstChaudieresRemontees.Column(0)

            Case Else

               StrChaudieresRemontees = "ChdRemNum In("

               For Each VarItem In LstLstChaudieresRemontees.ItemsSelected

                  StrChaudieresRemontees = StrChaudieresRemontees & LstLstChaudieresRemontees.Column(0, VarItem) & ", "

               Next

               StrChaudieresRemontees = Left(StrChaudieresRemontees, Len(StrChaudieresRemontees) - 2) & ")"

         End Select

         DonneFiltreChaudieresRemontees = StrChaudieresRemontees

   End Select

Exit_DonneFiltreChaudieresRemontees:

   Exit Function

Err_DonneFiltreChaudieresRemontees:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltreChaudieresRemontees"

   Resume Exit_DonneFiltreChaudieresRemontees
End Function

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)

   On Error GoTo Err_Rafraichir

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         LstLstChaudieresRemontees.Requery

      Case Else

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub ValideChaudieresRemontees()
   Dim IntEtat As Integer
   Dim StrSQLRequete As String

   On Error GoTo Err_ValideChaudieresRemontees

   Select Case LstLstChaudieresRemontees.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstChaudieresRemontees.ItemsSelected.Count = (LstLstChaudieresRemontees.ListCount - 1)
            Case True

               StrSQLRequete = GStrChaudieresRemonteesFiltre

            Case False

               StrSQLRequete = DonneFiltreChaudieresRemontees

         End Select

         DonneEtatFormulaire "FrmSaisieChaudieresRemontees", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieChaudieresRemontees.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLRequete

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ValideChaudieresRemontees

               End Select

               On Error GoTo Err_ValideChaudieresRemontees

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

              DoCmd.OpenForm "FrmSaisieChaudieresRemontees", acNormal, , StrSQLRequete, acFormEdit

         End Select

   End Select

Exit_ValideChaudieresRemontees:

   Exit Sub

Err_ValideChaudieresRemontees:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideChaudieresRemontees"

   Resume Exit_ValideChaudieresRemontees
End Sub

'******************************************************************************
'***  Declaration D'Evenement Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***  EVENEMENT:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub Form_Open(Cancel As Integer)
   Dim StrSQLChaudieresRemonteesDebut As String
   Dim StrSQLChaudieresRemonteesFin As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_Form_Open

   StrSQLChaudieresRemonteesDebut = "SELECT DISTINCT [SelSocietesSTAContactsTAInsChaudieresRemontees].* " & _
                                                           " FROM [SelSocietesSTAContactsTAInsChaudieresRemontees] "

   StrSQLChaudieresRemonteesFin = ";"

   Select Case GStrChaudieresRemonteesFiltre
      Case vbNullString

         LstLstChaudieresRemontees.RowSource = StrSQLChaudieresRemonteesDebut & StrSQLChaudieresRemonteesFin

      Case Else

         DonneFiltreElement GStrChaudieresRemonteesFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

         Select Case StrFiltreTableRequeteNom
            Case "[SelSocietesSTAContactsTAInsChaudieresRemontees]"

               LstLstChaudieresRemontees.RowSource = StrSQLChaudieresRemonteesDebut & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLChaudieresRemonteesFin

            Case Else

               LstLstChaudieresRemontees.RowSource = StrSQLChaudieresRemonteesDebut & " INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesSTAContactsTAInsChaudieresRemontees]." & StrFiltreChampSource & "=" & _
                   StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLChaudieresRemonteesFin

         End Select

   End Select

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
Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   ValideChaudieresRemontees

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   Select Case Err.Number
      Case 2501

      Case Else

         MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdFermer_Click"

   End Select

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

   OuvreRecherches 4

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
Private Sub CmdExporteExcel_Click()
   Dim StrSQLRequete As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_CmdExporteExcel

   Select Case LstLstChaudieresRemontees.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstChaudieresRemontees.ItemsSelected.Count = (LstLstChaudieresRemontees.ListCount - 1)
            Case True

               Select Case GStrChaudieresRemonteesFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsChaudieresRemontees].* FROM [SelSocietesSTAContactsTAInsChaudieresRemontees];"

                  Case Else

                     DonneFiltreElement GStrChaudieresRemonteesFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelSocietesSTAContactsTAInsChaudieresRemontees]"

                           StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsChaudieresRemontees].* FROM [SelSocietesSTAContactsTAInsChaudieresRemontees] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsChaudieresRemontees].* FROM [SelSocietesSTAContactsTAInsChaudieresRemontees] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesSTAContactsTAInsChaudieresRemontees]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsChaudieresRemontees].* FROM [SelSocietesSTAContactsTAInsChaudieresRemontees] WHERE " & DonneFiltreChaudieresRemontees & ";"

         End Select

         Select Case StrSQLRequete
            Case vbNullString

            Case Else

               ExporteExcelRequete (StrSQLRequete)

         End Select

   End Select

Exit_CmdExporteExcel:

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
Private Sub CmdSelectionnerTous_Click()
   Dim LngIndex As Long

   On Error GoTo Err_CmdSelectionnerTous_Click

   For LngIndex = 0 To LstLstChaudieresRemontees.ListCount - 1

      LstLstChaudieresRemontees.Selected(LngIndex) = Not LstLstChaudieresRemontees.Selected(LngIndex)

   Next

   TxtLstChaudieresRemonteesCount.Requery

Exit_CmdSelectionnerTous_Click:

   Exit Sub

Err_CmdSelectionnerTous_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSelectionnerTous_Click"

   Resume Exit_CmdSelectionnerTous_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstLstChaudieresRemontees_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstChaudieresRemontees_DblClick

   ValideChaudieresRemontees

Exit_LstLstChaudieresRemontees_DblClick:

   Exit Sub

Err_LstLstChaudieresRemontees_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstChaudieresRemontees_DblClick"

   Resume Exit_LstLstChaudieresRemontees_DblClick
End Sub
