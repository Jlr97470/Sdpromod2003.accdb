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
'***  Declaration De Procedure Priver                                                         ***
'******************************************************************************

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Function DonneFiltreCadeauxAchats() As String
   Dim StrCadeauxAchats As String
   Dim VarItem As Variant

   On Error GoTo Err_DonneFiltreCadeauxAchats

   DonneFiltreCadeauxAchats = vbNullString

   Select Case LstLstCadeauxAchats.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstCadeauxAchats.ItemsSelected.Count
            Case 1

               StrCadeauxAchats = "CadAchNum=" & LstLstCadeauxAchats.Column(0)

            Case Else

               StrCadeauxAchats = "CadAchNum In("

               For Each VarItem In LstLstCadeauxAchats.ItemsSelected

                  StrCadeauxAchats = StrCadeauxAchats & LstLstCadeauxAchats.Column(0, VarItem) & ", "

               Next

               StrCadeauxAchats = Left(StrCadeauxAchats, Len(StrCadeauxAchats) - 2) & ")"

         End Select

         DonneFiltreCadeauxAchats = StrCadeauxAchats

   End Select

Exit_DonneFiltreCadeauxAchats:

   Exit Function

Err_DonneFiltreCadeauxAchats:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltreCadeauxAchats"

   Resume Exit_DonneFiltreCadeauxAchats
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

         LstLstCadeauxAchats.Requery

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
Private Sub ValideCadeauxAchats()
   Dim IntEtat As Integer
   Dim StrSQLRequete As String

   On Error GoTo Err_ValideCadeauxAchats

   Select Case LstLstCadeauxAchats.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstCadeauxAchats.ItemsSelected.Count = (LstLstCadeauxAchats.ListCount - 1)
            Case True

               StrSQLRequete = GStrCadeauxAchatsFiltre

            Case False

               StrSQLRequete = DonneFiltreCadeauxAchats

         End Select

         DonneEtatFormulaire "FrmSaisieCadeauxAchats", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmSaisieCadeauxAchats.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLRequete

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ValideCadeauxAchats

               End Select

               On Error GoTo Err_ValideCadeauxAchats

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

              DoCmd.OpenForm "FrmSaisieCadeauxAchats", acNormal, , StrSQLRequete, acFormEdit

         End Select

   End Select

Exit_ValideCadeauxAchats:

   Exit Sub

Err_ValideCadeauxAchats:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideCadeauxAchats"

   Resume Exit_ValideCadeauxAchats
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
   Dim StrSQLCadeauxAchatsDebut As String
   Dim StrSQLCadeauxAchatsFin As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_Form_Load

   StrSQLCadeauxAchatsDebut = "SELECT DISTINCT [SelSocietesSTAContactsTAInsCadeauxAchats].* " & _
                                                 " FROM [SelSocietesSTAContactsTAInsCadeauxAchats] "

   StrSQLCadeauxAchatsFin = ";"

   Select Case GStrCadeauxAchatsFiltre
      Case vbNullString

         LstLstCadeauxAchats.RowSource = StrSQLCadeauxAchatsDebut & StrSQLCadeauxAchatsFin

      Case Else

         DonneFiltreElement GStrCadeauxAchatsFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

         Select Case StrFiltreTableRequeteNom
            Case "[SelSocietesSTAContactsTAInsCadeauxAchats]"

               LstLstCadeauxAchats.RowSource = StrSQLCadeauxAchatsDebut & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLCadeauxAchatsFin

            Case Else

               LstLstCadeauxAchats.RowSource = StrSQLCadeauxAchatsDebut & " INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesSTAContactsTAInsCadeauxAchats]." & StrFiltreChampSource & "=" & _
                   StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLCadeauxAchatsFin

         End Select

   End Select

   DoCmd.RunCommand acCmdWindowCascade

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

   ValideCadeauxAchats

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

   Select Case LstLstCadeauxAchats.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstCadeauxAchats.ItemsSelected.Count = (LstLstCadeauxAchats.ListCount - 1)
            Case True

               Select Case GStrCadeauxAchatsFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsCadeauxAchats].* FROM [SelSocietesSTAContactsTAInsCadeauxAchats];"

                  Case Else

                     DonneFiltreElement GStrCadeauxAchatsFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelSocietesSTAContactsTAInsCadeauxAchats]"

                           StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsCadeauxAchats].* FROM [SelSocietesSTAContactsTAInsCadeauxAchats] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsCadeauxAchats].* FROM [SelSocietesSTAContactsTAInsCadeauxAchats] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesSTAContactsTAInsCadeauxAchats]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelSocietesSTAContactsTAInsCadeauxAchats].* FROM [SelSocietesSTAContactsTAInsCadeauxAchats] WHERE " & DonneFiltreCadeauxAchats & ";"

         End Select

         ExporteExcelRequete (StrSQLRequete)

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

   For LngIndex = 0 To LstLstCadeauxAchats.ListCount - 1

      LstLstCadeauxAchats.Selected(LngIndex) = Not LstLstCadeauxAchats.Selected(LngIndex)

   Next

   TxtLstCadeauxAchatsCount.Requery

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
Private Sub LstLstCadeauxAchats_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstCadeauxAchats_DblClick

   ValideCadeauxAchats

Exit_LstLstCadeauxAchats_DblClick:

   Exit Sub

Err_LstLstCadeauxAchats_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstCadeauxAchats_DblClick"

   Resume Exit_LstLstCadeauxAchats_DblClick
End Sub
