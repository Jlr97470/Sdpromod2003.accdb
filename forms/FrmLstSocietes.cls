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
Private Function DonneFiltreSocietes() As String
   Dim StrSocietes As String
   Dim VarItem As Variant

   On Error GoTo Err_DonneFiltreSocietes

   DonneFiltreSocietes = vbNullString

   Select Case LstLstSocietes.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstSocietes.ItemsSelected.Count
            Case 1

               StrSocietes = "SocNum=" & LstLstSocietes.Column(0)

            Case Else

               StrSocietes = "SocNum In("

               For Each VarItem In LstLstSocietes.ItemsSelected

                  StrSocietes = StrSocietes & LstLstSocietes.Column(0, VarItem) & ", "

               Next

               StrSocietes = Left(StrSocietes, Len(StrSocietes) - 2) & ")"

         End Select

         DonneFiltreSocietes = StrSocietes

   End Select

Exit_DonneFiltreSocietes:

   Exit Function

Err_DonneFiltreSocietes:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltreSocietes"

   Resume Exit_DonneFiltreSocietes
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

         LstLstSocietes.Requery

         TxtLstSocietesCount.Requery

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
Private Sub ValideSocietes()
   Dim IntEtat As Integer
   Dim StrSQLRequete As String

   On Error GoTo Err_ValideSocietes

   Select Case LstLstSocietes.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstSocietes.ItemsSelected.Count = (LstLstSocietes.ListCount - 1)
            Case True

               StrSQLRequete = GStrSocietesFiltre

            Case False

               StrSQLRequete = DonneFiltreSocietes

         End Select

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLRequete

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ValideSocietes

               End Select

               On Error GoTo Err_ValideSocietes

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

              DoCmd.OpenForm "FrmFicheSocietes", acNormal, , StrSQLRequete, acFormEdit

         End Select
   End Select

Exit_ValideSocietes:

   Exit Sub

Err_ValideSocietes:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideSocietes"

   Resume Exit_ValideSocietes
End Sub

'******************************************************************************
'***  declaration D'Evenement Priver                                                           ***
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
   Dim StrSQLSocietesDebut As String
   Dim StrSQLSocietesFin As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_Form_Open

   StrSQLSocietesDebut = "SELECT DISTINCT [SelSocietesVilles].[SocNum], [SelSocietesVilles].[SocTypCode], [SelSocietesVilles].[SocSTA], " & _
                                       "[SelSocietesVilles].[SocRaison], [SelSocietesVilles].[SocAdh], [SelSocietesVilles].[SocAdresse], [SelSocietesVilles].[SocVilCodePostal], " & _
                                       "[SelSocietesVilles].[SocVilNom], [SelSocietesVilles].[SocTel1], [SelSocietesVilles].[SocTel2] FROM [SelSocietesVilles] "

   StrSQLSocietesFin = ";"

   Select Case GStrSocietesFiltre
      Case vbNullString

         LstLstSocietes.RowSource = StrSQLSocietesDebut & StrSQLSocietesFin

      Case Else

         DonneFiltreElement GStrSocietesFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

         Select Case StrFiltreTableRequeteNom
            Case "[SelSocietesVilles]"

               LstLstSocietes.RowSource = StrSQLSocietesDebut & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLSocietesFin

            Case Else

               LstLstSocietes.RowSource = StrSQLSocietesDebut & " INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesVilles]." & StrFiltreChampSource & "=" & _
                   StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLSocietesFin

         End Select

   End Select

   DoCmd.RunCommand acCmdWindowCascade

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

   ValideSocietes

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
Private Sub CmdSelectionnerTous_Click()
   Dim LngIndex As Long

   On Error GoTo Err_CmdSelectionnerTous_Click

   For LngIndex = 0 To LstLstSocietes.ListCount - 1

      LstLstSocietes.Selected(LngIndex) = Not LstLstSocietes.Selected(LngIndex)

   Next

   TxtLstSocietesCount.Requery

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
Private Sub CmdExporteMailing_Click()
   Dim StrSQLRequete As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_CmdExporteMailing_Click

   Select Case LstLstSocietes.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstSocietes.ItemsSelected.Count = (LstLstSocietes.ListCount - 1)
            Case True

               Select Case GStrSocietesFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles];"

                  Case Else

                     DonneFiltreElement GStrSocietesFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelSocietesVilles]"

                           StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesVilles]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] WHERE " & DonneFiltreSocietes & ";"

         End Select

         Select Case StrSQLRequete
            Case vbNullString

            Case Else

               DoCmd.OpenForm "FrmGestionMail", acNormal, , , acFormEdit, acWindowNormal, StrSQLRequete

         End Select

   End Select

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

   Select Case LstLstSocietes.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstSocietes.ItemsSelected.Count = (LstLstSocietes.ListCount - 1)
            Case True

               Select Case GStrSocietesFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles];"

                  Case Else

                     DonneFiltreElement GStrSocietesFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelSocietesVilles]"

                           StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelSocietesVilles]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelSocietesVilles].* FROM [SelSocietesVilles] WHERE " & DonneFiltreSocietes & ";"

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
Private Sub LstLstSocietes_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstSocietes_DblClick

   ValideSocietes

Exit_LstLstSocietes_DblClick:

   Exit Sub

Err_LstLstSocietes_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstSocietes_DblClick"

   Resume Exit_LstLstSocietes_DblClick
End Sub
