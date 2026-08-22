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
Private Function DonneFiltreContacts() As String
   Dim StrContacts As String
   Dim VarItem As Variant

   On Error GoTo Err_DonneFiltreContacts

   DonneFiltreContacts = vbNullString

   Select Case LstLstContacts.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstContacts.ItemsSelected.Count
            Case 1

               StrContacts = "ConNum=" & LstLstContacts.Column(0)

            Case Else

               StrContacts = "ConNum In("

               For Each VarItem In LstLstContacts.ItemsSelected

                  StrContacts = StrContacts & LstLstContacts.Column(0, VarItem) & ", "

               Next

               StrContacts = Left(StrContacts, Len(StrContacts) - 2) & ")"

         End Select

         DonneFiltreContacts = StrContacts

   End Select

Exit_DonneFiltreContacts:

   Exit Function

Err_DonneFiltreContacts:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltreContacts"

   Resume Exit_DonneFiltreContacts
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

         LstLstContacts.Requery

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
Private Sub ValideContacts()
   Dim IntEtat As Integer
   Dim StrSQLRequete As String

   On Error GoTo Err_ValideContacts

   Select Case LstLstContacts.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstContacts.ItemsSelected.Count = (LstLstContacts.ListCount - 1)
            Case True

               StrSQLRequete = GStrContactsFiltre

            Case False

               StrSQLRequete = DonneFiltreContacts

         End Select

         DonneEtatFormulaire "FrmFicheContacts", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheContacts.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , StrSQLRequete

               Select Case Err.Number
                  Case 0

                  Case 2501

                     Err.Clear

                  Case Else

                     GoTo Err_ValideContacts

               End Select

               On Error GoTo Err_ValideContacts

               DoCmd.RunCommand acCmdWindowCascade

           Case Else

               DoCmd.OpenForm "FrmFicheContacts", acNormal, , StrSQLRequete, acFormEdit

      End Select

   End Select

Exit_ValideContacts:

   Exit Sub

Err_ValideContacts:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideContacts"

   Resume Exit_ValideContacts
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
   Dim StrSQLContactsDebut As String
   Dim StrSQLContactsFin As String
   Dim StrFiltreTableRequeteNom As String
   Dim StrFiltreChampSource As String
   Dim StrFiltreChampDestination As String
   Dim StrFiltreWhere As String

   On Error GoTo Err_Form_Open

   StrSQLContactsDebut = "SELECT DISTINCT [SelContactsVilles].[ConNum], [SelContactsVilles].[ConFonCode], [SelContactsVilles].[ConNom], " & _
                                       "[SelContactsVilles].[ConPrenom], [SelContactsVilles].[ConAdresse], [SelContactsVilles].[ConVilCodePostal], " & _
                                       "[SelContactsVilles].[ConVilNom], [SelContactsVilles].[ConTelB], [SelContactsVilles].[ConTelPort], [SelContactsVilles].[ConInscrit] " & _
                                       "FROM [SelContactsVilles] "

   StrSQLContactsFin = ";"

   Select Case GStrContactsFiltre
      Case vbNullString

         LstLstContacts.RowSource = StrSQLContactsDebut & StrSQLContactsFin

      Case Else

         DonneFiltreElement GStrContactsFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

         Select Case StrFiltreTableRequeteNom
            Case "[SelContactsVilles]"

               LstLstContacts.RowSource = StrSQLContactsDebut & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLContactsFin

            Case Else

               LstLstContacts.RowSource = StrSQLContactsDebut & " INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelContactsVilles]." & StrFiltreChampSource & "=" & _
                   StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & StrSQLContactsFin

         End Select

   End Select

   DoCmd.RunCommand acCmdWindowCascade

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         Me.MenuBar = "Menu Coeur De Pro 2003"

         Me.Toolbar = "Outil Coeur De Pro 2003"

      Case "CALL"

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

   ValideContacts

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

   OuvreRecherches 2

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

   Select Case LstLstContacts.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstContacts.ItemsSelected.Count = (LstLstContacts.ListCount - 1)
            Case True

               Select Case GStrContactsFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles];"

                  Case Else

                     DonneFiltreElement GStrContactsFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelContactsVilles]"

                           StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelContactsVilles]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] WHERE " & DonneFiltreContacts & ";"

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

   Select Case LstLstContacts.ItemsSelected.Count
      Case 0

      Case Else

         Select Case LstLstContacts.ItemsSelected.Count = (LstLstContacts.ListCount - 1)
            Case True

               Select Case GStrContactsFiltre
                  Case vbNullString

                     StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles];"

                  Case Else

                     DonneFiltreElement GStrContactsFiltre, StrFiltreTableRequeteNom, StrFiltreChampSource, StrFiltreChampDestination, StrFiltreWhere

                     Select Case StrFiltreTableRequeteNom
                        Case "[SelContactsVilles]"

                           StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] " & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                        Case Else

                           StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] INNER JOIN " & StrFiltreTableRequeteNom & " ON [SelContactsVilles]." & StrFiltreChampSource & "=" & _
                               StrFiltreTableRequeteNom & "." & StrFiltreChampDestination & IIf(StrFiltreWhere = vbNullString, vbNullString, " WHERE " & StrFiltreWhere) & ";"

                     End Select

               End Select

            Case False

               StrSQLRequete = "SELECT [SelContactsVilles].* FROM [SelContactsVilles] WHERE " & DonneFiltreContacts & ";"

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

   For LngIndex = 0 To LstLstContacts.ListCount - 1

      LstLstContacts.Selected(LngIndex) = Not LstLstContacts.Selected(LngIndex)

   Next

   TxtLstContactsCount.Requery

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
Private Sub LstLstContacts_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstContacts_DblClick

   ValideContacts

Exit_LstLstContacts_DblClick:

   Exit Sub

Err_LstLstContacts_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstContacts_DblClick"

   Resume Exit_LstLstContacts_DblClick
End Sub
