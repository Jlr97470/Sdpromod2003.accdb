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

         EtiValeurRecherche.Caption = "Valeur Recherchée : " & TbsRecherche.Pages(TbsRecherche.Value).Name

         Select Case TbsRecherche.Value + 1
            Case 1

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = "Recherche Code STA"

               EtiTypeRecherche5.Caption = "Recherche Raison Sociale"

               EtiTypeRecherche6.Caption = "Recherche Code Postal"

               EtiTypeRecherche7.Caption = "Recherche Telephone"

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = True

               EtiTypeRecherche5.Visible = True

               EtiTypeRecherche6.Visible = True

               EtiTypeRecherche7.Visible = True

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = True

               CbxTypeRecherche5.Visible = True

               CbxTypeRecherche6.Visible = True

               CbxTypeRecherche7.Visible = True

               CadTypeRecherche = 1

            Case 2

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = "Recherche Code Postal"

               EtiTypeRecherche5.Caption = "Recherche Telephone"

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = True

               EtiTypeRecherche5.Visible = True

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = True

               CbxTypeRecherche5.Visible = True

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False

               CadTypeRecherche = 1

            Case 3

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = vbNullString

               EtiTypeRecherche5.Caption = vbNullString

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = False

               EtiTypeRecherche5.Visible = False

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = False

               CbxTypeRecherche5.Visible = False

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False

               CadTypeRecherche = 1

            Case 4

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = "Recherche Numero De Serie"

               EtiTypeRecherche5.Caption = vbNullString

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = True

               EtiTypeRecherche5.Visible = False

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = True

               CbxTypeRecherche5.Visible = False

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False

               CadTypeRecherche = 4

         End Select

      Case Else

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

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
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' AND ParCode='INTERFACE'")
      Case "GESTION"

      Case Else

        ' TbsRecherche.Tabs.Remove (TbsRecherche.Tabs.Count)

       '  TbsRecherche.Tabs.Remove (TbsRecherche.Tabs.Count)

   End Select

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

   Resume Err_Form_Activate
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
   Dim IntEtat As Integer
   Dim QdfSocietesVilles As DAO.QueryDef
   Dim QdfContactsVilles As DAO.QueryDef
   Dim QdfCadeauxAchats As DAO.QueryDef
   Dim QdfChaudieresRemontees As DAO.QueryDef
   Dim RsSocietesVilles As DAO.Recordset
   Dim RsContactsVilles As DAO.Recordset
   Dim RsCadeauxAchats As DAO.Recordset
   Dim RsChaudieresRemontees As DAO.Recordset
   Dim FldChamp As DAO.Field

   On Error GoTo Err_CmdValider_Click

   Select Case CadTypeRecherche
      Case 2

         Select Case IsNumeric(TxtValeurRecherche)
            Case True

            Case False

               CadTypeRecherche = 1

         End Select

      Case 3

         Select Case IsDate(TxtValeurRecherche)
            Case True

            Case False

               CadTypeRecherche = 1

         End Select

   End Select

   Select Case TbsRecherche.Value + 1
      Case 1

         GStrSocietesFiltre = vbNullString

         Set QdfSocietesVilles = CurrentDb.QueryDefs("SelSocietesVilles")

         GStrSocietesFiltre = "SELECT [SelSocietesVilles].[SocNum] FROM [SelSocietesVilles] WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In QdfSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In QdfSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[" & FldChamp.Name & "]=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In QdfSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[" & FldChamp.Name & "]=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

            Case 4

               GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[SocSTA] LIKE '*" & TxtValeurRecherche & "*' OR "

            Case 5

               GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[SocRaison] LIKE '*" & TxtValeurRecherche & "*' OR "

            Case 6

               For Each FldChamp In QdfSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 5

                              GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

            Case 7

               For Each FldChamp In QdfSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 10

                              GStrSocietesFiltre = GStrSocietesFiltre & "[SelSocietesVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

         End Select

         GStrSocietesFiltre = Mid(GStrSocietesFiltre, 1, Len(GStrSocietesFiltre) - 4)

         Set RsSocietesVilles = CurrentDb.OpenRecordset(GStrSocietesFiltre & ";")

         GStrSocietesFiltre = "[SocNum] In (" & GStrSocietesFiltre & ";)"

         Select Case RsSocietesVilles.EOF
            Case True

            Case False

               RsSocietesVilles.MoveLast

         End Select

         Select Case RsSocietesVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmFicheSocietes", IntEtat

               Select Case IntEtat
                  Case ETATFORMULAIRE

                     Forms!FrmFicheSocietes.SetFocus

                     On Error Resume Next

                     DoCmd.ApplyFilter , "SocNum=" & RsSocietesVilles!SocNum

                     Select Case Err.Number
                        Case 0

                        Case 2501

                           Err.Clear

                        Case Else

                           GoTo Err_CmdValider_Click

                     End Select

                     On Error GoTo Err_CmdValider_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & RsSocietesVilles!SocNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstSocietes", acSavePrompt

               DoCmd.OpenForm "FrmLstSocietes", acNormal, , , acFormEdit, acWindowNormal, GStrSocietesFiltre

         End Select

         RsSocietesVilles.Close

      Case 2

         GStrContactsFiltre = vbNullString

         Set QdfContactsVilles = CurrentDb.QueryDefs("SelContactsVilles")

         GStrContactsFiltre = "SELECT [SelContactsVilles].[ConNum] FROM [SelContactsVilles] WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In QdfContactsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrContactsFiltre = GStrContactsFiltre & "[SelContactsVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In QdfContactsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrContactsFiltre = GStrContactsFiltre & "[SelContactsVilles].[" & FldChamp.Name & "]=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In QdfContactsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrContactsFiltre = GStrContactsFiltre & "[SelContactsVilles].[" & FldChamp.Name & "]=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

            Case 4

               For Each FldChamp In QdfContactsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 5

                              GStrContactsFiltre = GStrContactsFiltre & "[SelContactsVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

            Case 5

               For Each FldChamp In QdfContactsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 10

                              GStrContactsFiltre = GStrContactsFiltre & "[SelContactsVilles].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

         End Select

         GStrContactsFiltre = Mid(GStrContactsFiltre, 1, Len(GStrContactsFiltre) - 4)

         Set RsContactsVilles = CurrentDb.OpenRecordset(GStrContactsFiltre & ";")

         GStrContactsFiltre = "[ConNum] In (" & GStrContactsFiltre & ";)"

         Select Case RsContactsVilles.EOF
            Case True

            Case False

               RsContactsVilles.MoveLast

         End Select

         Select Case RsContactsVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmFicheContacts", IntEtat

               Select Case IntEtat
                  Case ETATFORMULAIRE

                     Forms!FrmFicheContacts.SetFocus

                     On Error Resume Next

                     DoCmd.ApplyFilter , "ConNum=" & RsContactsVilles!ConNum

                     Select Case Err.Number
                        Case 0

                        Case 2501

                           Err.Clear

                        Case Else

                           GoTo Err_CmdValider_Click

                     End Select

                     On Error GoTo Err_CmdValider_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & RsContactsVilles!ConNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstContacts", acSavePrompt

               DoCmd.OpenForm "FrmLstContacts", acNormal, , , acFormEdit, acWindowNormal, GStrContactsFiltre

         End Select

         RsContactsVilles.Close

      Case 3

         GStrCadeauxAchatsFiltre = vbNullString

         Set QdfCadeauxAchats = CurrentDb.QueryDefs("SelSocietesSTAContactsTAInsCadeauxAchats")

         GStrCadeauxAchatsFiltre = "SELECT [SelSocietesSTAContactsTAInsCadeauxAchats].[CadAchNum] FROM [SelSocietesSTAContactsTAInsCadeauxAchats] WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In QdfCadeauxAchats.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrCadeauxAchatsFiltre = GStrCadeauxAchatsFiltre & "[SelSocietesSTAContactsTAInsCadeauxAchats].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In QdfCadeauxAchats.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrCadeauxAchatsFiltre = GStrCadeauxAchatsFiltre & "[SelSocietesSTAContactsTAInsCadeauxAchats].[" & FldChamp.Name & "]=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In QdfCadeauxAchats.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrCadeauxAchatsFiltre = GStrCadeauxAchatsFiltre & "[SelSocietesSTAContactsTAInsCadeauxAchats].[" & FldChamp.Name & "]=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

         End Select

         GStrCadeauxAchatsFiltre = Mid(GStrCadeauxAchatsFiltre, 1, Len(GStrCadeauxAchatsFiltre) - 4)

         Set RsCadeauxAchats = CurrentDb.OpenRecordset(GStrCadeauxAchatsFiltre & ";")

         GStrCadeauxAchatsFiltre = "[CadAchNum] In (" & GStrCadeauxAchatsFiltre & ";)"

         Select Case RsCadeauxAchats.EOF
            Case True

            Case False

               RsCadeauxAchats.MoveLast

         End Select

         Select Case RsCadeauxAchats.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmSaisieCadeauxAchats", IntEtat

               Select Case IntEtat
                  Case ETATFORMULAIRE

                     Forms!FrmSaisieCadeauxAchats.SetFocus

                     On Error Resume Next

                     DoCmd.ApplyFilter , "CadAchNum=" & RsCadeauxAchats!CadAchNum

                     Select Case Err.Number
                        Case 0

                        Case 2501

                           Err.Clear

                        Case Else

                           GoTo Err_CmdValider_Click

                     End Select

                     On Error GoTo Err_CmdValider_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmSaisieCadeauxAchats", acNormal, , "CadAchNum=" & RsCadeauxAchats!CadAchNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstCadeauxAchats", acSavePrompt

               DoCmd.OpenForm "FrmLstCadeauxAchats", acNormal, , , acFormEdit, acWindowNormal, GStrCadeauxAchatsFiltre

         End Select

         RsCadeauxAchats.Close

      Case 4

         GStrChaudieresRemonteesFiltre = vbNullString

         Set QdfChaudieresRemontees = CurrentDb.QueryDefs("SelSocietesSTAContactsTAInsChaudieresRemontees")

         GStrChaudieresRemonteesFiltre = "SELECT [SelSocietesSTAContactsTAInsChaudieresRemontees].[ChdRemNum] FROM [SelSocietesSTAContactsTAInsChaudieresRemontees] WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In QdfChaudieresRemontees.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrChaudieresRemonteesFiltre = GStrChaudieresRemonteesFiltre & "[SelSocietesSTAContactsTAInsChaudieresRemontees].[" & FldChamp.Name & "] LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In QdfChaudieresRemontees.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrChaudieresRemonteesFiltre = GStrChaudieresRemonteesFiltre & "[SelSocietesSTAContactsTAInsChaudieresRemontees].[" & FldChamp.Name & "]=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In QdfChaudieresRemontees.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrChaudieresRemonteesFiltre = GStrChaudieresRemonteesFiltre & "[SelSocietesSTAContactsTAInsChaudieresRemontees].[" & FldChamp.Name & "]=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

            Case 4

               GStrChaudieresRemonteesFiltre = GStrChaudieresRemonteesFiltre & "[SelSocietesSTAContactsTAInsChaudieresRemontees].[ChdRemSerie] LIKE '*" & TxtValeurRecherche & "*' OR "

         End Select

         GStrChaudieresRemonteesFiltre = Mid(GStrChaudieresRemonteesFiltre, 1, Len(GStrChaudieresRemonteesFiltre) - 4)

         Set RsChaudieresRemontees = CurrentDb.OpenRecordset(GStrChaudieresRemonteesFiltre & ";")

         GStrChaudieresRemonteesFiltre = "[ChdRemNum] In (" & GStrChaudieresRemonteesFiltre & ";)"

         Select Case RsChaudieresRemontees.EOF
            Case True

            Case False

               RsChaudieresRemontees.MoveLast

         End Select

         Select Case RsChaudieresRemontees.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmSaisieChaudieresRemontees", IntEtat

               Select Case IntEtat
                  Case ETATFORMULAIRE

                     Forms!FrmSaisieChaudieresRemontees.SetFocus

                     On Error Resume Next

                     DoCmd.ApplyFilter , "ChdRemNum=" & RsChaudieresRemontees!ChdRemNum

                     Select Case Err.Number
                        Case 0

                        Case 2501

                           Err.Clear

                        Case Else

                           GoTo Err_CmdValider_Click

                     End Select

                     On Error GoTo Err_CmdValider_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmSaisieChaudieresRemontees", acNormal, , "ChdRemNum=" & RsChaudieresRemontees!ChdRemNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstChaudieresRemontees", acSavePrompt

               DoCmd.OpenForm "FrmLstChaudieresRemontees", acNormal, , , acFormEdit, acWindowNormal, GStrChaudieresRemonteesFiltre

         End Select

         RsChaudieresRemontees.Close

   End Select

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdValider_Click:

   Set QdfSocietesVilles = Nothing

   Set QdfContactsVilles = Nothing

   Set QdfCadeauxAchats = Nothing

   Set QdfChaudieresRemontees = Nothing

   Set RsSocietesVilles = Nothing

   Set RsContactsVilles = Nothing

   Set RsCadeauxAchats = Nothing

   Set RsChaudieresRemontees = Nothing

   Set FldChamp = Nothing

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
Private Sub TbsRecherche_Change()

   On Error GoTo Err_TbsRecherche_Change

   Rafraichir (RAFRAICHIRCHOIX)

Exit_TbsRecherche_Change:

   Exit Sub

Err_TbsRecherche_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TbsRecherche_Change"

   Resume Exit_TbsRecherche_Change
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtValeurRecherche_Change()

   On Error GoTo Err_TxtValeurRecherche_Change

   Select Case Len(TxtValeurRecherche.Text)
      Case Is > vbEmpty

         CmdValider.Enabled = True

      Case Else

         CmdValider.Enabled = False

   End Select

Exit_TxtValeurRecherche_Change:

   Exit Sub

Err_TxtValeurRecherche_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtValeurRecherche_Change"

   Resume Exit_TxtValeurRecherche_Change
End Sub
