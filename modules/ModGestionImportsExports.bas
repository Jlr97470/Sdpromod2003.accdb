Attribute VB_Name = "ModGestionImportsExports"
'******************************************************************************
'***    DataSource Copyright                                                                       ***
'******************************************************************************
'***    MODULE:                                                                                          ***
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
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteExcelATCResultats() As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim VpbFeuille As Excel.VPageBreak
   Dim HpbFeuille As Excel.HPageBreak
   Dim RsContacts As DAO.Recordset
   Dim RsSocietesResultats As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSocietesResultatsDebut As String
   Dim StrSocietesResultatsFin As String
   Dim LngPosLig As Long
   Dim LngPosCol As Long
   Dim LngPosColChamp As Long
   Dim IntFeuilleNumero As Integer
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteExcelATCResultats

   StrSocietesResultatsDebut = "SELECT SocSTA, " & _
            " SocRaison, SocAdresse, " & _
            " SocVilCodePostal, SocVilNom, " & _
            " TotalCredit2004, Solde2002, " & _
            " Solde2003, Solde2004, " & _
            " MvtISO2002Sem1, MvtISO2002Sem2, " & _
            " MvtISO2003Sem1, MvtISO2003Sem2 " & _
            " FROM SelExtSocDRConATCSocSTAAdhConTAInsResISOSoldesDetailler " & _
            " WHERE ATCConNum="

   StrSocietesResultatsFin = " ORDER BY SocRaison;"

   ExporteExcelATCResultats = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   GTxtTitTraitements = "Mailing ATC Societes Resultats"

   GTxtFonTraitements = "ExporteExcelATCResultats"

   DoEvents

   Set RsContacts = CurrentDb.OpenRecordset("SELECT * FROM SelSocietesDRContactsATCInscrit ORDER BY ATCConNom ;")

   Select Case RsContacts.EOF
      Case True

      Case False

         RsContacts.MoveLast

         RsContacts.MoveFirst

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelATCResultats

         Set WbkClasseur = AppExcel.Workbooks.Add

         For IntFeuilleNumero = 3 To RsContacts.RecordCount - 1

            Set WstFeuille = WbkClasseur.Worksheets.Add

         Next

         IntFeuilleNumero = 1

         Do Until RsContacts.EOF = True

               GTxtObjTraitements = "Traitement En Cours : " & vbCrLf & RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

               DoEvents

               Set WstFeuille = WbkClasseur.Worksheets(IntFeuilleNumero)

               Set RsSocietesResultats = CurrentDb.OpenRecordset(StrSocietesResultatsDebut & RsContacts!ATCConNum & StrSocietesResultatsFin)

               With WstFeuille

                  .Name = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

                  .Rows.RowHeight = 14

                  .Cells(2, 1).Value = "CLUB COEUR DE PRO - Situation Points"

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesResultats.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesResultats.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesResultats.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesResultats.Fields.Count - 1)).Font.Bold = True

                  .Cells(3, 1).Value = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom & " - " & RsContacts!SocRaison

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesResultats.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesResultats.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesResultats.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesResultats.Fields.Count - 1)).Font.Bold = True

                  .Cells(4, 1).Value = Date

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesResultats.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesResultats.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesResultats.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesResultats.Fields.Count - 1)).Font.Bold = True

                  LngPosLig = 5

                  LngPosCol = 1

                  LngPosColChamp = vbEmpty

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Value = 1

                  GPbrTraitements.Max = RsSocietesResultats.Fields.Count

                  For Each FldChamp In RsSocietesResultats.Fields

                     Select Case LngPosColChamp
                        Case Is <= 4

                           On Error Resume Next

                           .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = IIf(FldChamp.Properties("Description").Value = vbNullString, FldChamp.Name, FldChamp.Properties("Description").Value)

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCResultats

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).MergeCells = True

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.LineStyle = 1

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.Weight = 3

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).HorizontalAlignment = -4108

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).VerticalAlignment = xlVAlignCenter

                        Case Else

                           Select Case LngPosColChamp
                              Case 5

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "POINTS ACQUIS"

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Font.Size = 8

                              Case 6

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "POINTS RESTANTS"

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Font.Size = 8

                              Case 9

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "BONUS ANIMATION ISO"

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Font.Size = 8

                           End Select

                           Select Case LngPosColChamp
                              Case 9

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "2002"

                              Case 11

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "2003"

                           End Select

                           On Error Resume Next

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = IIf(FldChamp.Properties("Description").Value = vbNullString, FldChamp.Name, FldChamp.Properties("Description").Value)

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCResultats

                           Select Case LngPosColChamp
                              Case 5

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders(8).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders(9).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders(7).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders(10).Weight = 3

                              Case 6

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 2 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).Borders(8).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).Borders(9).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).Borders(7).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 2 + LngPosColChamp)).Borders(10).Weight = 3

                              Case 9

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 3 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).Borders(8).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).Borders(9).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).Borders(7).Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 3 + LngPosColChamp)).Borders(10).Weight = 3

                           End Select

                           Select Case LngPosColChamp

                              Case 9, 11

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).Borders(8).Weight = 3

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).Borders(9).Weight = 3

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).Borders(7).Weight = 3

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + 1 + LngPosColChamp)).Borders(10).Weight = 3

                           End Select

                     End Select

                     LngPosColChamp = LngPosColChamp + 1

                     GTxtObjTraitements = "Exportation Des Libelles Des Champs : " & FldChamp.Name

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  GTxtObjTraitements = "Exportation Des Données"

                  DoEvents

                  .Cells(LngPosLig + 4, LngPosCol).CopyFromRecordset RsSocietesResultats

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesResultats.RecordCount - 1, LngPosCol + RsSocietesResultats.Fields.Count - 1)).Borders.LineStyle = 1

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesResultats.RecordCount - 1, LngPosCol + RsSocietesResultats.Fields.Count - 1)).Borders(8).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesResultats.RecordCount - 1, LngPosCol + RsSocietesResultats.Fields.Count - 1)).Borders(9).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesResultats.RecordCount - 1, LngPosCol + RsSocietesResultats.Fields.Count - 1)).Borders(7).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesResultats.RecordCount - 1, LngPosCol + RsSocietesResultats.Fields.Count - 1)).Borders(10).Weight = 3

                  LngPosColChamp = vbEmpty

                  GTxtObjTraitements = "Mise En Forme De La Feuille Excel"

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Max = RsSocietesResultats.Fields.Count

                  GPbrTraitements.Value = 1

                  DoEvents

                  For Each FldChamp In RsSocietesResultats.Fields

                     .Columns(LngPosCol + LngPosColChamp).AutoFit

                     LngPosColChamp = LngPosColChamp + 1

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  .PageSetup.Orientation = 2

                  .PageSetup.TopMargin = 20

                  .PageSetup.LeftMargin = 20

                  .PageSetup.BottomMargin = 20

                  .PageSetup.RightMargin = 20

                  .PageSetup.PrintTitleRows = .Range(.Cells(1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Address

                  .PageSetup.Zoom = False

                  .PageSetup.FitToPagesWide = 1

                  .PageSetup.FitToPagesTall = 10

               End With

               IntFeuilleNumero = IntFeuilleNumero + 1

               RsContacts.MoveNext

         Loop

         RsSocietesResultats.Close

         AppExcel.Visible = True

   End Select

   RsContacts.Close

Exit_ExporteExcelATCResultats:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set RsContacts = Nothing

   Set RsSocietesResultats = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Exit Function

Err_ExporteExcelATCResultats:

   ExporteExcelATCResultats = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelATCResultats"

   Resume Exit_ExporteExcelATCResultats
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteExcelATCChaudieres() As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim VpbFeuille As Excel.VPageBreak
   Dim HpbFeuille As Excel.HPageBreak
   Dim RsContacts As DAO.Recordset
   Dim RsSocietesChaudieres As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSocietesChaudieresDebut As String
   Dim StrSocietesChaudieresFin As String
   Dim LngPosLig As Long
   Dim LngPosCol As Long
   Dim LngPosColChamp As Long
   Dim IntFeuilleNumero As Integer
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteExcelATCChaudieres

   StrSocietesChaudieresDebut = "SELECT SocSTA, SocRaison " & _
                                                  ", SocResNbISO2003S1, SocResNbISO2003S2, SocResNbTHEMA2003S1, SocResNbTHEMA2003S2, SocResNbChd2003S1, SocResNbChd2003S2 " & _
                                                  ", SocResNbISO2004S1, SocResNbISO2004S2, SocResNbTHEMA2004S1, SocResNbTHEMA2004S2, SocResNbChd2004S1, SocResNbChd2004S2 " & _
                                                  ", SocResPourISO2004S1, SocResPourISO2004S2 " & _
                                                  " FROM SelExtSocDRConATCSocSTAAdhConTAInsResSemestresDetailler WHERE ATCConNum="

   StrSocietesChaudieresFin = " ORDER BY SocRaison;"

   ExporteExcelATCChaudieres = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   GTxtTitTraitements = "Mailing ATC Societes Chaudieres"

   GTxtFonTraitements = "ExporteExcelATCChaudieres"

   DoEvents

   Set RsContacts = CurrentDb.OpenRecordset("SELECT * FROM SelSocietesDRContactsATCInscrit ORDER BY ATCConNom ;")

   Select Case RsContacts.EOF
      Case True

      Case False

         RsContacts.MoveLast

         RsContacts.MoveFirst

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelATCChaudieres

         Set WbkClasseur = AppExcel.Workbooks.Add

         For IntFeuilleNumero = 3 To RsContacts.RecordCount - 1

            Set WstFeuille = WbkClasseur.Worksheets.Add

         Next

         IntFeuilleNumero = 1

         Do Until RsContacts.EOF = True

               GTxtObjTraitements = "Traitement En Cours : " & vbCrLf & RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

               DoEvents

               Set WstFeuille = WbkClasseur.Worksheets(IntFeuilleNumero)

               Set RsSocietesChaudieres = CurrentDb.OpenRecordset(StrSocietesChaudieresDebut & RsContacts!ATCConNum & StrSocietesChaudieresFin)

               With WstFeuille

                  .Name = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

                  .Rows.RowHeight = 14

                  .Cells(2, 1).Value = "CLUB COEUR DE PRO - Situation Des Chaudieres"

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  .Cells(3, 1).Value = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom & " - " & RsContacts!SocRaison

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  .Cells(4, 1).Value = Date

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(4, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  LngPosLig = 5

                  LngPosCol = 1

                  LngPosColChamp = vbEmpty

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Value = 1

                  GPbrTraitements.Max = RsSocietesChaudieres.Fields.Count

                  For Each FldChamp In RsSocietesChaudieres.Fields

                     Select Case LngPosColChamp
                        Case Is <= 1

                           On Error Resume Next

                           .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCChaudieres

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).MergeCells = True

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.LineStyle = 1

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.Weight = 3

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).HorizontalAlignment = -4108

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).VerticalAlignment = xlVAlignCenter

                        Case Else

                           Select Case LngPosColChamp
                              Case 2

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "2003"

                              Case 8

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "2004"

                           End Select

                           Select Case LngPosColChamp
                              Case 2, 8

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "ISO"

                              Case 4, 10

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "THEMA"

                              Case 6, 12

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "TOTAL"

                              Case 14

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "% ISO"

                           End Select

                           On Error Resume Next

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value = vbNullString

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCChaudieres

                           Select Case LngPosColChamp
                              Case 2

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 5 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 5 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 5 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 5 + LngPosColChamp)).HorizontalAlignment = -4108

                              Case 8

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).HorizontalAlignment = -4108

                           End Select

                           Select Case LngPosColChamp
                              Case 2, 4, 6, 8, 10, 12, 14

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Font.Size = 8

                              Case Else

                           End Select

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Borders.LineStyle = 1

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Borders.Weight = 3

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).HorizontalAlignment = -4108

                     End Select

                     LngPosColChamp = LngPosColChamp + 1

                     GTxtObjTraitements = "Exportation Des Libelles Des Champs : " & FldChamp.Name

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  GTxtObjTraitements = "Exportation Des Données"

                  DoEvents

                  .Cells(LngPosLig + 4, LngPosCol).CopyFromRecordset RsSocietesChaudieres

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders.LineStyle = 1

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(8).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(9).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(7).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(10).Weight = 3

                  LngPosColChamp = vbEmpty

                  GTxtObjTraitements = "Mise En Forme De La Feuille Excel"

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Max = RsSocietesChaudieres.Fields.Count

                  GPbrTraitements.Value = 1

                  DoEvents

                  For Each FldChamp In RsSocietesChaudieres.Fields

                     Select Case LngPosColChamp
                        Case 14

                           .Range(.Cells(LngPosLig + 4, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "0.00%"

                        Case Else

                     End Select

                     .Columns(LngPosCol + LngPosColChamp).AutoFit

                     LngPosColChamp = LngPosColChamp + 1

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  .PageSetup.Orientation = 2

                  .PageSetup.TopMargin = 20

                  .PageSetup.LeftMargin = 20

                  .PageSetup.BottomMargin = 20

                  .PageSetup.RightMargin = 20

                  .PageSetup.PrintTitleRows = .Range(.Cells(1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Address

                  .PageSetup.Zoom = False

                  .PageSetup.FitToPagesWide = 1

                  .PageSetup.FitToPagesTall = 10

               End With

               IntFeuilleNumero = IntFeuilleNumero + 1

               RsContacts.MoveNext

         Loop

         RsSocietesChaudieres.Close

         AppExcel.Visible = True

   End Select

   RsContacts.Close

Exit_ExporteExcelATCChaudieres:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set RsContacts = Nothing

   Set RsSocietesChaudieres = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Exit Function

Err_ExporteExcelATCChaudieres:

   ExporteExcelATCChaudieres = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelATCChaudieres"

   Resume Exit_ExporteExcelATCChaudieres
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteExcelATCChaudieresSPLIT() As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim VpbFeuille As Excel.VPageBreak
   Dim HpbFeuille As Excel.HPageBreak
   Dim RsContacts As DAO.Recordset
   Dim RsSocietesChaudieres As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSocietesChaudieresDebut As String
   Dim StrSocietesChaudieresFin As String
   Dim LngPosLig As Long
   Dim LngPosCol As Long
   Dim LngPosColChamp As Long
   Dim IntFeuilleNumero As Integer
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteExcelATCChaudieresSPLIT

   StrSocietesChaudieresDebut = "SELECT SocSTA, SocRaison " & _
                                                  ", SocResNbISOSTD2003S1, SocResNbISOSTD2003S2, SocResNbISOSPLIT2003S1, SocResNbISOSPLIT2003S2, SocResNbTHEMA2003S1, SocResNbTHEMA2003S2, SocResNbChd2003S1, SocResNbChd2003S2 " & _
                                                  ", SocResNbISOSTD2004S1, SocResNbISOSTD2004S2, SocResNbISOSPLIT2004S1, SocResNbISOSPLIT2004S2, SocResNbTHEMA2004S1, SocResNbTHEMA2004S2, SocResNbChd2004S1, SocResNbChd2004S2 " & _
                                                  ", SocResPourISO2004S1, SocResPourISO2004S2 " & _
                                                  " FROM SelExtSocDRConATCSocSTAAdhConTAInsResSemestresDetaillerSPLIT WHERE ATCConNum="

   StrSocietesChaudieresFin = " ORDER BY SocRaison;"

   ExporteExcelATCChaudieresSPLIT = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   GTxtTitTraitements = "Mailing ATC Societes Chaudieres"

   GTxtFonTraitements = "ExporteExcelATCChaudieresSPLIT"

   DoEvents

   Set RsContacts = CurrentDb.OpenRecordset("SELECT * FROM SelSocietesDRContactsATCInscrit ORDER BY ATCConNom ;")

   Select Case RsContacts.EOF
      Case True

      Case False

         RsContacts.MoveLast

         RsContacts.MoveFirst

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelATCChaudieresSPLIT

         Set WbkClasseur = AppExcel.Workbooks.Add

         For IntFeuilleNumero = 3 To RsContacts.RecordCount - 1

            Set WstFeuille = WbkClasseur.Worksheets.Add

         Next

         IntFeuilleNumero = 1

         Do Until RsContacts.EOF = True

               GTxtObjTraitements = "Traitement En Cours : " & vbCrLf & RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

               DoEvents

               Set WstFeuille = WbkClasseur.Worksheets(IntFeuilleNumero)

               Set RsSocietesChaudieres = CurrentDb.OpenRecordset(StrSocietesChaudieresDebut & RsContacts!ATCConNum & StrSocietesChaudieresFin)

               With WstFeuille

                  .Name = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom

                  .Rows.RowHeight = 14

                  .Cells(2, 1).Value = "CLUB COEUR DE PRO - Situation Des Chaudieres"

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(2, 1), .Cells(2, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  .Cells(3, 1).Value = RsContacts!ATCConNom & " " & RsContacts!ATCConPrenom & " - " & RsContacts!SocRaison

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(3, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  .Cells(4, 1).Value = Date

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).MergeCells = True

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).HorizontalAlignment = -4108

                  .Range(.Cells(4, 1), .Cells(4, 1 + RsSocietesChaudieres.Fields.Count - 1)).RowHeight = 20

                  .Range(.Cells(4, 1), .Cells(3, 1 + RsSocietesChaudieres.Fields.Count - 1)).Font.Bold = True

                  LngPosLig = 5

                  LngPosCol = 1

                  LngPosColChamp = vbEmpty

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Value = 1

                  GPbrTraitements.Max = RsSocietesChaudieres.Fields.Count

                  For Each FldChamp In RsSocietesChaudieres.Fields

                     Select Case LngPosColChamp
                        Case Is <= 1

                           On Error Resume Next

                           .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCChaudieresSPLIT

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).MergeCells = True

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.LineStyle = 1

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Borders.Weight = 3

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).HorizontalAlignment = -4108

                           .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).VerticalAlignment = xlVAlignCenter

                        Case Else

                           Select Case LngPosColChamp
                              Case 2

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "2003"

                              Case 10

                                 .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = "2004"

                           End Select

                           Select Case LngPosColChamp
                              Case 2, 10

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "ISO"

                              Case 4, 12

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "ISOSPLIT"

                              Case 6, 14

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "THEMA"

                              Case 8, 16

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "TOTAL"

                              Case 18

                                 .Cells(LngPosLig + 2, LngPosCol + LngPosColChamp).Value = "% ISO"

                           End Select

                           On Error Resume Next

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Value = FldChamp.Name

                           End Select

                           On Error GoTo Err_ExporteExcelATCChaudieresSPLIT

                           Select Case LngPosColChamp
                              Case 2

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 7 + LngPosColChamp)).HorizontalAlignment = -4108

                              Case 10

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 9 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 9 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 9 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 1, LngPosCol + 9 + LngPosColChamp)).HorizontalAlignment = -4108

                           End Select

                           Select Case LngPosColChamp
                              Case 2, 4, 6, 8, 10, 12, 14, 16, 18

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).MergeCells = True

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Borders.LineStyle = 1

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Borders.Weight = 3

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).HorizontalAlignment = -4108

                                 .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2, LngPosCol + 1 + LngPosColChamp)).Font.Size = 8

                              Case Else

                           End Select

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Borders.LineStyle = 1

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).Borders.Weight = 3

                           .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp).HorizontalAlignment = -4108

                     End Select

                     LngPosColChamp = LngPosColChamp + 1

                     GTxtObjTraitements = "Exportation Des Libelles Des Champs : " & FldChamp.Name

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  GTxtObjTraitements = "Exportation Des Données"

                  DoEvents

                  .Cells(LngPosLig + 4, LngPosCol).CopyFromRecordset RsSocietesChaudieres

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders.LineStyle = 1

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(8).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(9).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(7).Weight = 3

                  .Range(.Cells(LngPosLig + 4, LngPosCol), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + RsSocietesChaudieres.Fields.Count - 1)).Borders(10).Weight = 3

                  LngPosColChamp = vbEmpty

                  GTxtObjTraitements = "Mise En Forme De La Feuille Excel"

                  GPbrTraitements.Min = 0

                  GPbrTraitements.Max = RsSocietesChaudieres.Fields.Count

                  GPbrTraitements.Value = 1

                  DoEvents

                  For Each FldChamp In RsSocietesChaudieres.Fields

                     Select Case LngPosColChamp
                        Case 18, 19

                           .Range(.Cells(LngPosLig + 4, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 4 + RsSocietesChaudieres.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "0.00%"

                        Case Else

                     End Select

                     .Columns(LngPosCol + LngPosColChamp).AutoFit

                     LngPosColChamp = LngPosColChamp + 1

                     GPbrTraitements.Value = LngPosColChamp

                     DoEvents

                  Next

                  .PageSetup.Orientation = 2

                  .PageSetup.TopMargin = 20

                  .PageSetup.LeftMargin = 20

                  .PageSetup.BottomMargin = 20

                  .PageSetup.RightMargin = 20

                  .PageSetup.PrintTitleRows = .Range(.Cells(1, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 3, LngPosCol + LngPosColChamp)).Address

                  .PageSetup.Zoom = False

                  .PageSetup.FitToPagesWide = 1

                  .PageSetup.FitToPagesTall = 10

               End With

               IntFeuilleNumero = IntFeuilleNumero + 1

               RsContacts.MoveNext

         Loop

         RsSocietesChaudieres.Close

         AppExcel.Visible = True

   End Select

   RsContacts.Close

Exit_ExporteExcelATCChaudieresSPLIT:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set RsContacts = Nothing

   Set RsSocietesChaudieres = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Exit Function

Err_ExporteExcelATCChaudieresSPLIT:

   ExporteExcelATCChaudieresSPLIT = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelATCChaudieresSPLIT"

   Resume Exit_ExporteExcelATCChaudieresSPLIT
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteExcelRequete(ByVal StrSQL As String, Optional StrFichier As String) As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim RsRequete As DAO.Recordset
   Dim StrCheminFichierExcel As String
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteExcelRequete

   ExporteExcelRequete = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   GTxtFonTraitements = "ExporteExcelRequete"

   GTxtObjTraitements = StrSQL

   DoEvents

   Set RsRequete = CurrentDb.OpenRecordset(StrSQL)

   Select Case RsRequete.RecordCount
      Case vbEmpty

      Case Else

         On Error Resume Next

         GTxtTitTraitements = RsRequete.Properties("Description").Value

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               GTxtTitTraitements = "Exportation Excel"

         End Select

         On Error GoTo Err_ExporteExcelRequete

         DoEvents

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelRequete

         Select Case StrFichier
            Case vbNullString

               Set WbkClasseur = AppExcel.Workbooks.Add

            Case Else

               StrCheminFichierExcel = Nz(DLookup("FicValeur", "SelFichiersDetailler", "FicType='FICXLS' AND FicCode LIKE '*=" & StrFichier & "' AND FicValide=True"))

               Set WbkClasseur = AppExcel.Workbooks.Open(StrCheminFichierExcel & "\" & StrFichier)

         End Select

         Set WstFeuille = WbkClasseur.Worksheets(1)

         On Error Resume Next

         WstFeuille.Name = Left(RsRequete.Properties("Description").Value, 31)

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               WstFeuille.Name = "Exportation Excel"

         End Select

         WstFeuille.Rows.RowHeight = 14

         ExporteExcelRequeteDonner WstFeuille, RsRequete, 1, 1

         AppExcel.Visible = True

   End Select

   RsRequete.Close

Exit_ExporteExcelRequete:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set RsRequete = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Set AppExcel = Nothing

   Exit Function

Err_ExporteExcelRequete:

   ExporteExcelRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelRequete"

   Resume Exit_ExporteExcelRequete
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteExcelRequeteDonner(ByRef WstFeuille As Worksheet, ByRef RsRequete As DAO.Recordset, ByVal LngPosCol As Long, ByVal LngPosLig As Long) As Boolean
   Dim FldChamp As DAO.Field
   Dim HypChamp As Excel.Hyperlinks
   Dim LngPosColChamp As Long
   Dim LngPosLigChamp As Long
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteExcelRequeteDonner

   ExporteExcelRequeteDonner = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   GTxtFonTraitements = "ExporteExcelRequeteDonner"

   GTxtObjTraitements = RsRequete.Name

   DoEvents

   With WstFeuille

         On Error Resume Next

         .Cells(LngPosLig, LngPosCol).Value = RsRequete.Properties("Description").Value

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               .Cells(LngPosLig, LngPosCol).Value = "Exportation Excel"

         End Select

         On Error GoTo Err_ExporteExcelRequeteDonner

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).MergeCells = True

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).HorizontalAlignment = -4108

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Interior.Color = RGB(100, 100, 100)

         LngPosColChamp = vbEmpty

         GPbrTraitements.Min = 0

         GPbrTraitements.Value = 1

         GPbrTraitements.Max = RsRequete.Fields.Count

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.LineStyle = 1

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.Weight = 3

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).HorizontalAlignment = -4108

            LngPosColChamp = LngPosColChamp + 1

            GTxtObjTraitements = "Exportation Des Libelles Des Champs : " & FldChamp.Name

            GPbrTraitements.Value = LngPosColChamp

            DoEvents

         Next

         GTxtObjTraitements = "Exportation Des Données"

         DoEvents

         .Cells(LngPosLig + 2, LngPosCol).CopyFromRecordset RsRequete

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(8).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(9).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(7).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(10).Weight = 3

         LngPosColChamp = vbEmpty

         GTxtObjTraitements = "Mise En Forme De La Feuille Excel"

         GPbrTraitements.Min = 0

         GPbrTraitements.Max = RsRequete.Fields.Count

         GPbrTraitements.Value = 1

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            Select Case FldChamp.Properties("Format")
               Case vbNullString

               Case "Percent"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "0.00%"

               Case "Short Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY"

               Case "General Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY HH:MM"

               Case Else

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            Select Case InStr(1, FldChamp.Name, "EMail")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "mailto:" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            Select Case InStr(1, FldChamp.Name, "Web")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "http://" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            .Columns(LngPosCol + LngPosColChamp).AutoFit

            LngPosColChamp = LngPosColChamp + 1

            GPbrTraitements.Value = LngPosColChamp

            DoEvents

         Next

         .PageSetup.Orientation = 2

         .PageSetup.Zoom = False

         .PageSetup.FitToPagesWide = 1

         .PageSetup.FitToPagesTall = 10

   End With

Exit_ExporteExcelRequeteDonner:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set HypChamp = Nothing

   Exit Function

Err_ExporteExcelRequeteDonner:

   ExporteExcelRequeteDonner = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelRequeteDonner"

   Resume Exit_ExporteExcelRequeteDonner
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function FusionWordNouveauxAdherents(ByVal StrFichier As String) As Boolean
   Dim StrDateDebut As String

   On Error GoTo Err_FusionWordNouveauxAdherents

   Do

      StrDateDebut = InputBox("A Partir De Quelle Date Voulez-Vous Lancez L'Edition Des Nouveaux Adherents ?", "FusionWordNouveauxAdherents")

      Select Case StrDateDebut
         Case vbNullString

            MsgBox "Aucunne Date Saisie !", vbInformation

            GoTo Exit_FusionWordNouveauxAdherents

         Case Else

            Select Case IsDate(StrDateDebut)
               Case True

               Case False

                  MsgBox "La Date Saisie Est Incorrect !", vbInformation

            End Select

      End Select

   Loop Until IsDate(StrDateDebut) = True

   FusionWordRequete "SELECT [SelSocDRConATCSocSTAAdhConTAInsDateAdh].* FROM [SelSocDRConATCSocSTAAdhConTAInsDateAdh] WHERE [SocDateAdh] >= CDate('" & StrDateDebut & "');", StrFichier

Exit_FusionWordNouveauxAdherents:

   Exit Function

Err_FusionWordNouveauxAdherents:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "FusionWordNouveauxAdherents"

   Resume Exit_FusionWordNouveauxAdherents
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function FusionWordRequete(ByVal StrSQL As String, ByVal StrFichier As String) As Boolean
   Dim WshShellSystem As New IWshRuntimeLibrary.WshShell
   Dim AppWord As Word.Application
   Dim DocWord As Word.Document
   Dim MerWord As Word.MailMerge
   Dim TdfRequete As DAO.QueryDef
   Dim TdsRequete As DAO.QueryDefs
   Dim StrRequeteNom As String
   Dim StrCheminFichierAppWord As String
   Dim StrCheminFichierWord As String

   On Error GoTo Err_FusionWordRequete

   Set AppWord = CreateObject("Word.Application")

   For Each DocWord In AppWord.Documents

      DocWord.Close 0

   Next

   AppWord.Visible = True

   StrCheminFichierWord = Nz(DLookup("FicValeur", "SelFichiersDetailler", "FicType='FICDOC' AND FicCode LIKE '*=" & StrFichier & "' AND FicValide=True"))

   Select Case StrCheminFichierWord
      Case vbNullString

      Case Else

         On Error Resume Next

         AppWord.Documents.Open (StrCheminFichierWord & "\" & StrFichier)

         Select Case Err.Number
            Case 0

            Case 5981

               Err.Clear

            Case Else

               GoTo Err_FusionWordRequete

         End Select

         On Error GoTo Err_FusionWordRequete

         Select Case Left(StrSQL, 6)
            Case "SELECT"

            Case Else

               StrSQL = "SELECT " & StrSQL & ".* FROM " & StrSQL & ";"

         End Select

         Set DocWord = AppWord.Documents(StrFichier)

         Set MerWord = DocWord.MailMerge

         Set TdsRequete = CurrentDb.QueryDefs

         StrRequeteNom = "SelFusionWord" & Int(Rnd() * 10) & Int(Rnd() * 10)

         Set TdfRequete = CurrentDb.CreateQueryDef(StrRequeteNom, StrSQL)

         TdsRequete.Refresh

         MerWord.OpenDataSource Name:=CurrentDb.Name, Connection:="DSN=MS Access Databases;DBQ=" & CurrentDb.Name & ";", SQLStatement:="SELECT " & StrRequeteNom & ".* FROM " & StrRequeteNom & ";"

         MerWord.Destination = 0

         DocWord.MailMerge.Execute

         DocWord.Close 0

         Set TdfRequete = Nothing

         TdsRequete.Delete (StrRequeteNom)

   End Select

Exit_FusionWordRequete:

   Set WshShellSystem = Nothing

   Set MerWord = Nothing

   Set DocWord = Nothing

   Set AppWord = Nothing

   On Error Resume Next

   Set TdfRequete = Nothing

   TdsRequete.Delete (StrRequeteNom)

   Set TdsRequete = Nothing

   Exit Function

Err_FusionWordRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "FusionWordRequete"

   Resume Exit_FusionWordRequete
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ExporteMailingRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_ExporteMailingRequete

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='MAILING'")
      Case "JAVAMAIL"

         ExporteMailingRequete = MailingJavaMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

      Case "VBSENDMAIL"

         ExporteMailingRequete = MailingSendMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

      Case "VBOUTLOOKMAIL"

         ExporteMailingRequete = MailingSendMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

      Case "VBMAPIMAIL"

         ExporteMailingRequete = MailingSendMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

      Case "VBLOTUSMAIL"

         ExporteMailingRequete = MailingSendMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

   End Select

Exit_ExporteMailingRequete:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Exit Function

Err_ExporteMailingRequete:

   ExporteMailingRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteMailingRequete"

   Resume Exit_ExporteMailingRequete
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function MailingJavaMailRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean
   Dim FsoSystem As Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim JavaMsg As Object 'jmail.Message
   Dim HtmObjectDocument As MSHTML.HTMLDocument
   Dim HtmDocument As MSHTML.HTMLDocument
   Dim RsRequete As DAO.Recordset
   Dim RsFichier As DAO.Recordset
   Dim StrEmailFrom As String
   Dim StrAttacheID As String
   Dim StrAttacheFichier As String
   Dim StrDocument As String
   Dim StrCheminFichierBody As String
   Dim StrChampNom As String
   Dim IntNb As Integer
   Dim IntPosDebut As Integer
   Dim IntPosFin As Integer
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_MailingJavaMailRequete

   MailingJavaMailRequete = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acWindowNormal

   GTxtTitTraitements = "Mailing : " & StrSujet

   GTxtFonTraitements = "MailingJavaMailRequete"

   GTxtObjTraitements = StrSQLRequete

   DoEvents

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicCode2='" & StrFichierBody & "' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         StrCheminFichierBody = RsFichier!FicValeur & "\" & RsFichier!FicCode2

         RsFichier.Close

         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case RsRequete.EOF
            Case True

            Case False

               Set HtmObjectDocument = New MSHTML.HTMLDocument

               Set HtmDocument = HtmObjectDocument.createDocumentFromUrl(StrCheminFichierBody, vbNullString)

               Do Until HtmDocument.readyState = "complete"

                  DoEvents

               Loop

           '    Set JavaMsg = New jmail.Message

               GTxtObjTraitements = "CONNEXION " & vbCrLf

               JavaMsg.from = StrEmailExpediteur

               JavaMsg.Subject = StrSujet

               JavaMsg.AppendText ("Apparemment, votre logiciel d'email ne supporte pas le format HTML.")

               JavaMsg.Priority = 3

               JavaMsg.ReturnReceipt = True

               JavaMsg.Logging = True

               JavaMsg.MailServerUserName = vbNullString

               JavaMsg.MailServerPassWord = vbNullString

               Set FsoSystem = New Scripting.FileSystemObject

               Select Case HtmDocument.images.Length
                  Case 0

                  Case Else

                     For IntNb = 0 To HtmDocument.images.Length - 1

                        StrAttacheFichier = RemplaceChr(RemplaceChr(RemplaceChr(HtmDocument.images(IntNb).src, "file:///", vbNullString), "%20", " "), "/", "\")

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttacheID = JavaMsg.AddAttachment(StrAttacheFichier, True)

                              Set FilFichier = FsoSystem.GetFile(StrAttacheFichier)

                              HtmDocument.images(IntNb).src = "cid:" & StrAttacheID

                           Case False

                        End Select

                     Next

               End Select

               On Error Resume Next

               Select Case StrFichierAttache(LBound(StrFichierAttache))
                  Case vbNullString

                  Case Else

                     For IntNb = LBound(StrFichierAttache) To UBound(StrFichierAttache)

                        StrAttacheFichier = StrFichierAttache(IntNb)

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              JavaMsg.AddCustomAttachment StrAttacheFichier, Right(StrAttacheFichier, Len(StrAttacheFichier) - InStrRev(StrAttacheFichier, "\")), False

                           Case False

                        End Select

                     Next

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

               End Select

               On Error GoTo Err_MailingJavaMailRequete

               RsRequete.MoveLast

               RsRequete.MoveFirst

               GPbrTraitements.Min = 0

               GPbrTraitements.Value = 0

               GPbrTraitements.Max = RsRequete.RecordCount

               Do Until RsRequete.EOF = True

                  JavaMsg.ClearRecipients

                  Select Case BlnSimulation
                        Case True

                        Case False

                           JavaMsg.AddRecipientBCC StrEmailCopieCacher

                  End Select

                  Select Case IsNull(RsRequete(StrChampDestination))
                     Case False

                        Select Case InStr(1, RsRequete(StrChampDestination), "mailto:")
                           Case Is <= 0

                              Select Case InStr(1, RsRequete(StrChampDestination), "#")
                                 Case Is <= 0

                                    StrEmailFrom = RsRequete(StrChampDestination)

                                 Case Else

                                    StrEmailFrom = Left(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "#") - 1)

                              End Select

                           Case Else

                              StrEmailFrom = RemplaceChr(Mid(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "mailto:") + 7), "#", vbNullString)

                        End Select

                        Select Case BlnSimulation
                           Case True

                              JavaMsg.AddRecipient StrEmailExpediteur

                              JavaMsg.Subject = StrSujet & " (" & StrEmailFrom & ")"

                           Case False

                              JavaMsg.AddRecipient StrEmailFrom

                        End Select

                        StrDocument = "<HTML>" & vbCrLf & HtmDocument.documentElement.innerHTML & vbCrLf & "</HTML>"

                        IntPosDebut = 1

                        Do Until InStr(IntPosDebut, StrDocument, "<D VALUE=") <= 0

                           IntPosDebut = InStr(IntPosDebut, StrDocument, "<D VALUE=")

                           IntPosFin = InStr(IntPosDebut, StrDocument, ">")

                           StrChampNom = RemplaceChr(Mid(StrDocument, IntPosDebut + 10, IntPosFin - IntPosDebut - 11), Chr$(34), "")

                           IntPosFin = InStr(IntPosDebut, StrDocument, "</D>")

                           On Error Resume Next

                           StrDocument = Left(StrDocument, IntPosDebut - 1) & RsRequete(StrChampNom) & Mid(StrDocument, IntPosFin + 4)

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 IntPosDebut = IntPosFin

                           End Select

                           On Error GoTo Err_MailingJavaMailRequete

                        Loop

                        JavaMsg.HTMLBody = StrDocument

                     Case True

                  End Select

                  GPbrTraitements.Value = GPbrTraitements.Value + 1

                 Select Case BlnSimulation
                     Case True

                        GTxtObjTraitements = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                        GTxtNotesTraitements = GTxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                     Case False

                        GTxtObjTraitements = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                        GTxtNotesTraitements = GTxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                  End Select

                  Select Case JavaMsg.Send("smtp.9online.fr")
                     Case True

                        GTxtNotesTraitements = GTxtNotesTraitements & JavaMsg.Log & vbCrLf

                        GTxtNotesTraitements = GTxtNotesTraitements & "Reussi" & vbCrLf

                     Case False

                        GTxtNotesTraitements = GTxtNotesTraitements & JavaMsg.Log & vbCrLf

                        GTxtNotesTraitements = GTxtNotesTraitements & "Echec" & vbCrLf

                  End Select

                  DoEvents

                  RsRequete.MoveNext

               Loop

         End Select

         RsRequete.Close

   End Select

Exit_MailingJavaMailRequete:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set FsoSystem = Nothing

   Set FilFichier = Nothing

   Set HtmObjectDocument = Nothing

   Set HtmDocument = Nothing

   Set JavaMsg = Nothing

   Set RsFichier = Nothing

   Set RsRequete = Nothing

   Exit Function

Err_MailingJavaMailRequete:

   MailingJavaMailRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingJavaMailRequete"

   Resume Exit_MailingJavaMailRequete
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function MailingSendMailRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean
   Dim FsoSystem As Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim SndMsg As Object 'clsSendMail
   Dim SndMsgStatus As ClsSendMailStatus
   Dim HtmObjectDocument As MSHTML.HTMLDocument
   Dim HtmDocument As MSHTML.HTMLDocument
   Dim RsRequete As DAO.Recordset
   Dim RsFichier As DAO.Recordset
   Dim StrEmailFrom As String
   Dim StrAttache As String
   Dim StrAttacheFichier As String
   Dim StrDocument As String
   Dim StrCheminFichierBody As String
   Dim StrChampNom As String
   Dim IntNb As Integer
   Dim IntPosDebut As Integer
   Dim IntPosFin As Integer
   Dim IntEtatForm As Integer
   Dim BolFormFerme As Boolean

   On Error GoTo Err_MailingSendMailRequete

   MailingSendMailRequete = True

   DonneEtatFormulaire "FrmGestionTraitements", IntEtatForm

   Select Case IntEtatForm
      Case ETATFORMULAIRE

      Case Else

         BolFormFerme = True

         DoCmd.Hourglass True

         DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   End Select

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acWindowNormal

   GTxtTitTraitements = "Mailing : " & StrSujet

   GTxtFonTraitements = "MailingSendMailRequete"

   GTxtObjTraitements = StrSQLRequete

   DoEvents

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicCode2='" & StrFichierBody & "' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         StrCheminFichierBody = RsFichier!FicValeur & "\" & RsFichier!FicCode2

         RsFichier.Close

         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case RsRequete.EOF
            Case True

            Case False

               Set HtmObjectDocument = New MSHTML.HTMLDocument

               Set HtmDocument = HtmObjectDocument.createDocumentFromUrl(StrCheminFichierBody, vbNullString)

               Do Until HtmDocument.readyState = "complete"

                  DoEvents

               Loop

             '  Set SndMsg = New clsSendMail

             '  SndMsg.SMTPHostValidation = VALIDATE_NONE

           '    SndMsg.EmailAddressValidation = VALIDATE_SYNTAX

               SndMsg.Delimiter = ";"

               SndMsg.SMTPHost = "smtp.9online.fr"

               GTxtObjTraitements = "CONNEXION A " & SndMsg.SMTPHost & vbCrLf

               SndMsg.from = StrEmailExpediteur

               SndMsg.ReplyToAddress = StrEmailExpediteur

               Select Case BlnSimulation
                     Case True

                     Case False

                        SndMsg.BccRecipient = StrEmailCopieCacher

               End Select

               SndMsg.Subject = StrSujet

               SndMsg.AsHTML = True

               SndMsg.ContentBase = vbNullString

        '       SndMsg.EncodeType = MIME_ENCODE

          '     SndMsg.Priority = NORMAL_PRIORITY

               SndMsg.Receipt = True

               SndMsg.UseAuthentication = False

               SndMsg.UsePopAuthentication = False

               SndMsg.UserName = vbNullString

               SndMsg.Password = vbNullString

               SndMsg.POP3Host = vbNullString

               SndMsg.MaxRecipients = 100

               StrAttache = vbNullString

               Set FsoSystem = New Scripting.FileSystemObject

               Select Case HtmDocument.images.Length
                  Case 0

                  Case Else

                     For IntNb = 0 To HtmDocument.images.Length - 1

                        StrAttacheFichier = RemplaceChr(RemplaceChr(RemplaceChr(HtmDocument.images(IntNb).src, "file:///", vbNullString), "%20", " "), "/", "\")

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttache = StrAttache & StrAttacheFichier & SndMsg.Delimiter

                              Set FilFichier = FsoSystem.GetFile(StrAttacheFichier)

                              HtmDocument.images(IntNb).src = "cid:" & FilFichier.Name

                           Case False

                        End Select

                     Next

               End Select

               On Error Resume Next

               Select Case StrFichierAttache(LBound(StrFichierAttache))
                  Case vbNullString

                  Case Else

                     For IntNb = LBound(StrFichierAttache) To UBound(StrFichierAttache)

                        StrAttacheFichier = StrFichierAttache(IntNb)

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttache = StrAttache & StrAttacheFichier & SndMsg.Delimiter

                           Case False

                        End Select

                     Next

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

               End Select

               On Error GoTo Err_MailingSendMailRequete

               Select Case StrAttache
                  Case vbNullString

                  Case Else

                     StrAttache = Left(StrAttache, Len(StrAttache) - 1)

               End Select

               SndMsg.Attachment = StrAttache

               Set SndMsgStatus = New ClsSendMailStatus

               Set SndMsgStatus.Mail = SndMsg

               RsRequete.MoveLast

               RsRequete.MoveFirst

               GPbrTraitements.Min = 0

               GPbrTraitements.Value = 0

               GPbrTraitements.Max = RsRequete.RecordCount

               Select Case SndMsg.Connect
                  Case True

                     Do Until RsRequete.EOF = True

                        Select Case IsNull(RsRequete(StrChampDestination))
                           Case False

                              Select Case InStr(1, RsRequete(StrChampDestination), "mailto:")
                                 Case Is <= 0

                                    Select Case InStr(1, RsRequete(StrChampDestination), "#")
                                       Case Is <= 0

                                          StrEmailFrom = RsRequete(StrChampDestination)

                                       Case Else

                                          StrEmailFrom = Left(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "#") - 1)

                                    End Select

                                 Case Else

                                    StrEmailFrom = RemplaceChr(Mid(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "mailto:") + 7), "#", vbNullString)

                              End Select

                              Select Case BlnSimulation
                                 Case True

                                    SndMsg.Recipient = StrEmailExpediteur

                                    SndMsg.Subject = StrSujet & " (" & StrEmailFrom & ")"

                                 Case False

                                    SndMsg.Recipient = StrEmailFrom

                              End Select

                              StrDocument = "<HTML>" & vbCrLf & HtmDocument.documentElement.innerHTML & vbCrLf & "</HTML>"

                              IntPosDebut = 1

                              Do Until InStr(IntPosDebut, StrDocument, "<D VALUE=") <= 0

                                 IntPosDebut = InStr(IntPosDebut, StrDocument, "<D VALUE=")

                                 IntPosFin = InStr(IntPosDebut, StrDocument, ">")

                                 StrChampNom = RemplaceChr(Mid(StrDocument, IntPosDebut + 10, IntPosFin - IntPosDebut - 11), Chr$(34), "")

                                 IntPosFin = InStr(IntPosDebut, StrDocument, "</D>")

                                 On Error Resume Next

                                 StrDocument = Left(StrDocument, IntPosDebut - 1) & RsRequete(StrChampNom) & Mid(StrDocument, IntPosFin + 4)

                                 Select Case Err.Number
                                    Case vbEmpty

                                    Case Else

                                       Err.Clear

                                       IntPosDebut = IntPosFin

                                 End Select

                                 On Error GoTo Err_MailingSendMailRequete

                              Loop

                              SndMsg.Message = StrDocument

                           Case True

                        End Select

                        GPbrTraitements.Value = GPbrTraitements.Value + 1

                        Select Case BlnSimulation
                            Case True

                               GTxtObjTraitements = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                               GTxtNotesTraitements = GTxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                            Case False

                               GTxtObjTraitements = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                               GTxtNotesTraitements = GTxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                         End Select

                        SndMsg.Send

                        DoEvents

                        RsRequete.MoveNext

                     Loop

                     SndMsg.Disconnect

               End Select

               RsRequete.Close

         End Select

   End Select

Exit_MailingSendMailRequete:

   Select Case BolFormFerme
      Case True

         DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

         DoCmd.Hourglass False

      Case False

   End Select

   Set FsoSystem = Nothing

   Set FilFichier = Nothing

   Set HtmObjectDocument = Nothing

   Set HtmDocument = Nothing

   Set SndMsg = Nothing

   Set SndMsgStatus = Nothing

   Set RsFichier = Nothing

   Set RsRequete = Nothing

   Exit Function

Err_MailingSendMailRequete:

   MailingSendMailRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingSendMailRequete"

   Resume Exit_MailingSendMailRequete
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
'Public Function MailingMAPI(ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String) As Boolean
'   Dim MapSession As MAPI.Session
'   Dim MapMsg As MAPI.Message
'   Dim MapRecipients As MAPI.Recipients
'   Dim MapRecipient As MAPI.Recipient
'   Dim MapAttachs As MAPI.Attachments
'   Dim MapAttach As MAPI.Attachment
'   Dim MapFields As MAPI.Fields
'   Dim MapField As MAPI.Field
'   Dim HtmObjectDocument As MSHTML.HTMLDocument
'   Dim HtmDocument As MSHTML.HTMLDocument
'   Dim RsRequete As DAO.Recordset
'   Dim RsFichier As DAO.Recordset
'   Dim StrDocument As String
'   Dim StrCheminFichierBody As String
'   Dim StrChampNom As String
'   Dim IntNb As Integer
'   Dim IntPosDebut As Integer
'   Dim IntPosFin As Integer
'
'   On Error GoTo Err_MailingMAPI
'
'   MailingMAPI = True
'
'   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicCode2='" & StrFichierBody & "' AND FicValide=True")
'
'   Select Case RsFichier.EOF
'      Case True
'
'      Case False
'
'         StrCheminFichierBody = RsFichier!FicValeur & "\" & RsFichier!FicCode2
'
'         RsFichier.Close
'
'         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)
'
'         Select Case RsRequete.EOF
'            Case True
'
'            Case False
'
'               Set MapSession = CreateObject("MAPI.Session")
'
'               MapSession.Logon "", "", False, False
'
'               Do Until RsRequete.EOF = True
'
'                  Set HtmObjectDocument = New MSHTML.HTMLDocument
'
'                  Set HtmDocument = HtmObjectDocument.createDocumentFromUrl(StrCheminFichierBody, vbNullString)
'
'                  Do Until HtmDocument.readyState = "complete"
'
'                     DoEvents
'
'                  Loop
'
'                  Set MapMsg = MapSession.Outbox.Messages.Add
'
'                  Set MapAttachs = MapMsg.Attachments
'
'                  For IntNb = 0 To HtmDocument.images.Length - 1
'
'                     Set MapAttach = MapAttachs.Add(, , , RemplaceChr(RemplaceChr(HtmDocument.images(IntNb).src, "file:///", vbNullString), "%20", " "))
'
'                     Set MapFields = MapAttach.Fields
'
'                     Set MapField = MapFields.Add(CdoPR_ATTACH_MIME_TAG, "image/jpeg")
'
'                     Set MapField = MapFields.Add(&H3712001E, "IMAGE" & MapAttach.Index)
'
'                     HtmDocument.images(MapAttach.Index - 1).src = "cid:IMAGE" & MapAttach.Index
'
'                  Next
'
'                  MapMsg.Fields.Add "{0820060000000000C000000000000046}0x8514", 11, True
'
'                  MapMsg.Update
'
'                  Set MapAttachs = Nothing
'
'                  Set MapAttach = Nothing
'
'                  Set MapFields = Nothing
'
'                  Set MapField = Nothing
'
'                  Set MapRecipients = MapMsg.Recipients
'
'                  Select Case InStr(1, RsRequete(StrChampDestination), "mailto:")
'                     Case Is <= 0
'
'                        Set MapRecipient = MapRecipients.Add(RsRequete(StrChampDestination), , CdoTo)
'
'                     Case Else
'
'                        Set MapRecipient = MapRecipients.Add(RemplaceChr(Mid(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "mailto:") + 7), "#", vbNullString), , CdoTo)
'
'                  End Select
'
'                  MapMsg.Subject = StrSujet
'
'                  StrDocument = "<HTML>" & vbCrLf & HtmDocument.documentElement.innerHTML & vbCrLf & "</HTML>"
'
'                  IntPosDebut = 1
'
'                  Do Until InStr(IntPosDebut, StrDocument, "<D Value=") <= 0
'
'                     IntPosDebut = InStr(IntPosDebut, StrDocument, "<D Value=")
'
'                     IntPosFin = InStr(IntPosDebut, StrDocument, ">")
'
'                     StrChampNom = Mid(StrDocument, IntPosDebut + 10, IntPosFin - IntPosDebut - 11)
'
'                     IntPosFin = InStr(IntPosDebut, StrDocument, "</D>")
'
'                     StrDocument = Left(StrDocument, IntPosDebut - 1) & RsRequete(StrChampNom) & Mid(StrDocument, IntPosFin + 4)
'
'                  Loop
'
'                  MapMsg.Text = StrDocument
'
'                  MapMsg.Update
'
'                  MapMsg.Send True, False
'
'                  RsRequete.MoveNext
'
'               Loop
'
'               RsRequete.Close
'
'         End Select
'
'   End Select
'
'Exit_MailingMAPI:
'
'   Set MapSession = Nothing
'
'   Set MapRecipients = Nothing
'
'   Set MapRecipient = Nothing
'
'   Set MapAttachs = Nothing
'
'   Set MapAttach = Nothing
'
'   Set MapFields = Nothing
'
'   Set MapField = Nothing
'
'   Set MapMsg = Nothing
'
'   Set RsFichier = Nothing
'
'   Set RsRequete = Nothing
'
'   Exit Function
'
'Err_MailingMAPI:
'
'   MailingMAPI = False
'
'   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingMAPI"
'
'   Resume Exit_MailingMAPI
'End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function MailingLotusNotes(ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String) As Boolean

   On Error GoTo Err_MailingLotusNotes

   MailingLotusNotes = True

Exit_MailingLotusNotes:

   Exit Function

Err_MailingLotusNotes:

   MailingLotusNotes = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingOutlook"

   Resume Exit_MailingLotusNotes
End Function
