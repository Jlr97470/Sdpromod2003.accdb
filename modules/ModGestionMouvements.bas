Attribute VB_Name = "ModGestionMouvements"
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
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***    Constante Pour Les Exercices                                                           ***
'******************************************************************************
Public Const EXE2000 = 1
Public Const EXE2001 = 2
Public Const EXE2002 = 3
Public Const EXE2003 = 4
Public Const EXE2004 = 5
Public Const EXE2005 = 6
Public Const EXE2006 = 7
Public Const EXE2007 = 8

'******************************************************************************
'***    Constante Pour Les Semestres                                                          ***
'******************************************************************************
Public Const SEM2000S1 = 1
Public Const SEM2000S2 = 2
Public Const SEM2001S1 = 3
Public Const SEM2001S2 = 4
Public Const SEM2002S1 = 5
Public Const SEM2002S2 = 6
Public Const SEM2003S1 = 7
Public Const SEM2003S2 = 8
Public Const SEM2004S1 = 9
Public Const SEM2004S2 = 10
Public Const SEM2005S1 = 11
Public Const SEM2005S2 = 12
Public Const SEM2006S1 = 13
Public Const SEM2006S2 = 14
Public Const SEM2007S1 = 15
Public Const SEM2007S2 = 16

'******************************************************************************
'***    Constante Pour Les Mouvements                                                       ***
'******************************************************************************
Public Const MVTNATURESOLDE = 1
Public Const MVTNATURECADEAUX = 2
Public Const MVTNATURECHAUDIERES = 3
Public Const MVTNATUREBONUSANIMATION = 4

'******************************************************************************
'***    Constante Pour Les Chaudieres                                                         ***
'******************************************************************************
Private Const CHAUDIERES30 = 1000
Private Const CHAUDIERES50 = 2000
Private Const CHAUDIERES100 = 3000
Private Const CHAUDIERES101 = 4000
Private Const CHAUDIERESISO = 3000
Private Const CHAUDIERESRAP = 100

'******************************************************************************
'***    Declaration De Fonction Public                                                         ***
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
Public Function SupprimeMouvements(ByVal LngMvtSocNum As Long, ByVal lngMvtMvtNatNum As Long, ByVal lngMvtDossier As Long) As Boolean
   Dim StrSQLMouvements As String

   On Error GoTo Err_SupprimeMouvement

   SupprimeMouvements = True

   StrSQLMouvements = "DELETE * FROM TBLMOUVEMENTS WHERE TBLMOUVEMENTS.MvtSocNum=" & LngMvtSocNum & _
                                    " AND TBLMOUVEMENTS.MvtMvtNatNum=" & lngMvtMvtNatNum & _
                                    " AND TBLMOUVEMENTS.MvtDossier=" & lngMvtDossier & ";"

   CurrentDb.Execute StrSQLMouvements

Exit_SupprimeMouvement:

   Exit Function

Err_SupprimeMouvement:

   SupprimeMouvements = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "SupprimeMouvement"

   Resume Exit_SupprimeMouvement
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
Public Function AjouteMouvementsChaudieres(Optional ByVal LngChdRemSocNum As Long, Optional ByVal LngChdRemNum As Long) As Boolean
   Dim RsChaudieresRemontees As DAO.Recordset
   Dim RsMouvements As DAO.Recordset
   Dim StrSQLChaudieresRemonteesDebut As String
   Dim StrSQLChaudieresRemonteesFin As String
   Dim StrSQLChaudieresRemonteesSocNum As String
   Dim StrSQLChaudieresRemonteesNum As String
   Dim LngChaudieresNombre As Long
   Dim LngPoints As Long

   On Error GoTo Err_AjouteMouvementsChaudieres

   AjouteMouvementsChaudieres = True

   StrSQLChaudieresRemonteesDebut = "SELECT TBLCHAUDIERESREMONTEES.*, TBLCHAUDIERESARTICLES.ChdArtGamme, " & _
                                                           "TBLCHAUDIERESARTICLES.ChdArtLibelle, TBLEXERCICES.ExeExercice, TBLEXERCICES.ExeNum " & _
                                                           "FROM (TBLEXERCICES INNER JOIN TBLSEMESTRES ON TBLEXERCICES.ExeNum = TBLSEMESTRES.SemExeNum) " & _
                                                           "INNER JOIN (TBLCHAUDIERESARTICLES INNER JOIN (TBLCHAUDIERESREMONTEES LEFT JOIN SelMouvementsChaudieres " & _
                                                           "ON TBLCHAUDIERESREMONTEES.ChdRemNum = SelMouvementsChaudieres.MvtDossier) ON " & _
                                                           "TBLCHAUDIERESARTICLES.ChdArtRef = TBLCHAUDIERESREMONTEES.ChdRemChdArtRef) ON " & _
                                                           "TBLSEMESTRES.SemNum = TBLCHAUDIERESREMONTEES.ChdRemSemNum " & _
                                                           "WHERE SelMouvementsChaudieres.MvtDossier Is Null AND TBLCHAUDIERESREMONTEES.ChdRemRemp Is Not Null "

   StrSQLChaudieresRemonteesSocNum = IIf(LngChdRemSocNum <> vbEmpty, " AND ChdRemSocNum=" & LngChdRemSocNum & " ", vbNullString)

   StrSQLChaudieresRemonteesNum = IIf(LngChdRemNum <> vbEmpty, " AND ChdRemNum=" & LngChdRemNum & " ", vbNullString)

   StrSQLChaudieresRemonteesFin = " ORDER BY ChdRemDate, ChdRemHeure;"

   Set RsChaudieresRemontees = CurrentDb.OpenRecordset(StrSQLChaudieresRemonteesDebut & StrSQLChaudieresRemonteesSocNum & StrSQLChaudieresRemonteesNum & StrSQLChaudieresRemonteesFin)

   Select Case RsChaudieresRemontees.EOF
      Case True

      Case False

         Select Case LngChdRemNum
            Case vbEmpty

               DoCmd.Hourglass True

               DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

               GTxtTitTraitements = "Mise A Jour - Mouvements Chaudiéres"

               GTxtFonTraitements = "AjouteMouvementsChaudieres"

               GPbrTraitements.Min = 0

               RsChaudieresRemontees.MoveLast

               RsChaudieresRemontees.MoveFirst

               GPbrTraitements.Max = RsChaudieresRemontees.RecordCount

               GTxtObjTraitements = "Nombre De Chaudieres : " & RsChaudieresRemontees.RecordCount

               DoEvents

            Case Else

         End Select

         Set RsMouvements = CurrentDb.OpenRecordset("TBLMOUVEMENTS")

         Do Until RsChaudieresRemontees.EOF

            LngPoints = vbEmpty

            LngChaudieresNombre = DCount("ChdRemNum", "TBLCHAUDIERESREMONTEES", "(ChdRemSocNum=" & RsChaudieresRemontees!ChdRemSocNum & _
                                                  " AND ChdRemDate<CDate('" & RsChaudieresRemontees!ChdRemDate & "') AND Year(ChdRemServ)=" & Year(RsChaudieresRemontees!ChdRemServ) & " AND ChdRemRemp Is Not Null) OR (ChdRemSocNum=" & _
                                                  RsChaudieresRemontees!ChdRemSocNum & " AND ChdRemDate=CDate('" & RsChaudieresRemontees!ChdRemDate & "') AND " & _
                                                  "ChdRemHeure<=CDate('" & RsChaudieresRemontees!ChdRemHeure & "') AND Year(ChdRemServ)=" & Year(RsChaudieresRemontees!ChdRemServ) & " AND ChdRemRemp Is Not Null)")

            Select Case LngChaudieresNombre
               Case Is <= 30

                  LngPoints = CHAUDIERES30

               Case Is <= 50

                  LngPoints = CHAUDIERES50

               Case Is <= 100

                  LngPoints = CHAUDIERES100

               Case Else

                  LngPoints = CHAUDIERES101

            End Select

            Select Case RsChaudieresRemontees!ChdArtGamme
               Case "ISO"

                  LngPoints = LngPoints + CHAUDIERESISO

                  Select Case Year(RsChaudieresRemontees!ChdRemDate)
                     Case 2003

                        Select Case Year(RsChaudieresRemontees!ChdRemServ)
                           Case 2003

                              Select Case Month(RsChaudieresRemontees!ChdRemServ)
                                 Case 4, 5, 6

                                    LngPoints = LngPoints * 2

                                 Case Else

                                    Select Case Left(RsChaudieresRemontees!ChdArtLibelle, 8)
                                       Case "ISOSPLIT"

                                          Select Case Month(RsChaudieresRemontees!ChdRemServ)
                                             Case 7, 8, 9

                                                LngPoints = LngPoints * 2

                                             Case Else

                                          End Select

                                       Case Else

                                    End Select

                              End Select

                           Case Else

                        End Select

                     Case Else

                  End Select

               Case "THEMA"

               Case Else

            End Select

            Select Case DateDiff("d", RsChaudieresRemontees!ChdRemServ, RsChaudieresRemontees!ChdRemSD)
               Case Is <= 30

                  LngPoints = LngPoints + CHAUDIERESRAP

               Case Else

            End Select

            RsMouvements.AddNew

            RsMouvements!MvtSocNum = RsChaudieresRemontees!ChdRemSocNum

            RsMouvements!MvtExeNum = DLookup("ExeNum", "TBLEXERCICES", "ExeExercice=" & Year(RsChaudieresRemontees!ChdRemServ))

            RsMouvements!MvtMvtNatNum = MVTNATURECHAUDIERES

            RsMouvements!MvtDossier = RsChaudieresRemontees!ChdRemNum

            RsMouvements!MvtDate = RsChaudieresRemontees!ChdRemDate

            RsMouvements!MvtIntitule = RsChaudieresRemontees!ChdArtLibelle & " N°" & RsChaudieresRemontees!ChdRemSerie

            RsMouvements!MvtDebit = vbEmpty

            RsMouvements!MvtCredit = LngPoints

            RsMouvements!MvtQuantite = 1

            RsMouvements.Update

            Select Case LngChdRemNum
               Case vbEmpty

                  GPbrTraitements.Value = RsChaudieresRemontees.AbsolutePosition + 1

                  DoEvents

               Case Else

            End Select

            RsChaudieresRemontees.MoveNext

         Loop

         Select Case LngChdRemNum
            Case vbEmpty

               DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

               DoCmd.Hourglass False

            Case Else

         End Select

         RsMouvements.Close

   End Select

   RsChaudieresRemontees.Close

Exit_AjouteMouvementsChaudieres:

   Set RsChaudieresRemontees = Nothing

   Set RsMouvements = Nothing

   Exit Function

Err_AjouteMouvementsChaudieres:

   AjouteMouvementsChaudieres = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MouvementsChaudieres"

   Resume Exit_AjouteMouvementsChaudieres
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
Public Function AjouteMouvementsCadeaux(Optional ByVal LngCadAchSocNum As Long, Optional ByVal LngCadAchNum As Long) As Boolean
   Dim rsSocietesSoldesExe As DAO.Recordset
   Dim RsCadeauxAchats As DAO.Recordset
   Dim RsMouvements As DAO.Recordset
   Dim StrSQLCadeauxAchatsDebut As String
   Dim StrSQLCadeauxAchatsFin As String
   Dim StrSQLCadeauxAchatsSocNum As String
   Dim StrSQLCadeauxAchatsNum As String
   Dim lngPointsTotal As Long
   Dim lngPointsDebit As Long

   On Error GoTo Err_AjouteMouvementsCadeaux

   AjouteMouvementsCadeaux = True

   StrSQLCadeauxAchatsDebut = "SELECT TBLCADEAUXACHATS.*, TBLCADEAUXARTICLES.CadArtDescription FROM (TBLCADEAUXARTICLES INNER JOIN " & _
                                                 "TBLCADEAUXACHATS ON TBLCADEAUXARTICLES.CadArtNum = TBLCADEAUXACHATS.CadAchCadArtNum) LEFT JOIN " & _
                                                 "SelMouvementsCadeaux ON TBLCADEAUXACHATS.CadAchNum = SelMouvementsCadeaux.MvtDossier WHERE TBLCADEAUXACHATS.CadAchEtat='Commandé' " & _
                                                 "AND SelMouvementsCadeaux.MvtDossier Is Null "

   StrSQLCadeauxAchatsSocNum = IIf(LngCadAchSocNum <> vbEmpty, " AND CadAchSocNum=" & LngCadAchSocNum & " ", vbNullString)

   StrSQLCadeauxAchatsNum = IIf(LngCadAchNum <> vbEmpty, " AND CadAchNum=" & LngCadAchNum & " ", vbNullString)

   StrSQLCadeauxAchatsFin = " ORDER BY CadAchDateSaisie, CadAchHeureSaisie;"

   Set RsCadeauxAchats = CurrentDb.OpenRecordset(StrSQLCadeauxAchatsDebut & StrSQLCadeauxAchatsSocNum & StrSQLCadeauxAchatsNum & StrSQLCadeauxAchatsFin)

   Select Case RsCadeauxAchats.EOF
      Case True

      Case False

         Select Case LngCadAchNum
            Case vbEmpty

               DoCmd.Hourglass True

               DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

               GTxtTitTraitements = "Mise A Jour - Mouvements Cadeaux"

               GTxtFonTraitements = "AjouteMouvementsCadeaux"

               GPbrTraitements.Min = 0

               RsCadeauxAchats.MoveLast

               RsCadeauxAchats.MoveFirst

               GPbrTraitements.Max = RsCadeauxAchats.RecordCount

               GTxtObjTraitements = "Nombre De Cadeaux : " & RsCadeauxAchats.RecordCount

               DoEvents

            Case Else

         End Select

         Set RsMouvements = CurrentDb.OpenRecordset("TBLMOUVEMENTS")

         Do Until RsCadeauxAchats.EOF

'            Set rsSocietesSoldesExe = CurrentDb.OpenRecordset("SELECT * FROM SelSocietesSoldesExercices WHERE SelSocietesSoldesExercices.Solde>0 AND SelSocietesSoldesExercices.MvtSocNum=" & rsCadeauxAchats!CadAchSocNum)

            Set rsSocietesSoldesExe = CurrentDb.OpenRecordset("SELECT SelSocietesSoldesExercices.MvtSocNum, SelSocietesSoldesExercices.MvtExeNum, Switch([MvtExeNum]<4,[SelSocietesSoldesExercices].[Solde],[MvtExeNum]=4,[SelSocietesSoldesExercices].[Solde]+[PointsRea2003],[MvtExeNum]=5,[SelSocietesSoldesExercices].[Solde]+[PointsRea2004]) AS Solde " & _
            " FROM SelSocietesSoldesExercices LEFT JOIN TBLSOCIETESRESULTATSBONUS ON SelSocietesSoldesExercices.MvtSocNum = TBLSOCIETESRESULTATSBONUS.SocNum " & _
            " WHERE Switch([MvtExeNum]<4,[SelSocietesSoldesExercices].[Solde],[MvtExeNum]=4,[SelSocietesSoldesExercices].[Solde]+[PointsRea2003],[MvtExeNum]=5,[SelSocietesSoldesExercices].[Solde]+[PointsRea2004])>0 AND SelSocietesSoldesExercices.MvtSocNum=" & RsCadeauxAchats!CadAchSocNum & _
            " ORDER BY SelSocietesSoldesExercices.MvtSocNum, SelSocietesSoldesExercices.MvtExeNum;")

            lngPointsTotal = RsCadeauxAchats!CadAchPoints

            Do Until lngPointsTotal = vbEmpty

               RsMouvements.AddNew

               RsMouvements!MvtSocNum = RsCadeauxAchats!CadAchSocNum

               Select Case rsSocietesSoldesExe.EOF
                  Case False

                     Select Case ((Year(Date) - 1999) > rsSocietesSoldesExe!MvtExeNum)
                        Case True

                           Select Case lngPointsTotal > rsSocietesSoldesExe!Solde
                              Case True

                                 lngPointsDebit = rsSocietesSoldesExe!Solde

                                 lngPointsTotal = lngPointsTotal - rsSocietesSoldesExe!Solde

                              Case False

                                 lngPointsDebit = lngPointsTotal

                                 lngPointsTotal = vbEmpty

                           End Select

                           RsMouvements!MvtExeNum = rsSocietesSoldesExe!MvtExeNum

                        Case False

                           lngPointsDebit = lngPointsTotal

                           lngPointsTotal = vbEmpty

                           RsMouvements!MvtExeNum = (Year(Date) - 1999)

                     End Select

                  Case True

                     lngPointsDebit = lngPointsTotal

                     lngPointsTotal = vbEmpty

                     RsMouvements!MvtExeNum = (Year(Date) - 1999)

               End Select

               RsMouvements!MvtMvtNatNum = MVTNATURECADEAUX

               RsMouvements!MvtDossier = RsCadeauxAchats!CadAchNum

               RsMouvements!MvtDate = RsCadeauxAchats!CadAchDateSaisie

               RsMouvements!MvtIntitule = RsCadeauxAchats!CadArtDescription

               RsMouvements!MvtDebit = lngPointsDebit

               RsMouvements!MvtCredit = vbEmpty

               RsMouvements!MvtQuantite = RsCadeauxAchats!CadAchQuantite

               RsMouvements.Update

               Select Case LngCadAchNum
                  Case vbEmpty

                     GPbrTraitements.Value = RsCadeauxAchats.AbsolutePosition + 1

                     DoEvents

                  Case Else

               End Select

               Select Case rsSocietesSoldesExe.EOF
                  Case False

                     rsSocietesSoldesExe.MoveNext

                  Case True

               End Select

            Loop

            rsSocietesSoldesExe.Close

            RsCadeauxAchats.MoveNext

         Loop

         Select Case LngCadAchNum
            Case vbEmpty

               DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

               DoCmd.Hourglass False

            Case Else

         End Select

         RsMouvements.Close

   End Select

   RsCadeauxAchats.Close

Exit_AjouteMouvementsCadeaux:

   Set rsSocietesSoldesExe = Nothing

   Set RsCadeauxAchats = Nothing

   Set RsMouvements = Nothing

   Exit Function

Err_AjouteMouvementsCadeaux:

   AjouteMouvementsCadeaux = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteMouvementsCadeaux"

   Resume Exit_AjouteMouvementsCadeaux
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
Public Function SupprimeMouvementsChaudieres(Optional ByVal LngChdRemSocNum As Long, Optional ByVal LngChdRemNum As Long) As Boolean
   Dim StrSQLMouvementsChaudieres As String
   Dim DatChdRemDate As Date
   Dim DatChdRemHeure As Date

   On Error GoTo Err_SupprimeMouvementsChaudieres

   SupprimeMouvementsChaudieres = True

   Select Case LngChdRemNum
      Case vbEmpty

         Select Case LngChdRemSocNum
            Case vbEmpty

            Case Else

               StrSQLMouvementsChaudieres = "DELETE SelMouvementsChaudieres.* FROM SelMouvementsChaudieres " & _
                                                                 "WHERE SelMouvementsChaudieres.MvtSocNum=" & LngChdRemSocNum & ";"

               CurrentDb.Execute StrSQLMouvementsChaudieres

            End Select

      Case Else

         LngChdRemSocNum = Nz(DLookup("ChdRemSocNum", "TBLCHAUDIERESREMONTEES", "ChdRemNum=" & LngChdRemNum))

         Select Case LngChdRemSocNum
            Case vbEmpty

            Case Else

               DatChdRemDate = Nz(DLookup("ChdRemDate", "TBLCHAUDIERESREMONTEES", "ChdRemNum=" & LngChdRemNum))

               DatChdRemHeure = Nz(DLookup("ChdRemHeure", "TBLCHAUDIERESREMONTEES", "ChdRemNum=" & LngChdRemNum))

               StrSQLMouvementsChaudieres = "DELETE SelMouvementsChaudieres.* FROM SelMouvementsChaudieres INNER JOIN " & _
                                                                 "TBLCHAUDIERESREMONTEES ON SelMouvementsChaudieres.MvtDossier = TBLCHAUDIERESREMONTEES.ChdRemNum " & _
                                                                 "WHERE (((TBLCHAUDIERESREMONTEES.ChdRemDate)>CDate('" & DatChdRemDate & "')) AND " & _
                                                                 "((TBLCHAUDIERESREMONTEES.ChdRemSocNum)=" & LngChdRemSocNum & ") AND TBLCHAUDIERESREMONTEES.ChdRemRemp Is Not Null) OR " & _
                                                                 "(((TBLCHAUDIERESREMONTEES.ChdRemDate)=CDate('" & DatChdRemDate & "')) AND " & _
                                                                 "((TBLCHAUDIERESREMONTEES.ChdRemHeure)>=CDate('" & DatChdRemHeure & "')) AND ((TBLCHAUDIERESREMONTEES.ChdRemSocNum)=" & LngChdRemSocNum & ") AND TBLCHAUDIERESREMONTEES.ChdRemRemp Is Not Null);"

               CurrentDb.Execute StrSQLMouvementsChaudieres

         End Select

   End Select

Exit_SupprimeMouvementsChaudieres:

   Exit Function

Err_SupprimeMouvementsChaudieres:

   SupprimeMouvementsChaudieres = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "SupprimeMouvementsChaudieres"

   Resume Exit_SupprimeMouvementsChaudieres
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
Public Function SupprimeMouvementsCadeaux(Optional ByVal LngCadAchSocNum As Long, Optional ByVal LngCadAchNum As Long) As Boolean
   Dim StrSQLMouvementsCadeaux As String
   Dim DatCadAchDate As Date
   Dim DatCadAchHeure As Date

   On Error GoTo Err_SupprimeMouvementsCadeaux

   SupprimeMouvementsCadeaux = True

   Select Case LngCadAchNum
      Case vbEmpty

         Select Case LngCadAchSocNum
            Case vbEmpty

            Case Else

               StrSQLMouvementsCadeaux = "DELETE SelMouvementsCadeaux.* FROM SelMouvementsCadeaux " & _
                                                             "WHERE SelMouvementsCadeaux.MvtSocNum=" & LngCadAchSocNum & ";"

               CurrentDb.Execute StrSQLMouvementsCadeaux

            End Select

      Case Else

         LngCadAchSocNum = Nz(DLookup("CadAchSocNum", "TBLCADEAUXACHATS", "CadAchNum=" & LngCadAchNum))

         Select Case LngCadAchSocNum
            Case vbEmpty

            Case Else

               DatCadAchDate = Nz(DLookup("CadAchDateSaisie", "TBLCADEAUXACHATS", "CadAchNum=" & LngCadAchNum))

               DatCadAchHeure = Nz(DLookup("CadAchHeureSaisie", "TBLCADEAUXACHATS", "CadAchNum=" & LngCadAchNum))

               StrSQLMouvementsCadeaux = "DELETE SelMouvementsCadeaux.* FROM SelMouvementsCadeaux INNER JOIN " & _
                                                                 "TBLCADEAUXACHATS ON SelMouvementsCadeaux.MvtDossier = TBLCADEAUXACHATS.CadAchNum " & _
                                                                 "WHERE (((TBLCADEAUXACHATS.CadAchDateSaisie)>CDate('" & DatCadAchDate & "')) AND " & _
                                                                 "((TBLCADEAUXACHATS.CadAchSocNum)=" & LngCadAchSocNum & ")) OR " & _
                                                                 "(((TBLCADEAUXACHATS.CadAchDateSaisie)=CDate('" & DatCadAchDate & "')) AND " & _
                                                                 "((TBLCADEAUXACHATS.CadAchHeureSaisie)>=CDate('" & DatCadAchHeure & "')) AND ((TBLCADEAUXACHATS.CadAchSocNum)=" & LngCadAchSocNum & "));"

               CurrentDb.Execute StrSQLMouvementsCadeaux

         End Select

   End Select

Exit_SupprimeMouvementsCadeaux:

   Exit Function

Err_SupprimeMouvementsCadeaux:

   SupprimeMouvementsCadeaux = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "SupprimeMouvementsCadeaux"

   Resume Exit_SupprimeMouvementsCadeaux
End Function
