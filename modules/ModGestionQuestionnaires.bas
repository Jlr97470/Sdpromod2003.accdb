Attribute VB_Name = "ModGestionQuestionnaires"
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
'***    Declaration De Constantes Public                                                     ***
'******************************************************************************

'******************************************************************************
'***    Declaration De Procedure Public                                                       ***
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
Public Sub CreeAnalyseSocietesQuestionnaires(ByVal IntQuestionnairesNumero As Integer)
   Dim RsQuestionnairesTypesID As DAO.Recordset
   Dim QdfTable As DAO.QueryDef
   Dim StrAnalyseQuestionnaires As String
   Dim StrSQLAnalyseQuestionnairesDebut As String
   Dim StrSQLAnalyseQuestionnairesWhere As String
   Dim StrSQLAnalyseQuestionnairesFin As String
   Dim StrSQLAnalyseQuestionnairesChamp As String
   Dim StrSQLQuestionnairesTypeID As String

   On Error GoTo Err_CreeAnalyseSocietesQuestionnaires

   StrSQLQuestionnairesTypeID = "SELECT * FROM TBLSOCIETESQUESTIONNAIRESTYPESID WHERE SocQstTypIDSocQstTypNum=" & IntQuestionnairesNumero & ";"

   Set RsQuestionnairesTypesID = CurrentDb.OpenRecordset(StrSQLQuestionnairesTypeID)

   Select Case RsQuestionnairesTypesID.EOF
      Case True

      Case False

         StrAnalyseQuestionnaires = "AnaSocietesQst" & IntQuestionnairesNumero

         StrSQLAnalyseQuestionnairesDebut = "TRANSFORM First(SocQstResValeur) AS SocQstResValeurValeur " & _
               " SELECT SocQstNum, SocQstSocQstTypNum, SocQstSocNum, SocQstDateSaisie, SocQstNotes " & _
               " FROM ((TBLSOCIETESQUESTIONNAIRESTYPES INNER JOIN TBLSOCIETESQUESTIONNAIRESTYPESID ON TBLSOCIETESQUESTIONNAIRESTYPES.SocQstTypNum = TBLSOCIETESQUESTIONNAIRESTYPESID.SocQstTypIDSocQstTypNum) INNER JOIN TBLSOCIETESQUESTIONNAIRESRESULTATS ON TBLSOCIETESQUESTIONNAIRESTYPESID.SocQstTypIDNum = TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocQstTypIDNum) INNER JOIN TBLSOCIETESQUESTIONNAIRES ON (TBLSOCIETESQUESTIONNAIRESRESULTATS.SocQstResSocNum = TBLSOCIETESQUESTIONNAIRES.SocQstSocNum) AND (TBLSOCIETESQUESTIONNAIRESTYPES.SocQstTypNum = TBLSOCIETESQUESTIONNAIRES.SocQstSocQstTypNum) "

         StrSQLAnalyseQuestionnairesWhere = " WHERE TBLSOCIETESQUESTIONNAIRESTYPES.SocQstTypNum = " & IntQuestionnairesNumero

         StrSQLAnalyseQuestionnairesFin = " GROUP BY SocQstNum, SocQstSocQstTypNum, SocQstSocNum, SocQstDateSaisie, SocQstNotes, SocQstTypCode, SocQstTypLibelle, SocQstTypDateDebut, SocQstTypDateFin, SocQstTypNbResultat PIVOT 'SocQstResValeur' & [SocQstTypID] In "

         StrSQLAnalyseQuestionnairesChamp = "("

         Do Until RsQuestionnairesTypesID.EOF = True

            StrSQLAnalyseQuestionnairesChamp = StrSQLAnalyseQuestionnairesChamp & "'SocQstResValeur" & RsQuestionnairesTypesID!SocQstTypID & "',"

            RsQuestionnairesTypesID.MoveNext

         Loop

         StrSQLAnalyseQuestionnairesChamp = Left(StrSQLAnalyseQuestionnairesChamp, Len(StrSQLAnalyseQuestionnairesChamp) - 1) & ")"

         On Error Resume Next

         Set QdfTable = CurrentDb.QueryDefs(StrAnalyseQuestionnaires)

         Select Case Err.Number
            Case 0

               QdfTable.SQL = StrSQLAnalyseQuestionnairesDebut & StrSQLAnalyseQuestionnairesWhere & StrSQLAnalyseQuestionnairesFin & StrSQLAnalyseQuestionnairesChamp & ";"

            Case Else

               Err.Clear

               Set QdfTable = CurrentDb.CreateQueryDef(StrAnalyseQuestionnaires, StrSQLAnalyseQuestionnairesDebut & StrSQLAnalyseQuestionnairesWhere & StrSQLAnalyseQuestionnairesFin & StrSQLAnalyseQuestionnairesChamp & ";")

               CurrentDb.QueryDefs.Append QdfTable

         End Select

         On Error GoTo Err_CreeAnalyseSocietesQuestionnaires

      End Select

Exit_CreeAnalyseSocietesQuestionnaires:

   Set QdfTable = Nothing

   Set RsQuestionnairesTypesID = Nothing

   Exit Sub

Err_CreeAnalyseSocietesQuestionnaires:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeAnalyseSocietesQuestionnaires"

   Resume Exit_CreeAnalyseSocietesQuestionnaires
End Sub
