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
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************
Private StrSocietesCle As String
Private StrContactsCle As String
Private StrChaudieresCle As String
Private StrCadeauxCle As String

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
Private Function CreeSQLRequete() As String
   Dim StrSQLRequeteTable As String
   Dim StrSQLRequeteTableWhere As String
   Dim StrSQLRequeteTableOrder As String
   Dim StrCondition As String
   Dim StrValeur As String
   Dim StrOrdre() As String
   Dim IntChamp As Integer

   On Error GoTo Err_CreeSQLRequete

   CreeSQLRequete = vbNullString

   Select Case CmbTableRequete.ListIndex
      Case Is < 0

      Case Else

         Select Case LstChampSelectionner.ListCount
            Case 0

            Case Else

               ReDim StrOrdre(LstChampSelectionner.ListCount - 1)

               StrSQLRequeteTable = vbNullString

               StrSQLRequeteTableWhere = vbNullString

               StrSQLRequeteTableOrder = vbNullString

               For IntChamp = 0 To LstChampSelectionner.ListCount - 1

                  Select Case (InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") <= 0 Or IsNull(InStr(1, LstChampSelectionner.Column(3, IntChamp), " ")))
                     Case True

                     Case False

                        StrCondition = vbNullString

                        StrValeur = vbNullString

                        Select Case Left(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") - 1)
                           Case "Comme"

                              StrCondition = "Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                           Case "Pas_Comme"

                              StrCondition = "Not Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                           Case Else

                              StrCondition = Left(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") - 1)

                              Select Case LstChampSelectionner.Column(2, IntChamp)
                                 Case dbBoolean

                                    Select Case Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1)
                                       Case "Vrai"

                                          StrValeur = "True"

                                       Case "Faux"

                                          StrValeur = "False"

                                    End Select

                                 Case dbDate, dbTime, dbTimeStamp

                                    StrValeur = "CDate('" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "')"

                                 Case dbChar, dbText, dbMemo

                                    StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                                 Case Else

                                    StrValeur = RemplaceChr(Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1), ",", ".")

                              End Select

                        End Select

                        StrSQLRequeteTableWhere = StrSQLRequeteTableWhere & "[" & CmbTableRequete.Column(0) & "].[" & LstChampSelectionner.Column(0, IntChamp) & "] " & StrCondition & " " & StrValeur & " AND "

                  End Select

                  Select Case (InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") <= 0) Or IsNull(InStr(1, LstChampSelectionner.Column(4, IntChamp), " "))
                     Case True

                     Case False

                        StrCondition = vbNullString

                        StrValeur = vbNullString

                        Select Case Left(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") - 1)
                           Case "Comme"

                              StrCondition = "Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                           Case "Pas_Comme"

                              StrCondition = "Not Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                           Case Else

                              StrCondition = Left(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") - 1)

                              Select Case LstChampSelectionner.Column(2, IntChamp)
                                 Case dbBoolean

                                    Select Case Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1)
                                       Case "Vrai", "Oui"

                                          StrValeur = "True"

                                       Case "Faux", "Non"

                                          StrValeur = "False"

                                    End Select

                                 Case dbDate, dbTime, dbTimeStamp

                                    StrValeur = "CDate('" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "')"

                                 Case dbChar, dbText, dbMemo

                                    StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                                 Case Else

                                    StrValeur = RemplaceChr(Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1), ",", ".")

                              End Select

                        End Select

                        StrSQLRequeteTableWhere = StrSQLRequeteTableWhere & "[" & CmbTableRequete.Column(0) & "].[" & LstChampSelectionner.Column(0, IntChamp) & "] " & StrCondition & " " & StrValeur & " AND "

                  End Select

                  Select Case (InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") <= 0) Or IsNull(InStr(1, LstChampSelectionner.Column(5, IntChamp), " "))
                     Case True

                     Case False

                        Select Case Mid(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") + 1)
                           Case "Croisant"

                              StrOrdre(Left(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") - 1)) = "[" & CmbTableRequete.Column(0) & "].[" & LstChampSelectionner.Column(0, IntChamp) & "] ASC"

                           Case "Decroisant"

                              StrOrdre(Left(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") - 1)) = "[" & CmbTableRequete.Column(0) & "].[" & LstChampSelectionner.Column(0, IntChamp) & "] DESC"

                        End Select

                  End Select

               Next

               StrSQLRequeteTable = " FROM [" & CmbTableRequete.Column(0) & "]"

               Select Case StrSQLRequeteTableWhere
                  Case vbNullString

                  Case Else

                     StrSQLRequeteTableWhere = Left(StrSQLRequeteTableWhere, Len(StrSQLRequeteTableWhere) - 5)

                     StrSQLRequeteTable = StrSQLRequeteTable & " WHERE " & StrSQLRequeteTableWhere

               End Select

               For IntChamp = 0 To LstChampSelectionner.ListCount - 1

                  Select Case StrOrdre(IntChamp)
                     Case vbNullString

                     Case Else

                        StrSQLRequeteTableOrder = StrSQLRequeteTableOrder & StrOrdre(IntChamp) & ", "

                  End Select

               Next

               Select Case StrSQLRequeteTableOrder
                  Case vbNullString

                  Case Else

                     StrSQLRequeteTableOrder = Left(StrSQLRequeteTableOrder, Len(StrSQLRequeteTableOrder) - 2)

                     StrSQLRequeteTable = StrSQLRequeteTable & " ORDER BY " & StrSQLRequeteTableOrder

               End Select

               StrSQLRequeteTable = StrSQLRequeteTable & ";"

         End Select

   End Select

   CreeSQLRequete = StrSQLRequeteTable

Exit_CreeSQLRequete:

   Exit Function

Err_CreeSQLRequete:

   CreeSQLRequete = vbNullString

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeSQLRequete"

   Resume Exit_CreeSQLRequete
End Function

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
Private Sub CreeListeTableRequete()
   Dim TdfTable As DAO.TableDef
   Dim QdfRequete As DAO.QueryDef
   Dim IntTableRequeteIndex As Integer

   On Error GoTo Err_CreeListeTableRequete

   Select Case CmbTableRequete.ListCount
      Case 0

      Case Else

         For IntTableRequeteIndex = 0 To CmbTableRequete.ListCount - 1

            CmbTableRequete.RemoveItem IntTableRequeteIndex

         Next

   End Select

   For Each TdfTable In CurrentDb.TableDefs

      Select Case (TdfTable.Attributes And dbSystemObject) <> 0 Or (TdfTable.Attributes And dbHiddenObject) <> 0
         Case True

         Case False

            On Error Resume Next

            CmbTableRequete.AddItem TdfTable.Name & ";" & TdfTable.Properties("Description") & ";" & "TABLE"

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  CmbTableRequete.AddItem TdfTable.Name & ";" & ";" & "TABLE"

            End Select

            On Error GoTo Err_CreeListeTableRequete

      End Select

   Next

   For Each QdfRequete In CurrentDb.QueryDefs

      Select Case Left(QdfRequete.Name, 4) = "~sq_" Or QdfRequete.Type = dbQProcedure Or QdfRequete.Type = dbQAction Or QdfRequete.Type = dbQAppend Or QdfRequete.Type = dbQDelete Or QdfRequete.Type = dbQMakeTable Or QdfRequete.Type = dbQUpdate
         Case True

         Case False

            On Error Resume Next

            CmbTableRequete.AddItem QdfRequete.Name & ";" & QdfRequete.Properties("Description") & ";" & "REQUETE"

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  CmbTableRequete.AddItem QdfRequete.Name & ";" & ";" & "REQUETE"

            End Select

            On Error GoTo Err_CreeListeTableRequete

      End Select

   Next

Exit_CreeListeTableRequete:

   Set TdfTable = Nothing

   Exit Sub

Err_CreeListeTableRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeListeTableRequete"

   Resume Exit_CreeListeTableRequete
End Sub

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
Private Sub CreeListeChampTableRequete()
   Dim TdsTable As DAO.TableDefs
   Dim TdfTable As DAO.TableDef
   Dim QdsRequete As DAO.QueryDefs
   Dim QdfRequete As DAO.QueryDef
   Dim FldChamp As DAO.Field

   On Error GoTo Err_CreeListeChampTableRequete

   LstChamp.RowSource = vbNullString

   Select Case CmbTableRequete.Column(0)
      Case vbNullString

      Case Else

         Select Case CmbTableRequete.Column(2)
            Case "TABLE"

               Set TdsTable = CurrentDb.TableDefs

               Set TdfTable = TdsTable(CmbTableRequete.Column(0))

               For Each FldChamp In TdfTable.Fields

                  On Error Resume Next

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description") & ";" & FldChamp.Properties("Type")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChamp.AddItem FldChamp.Name & ";;" & FldChamp.Properties("Type")

                  End Select

                  On Error GoTo Err_CreeListeChampTableRequete

               Next

            Case "REQUETE"

               Set QdsRequete = CurrentDb.QueryDefs

               Set QdfRequete = QdsRequete(CmbTableRequete.Column(0))

               For Each FldChamp In QdfRequete.Fields

                  On Error Resume Next

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description") & ";" & FldChamp.Properties("Type")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChamp.AddItem FldChamp.Name & ";;" & FldChamp.Properties("Type")

                  End Select

                  On Error GoTo Err_CreeListeChampTableRequete

               Next
      End Select

   End Select

Exit_CreeListeChampTableRequete:

   Exit Sub

Err_CreeListeChampTableRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeListeChampTableRequete"

   Resume Exit_CreeListeChampTableRequete
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

   CmbTableRequete.RowSource = vbNullString

   LstChamp.RowSource = vbNullString

   LstChampSelectionner.RowSource = vbNullString

   CreeListeTableRequete

   CmbChampCondition1 = vbNullString

   CmbChampCondition2 = vbNullString

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

   On Error Resume Next

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

   End Select

   On Error GoTo Err_CmdFermer_Click

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

   OuvreRecherches

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
Private Sub CmdMinimize_Click()

   On Error GoTo Err_CmdMinimize_Click

   DoCmd.Minimize

Exit_CmdMinimize_Click:

   Exit Sub

Err_CmdMinimize_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdMinimize_Click"

   Resume Exit_CmdMinimize_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherSocietes_Click()
   Dim RsSocietesVilles As DAO.Recordset
   Dim StrSQLRequete As String
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   StrSQLRequete = CreeSQLRequete

   Select Case StrSQLRequete
      Case vbNullString

      Case Else

         GStrSocietesFiltre = "[SocNum] In(SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrSocietesCle & "] " & StrSQLRequete & ")"

         Set RsSocietesVilles = CurrentDb.OpenRecordset("SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrSocietesCle & "] " & StrSQLRequete)

         Select Case RsSocietesVilles.EOF
            Case True

            Case False

               RsSocietesVilles.MoveLast

         End Select

         Select Case RsSocietesVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdAfficherSocietes_Click

            Case 1

               DonneEtatFormulaire "FrmFicheSocietes", IntEtat

               Select Case IntEtat
                  Case ETATFORMULAIRE

                     Forms!FrmFicheSocietes.SetFocus

                     On Error Resume Next

                     DoCmd.ApplyFilter , "SocNum = " & RsSocietesVilles!SocNum

                     Select Case Err.Number
                        Case 0

                        Case 2501

                           Err.Clear

                        Case Else

                           GoTo Err_CmdAfficherSocietes_Click

                     End Select

                     On Error GoTo Err_CmdAfficherSocietes_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum = " & RsSocietesVilles!SocNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstSocietes", acSavePrompt

               DoCmd.OpenForm "FrmLstSocietes", acNormal, , , acFormEdit, acWindowNormal, GStrSocietesFiltre

         End Select

         RsSocietesVilles.Close

   End Select

Exit_CmdAfficherSocietes_Click:

   Set RsSocietesVilles = Nothing

   Exit Sub

Err_CmdAfficherSocietes_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherSocietes_Click"

   Resume Exit_CmdAfficherSocietes_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherContacts_Click()
   Dim RsContactsVilles As DAO.Recordset
   Dim StrSQLRequete As String
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherContacts_Click

   StrSQLRequete = CreeSQLRequete

   Select Case StrSQLRequete
      Case vbNullString

      Case Else

         GStrContactsFiltre = "[ConNum] In(SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrContactsCle & "] " & StrSQLRequete & ")"

         Set RsContactsVilles = CurrentDb.OpenRecordset("SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrContactsCle & "] " & StrSQLRequete)

         Select Case RsContactsVilles.EOF
            Case True

            Case False

               RsContactsVilles.MoveLast

         End Select

         Select Case RsContactsVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdAfficherContacts_Click

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

                           GoTo Err_CmdAfficherContacts_Click

                     End Select

                     On Error GoTo Err_CmdAfficherContacts_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & RsContactsVilles!ConNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstContacts", acSavePrompt

               DoCmd.OpenForm "FrmLstContacts", acNormal, , , acFormEdit, acWindowNormal, GStrContactsFiltre

         End Select

         RsContactsVilles.Close

   End Select

Exit_CmdAfficherContacts_Click:

   Set RsContactsVilles = Nothing

   Exit Sub

Err_CmdAfficherContacts_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherContacts_Click"

   Resume Exit_CmdAfficherContacts_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherChaudieres_Click()
   Dim RsChaudieresRemontees As DAO.Recordset
   Dim StrSQLRequete As String
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherChaudieres_Click

   StrSQLRequete = CreeSQLRequete

   Select Case StrSQLRequete
      Case vbNullString

      Case Else

         GStrChaudieresRemonteesFiltre = "[ChdRemNum] In(SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrChaudieresCle & "] " & StrSQLRequete & ")"

         Set RsChaudieresRemontees = CurrentDb.OpenRecordset("SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrChaudieresCle & "] " & StrSQLRequete)

         Select Case RsChaudieresRemontees.EOF
            Case True

            Case False

               RsChaudieresRemontees.MoveLast

         End Select

         Select Case RsChaudieresRemontees.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdAfficherChaudieres_Click

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

                           GoTo Err_CmdAfficherChaudieres_Click

                     End Select

                     On Error GoTo Err_CmdAfficherChaudieres_Click

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

Exit_CmdAfficherChaudieres_Click:

   Set RsChaudieresRemontees = Nothing

   Exit Sub

Err_CmdAfficherChaudieres_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherChaudieres_Click"

   Resume Exit_CmdAfficherChaudieres_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherCadeaux_Click()
   Dim RsCadeauxAchats As DAO.Recordset
   Dim StrSQLRequete As String
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherCadeaux_Click

   StrSQLRequete = CreeSQLRequete

   Select Case StrSQLRequete
      Case vbNullString

      Case Else

         GStrCadeauxAchatsFiltre = "[CadAchNum] In(SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrCadeauxCle & "] " & StrSQLRequete & ")"

         Set RsCadeauxAchats = CurrentDb.OpenRecordset("SELECT DISTINCT [" & CmbTableRequete.Column(0) & "].[" & StrCadeauxCle & "] " & StrSQLRequete)

         Select Case RsCadeauxAchats.EOF
            Case True

            Case False

               RsCadeauxAchats.MoveLast

         End Select

         Select Case RsCadeauxAchats.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdAfficherCadeaux_Click

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

                           GoTo Err_CmdAfficherCadeaux_Click

                     End Select

                     On Error GoTo Err_CmdAfficherCadeaux_Click

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmSaisieCadeauxAchats", acNormal, , "CadAchNum=" & RsCadeauxAchats!CadAchNum, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstCadeauxAchats", acSavePrompt

               DoCmd.OpenForm "FrmLstCadeauxAchats", acNormal, , , acFormEdit, acWindowNormal, GStrCadeauxAchatsFiltre

         End Select

         RsCadeauxAchats.Close

   End Select

Exit_CmdAfficherCadeaux_Click:

   Set RsCadeauxAchats = Nothing

   Exit Sub

Err_CmdAfficherCadeaux_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCadeaux_Click"

   Resume Exit_CmdAfficherCadeaux_Click
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
   Dim IntIndex As Integer
   Dim IntOrdre As Integer
   Dim TxtChamp(2) As String
   Dim TxtCondition1Valeur As String
   Dim TxtCondition2Valeur As String
   Dim TxtOrdreValeur As String

   On Error GoTo Err_CmdValider_Click

   Select Case LstChampSelectionner.ListIndex
      Case Is >= 0

         IntOrdre = 0

         TxtCondition1Valeur = vbNullString

         TxtCondition2Valeur = vbNullString

         TxtOrdreValeur = vbNullString

         Select Case InStr(1, LstChampSelectionner.Column(5), " ")
            Case Is > 0

               IntOrdre = CInt(Left(LstChampSelectionner.Column(5), InStr(1, LstChampSelectionner.Column(5), " ") - 1))

            Case Else

               For IntIndex = 0 To LstChampSelectionner.ListCount - 1

                  Select Case InStr(1, LstChampSelectionner.Column(5, IntIndex), " ")
                     Case Is > 0

                        Select Case CInt(Left(LstChampSelectionner.Column(5, IntIndex), InStr(1, LstChampSelectionner.Column(5, IntIndex), " ") - 1))
                           Case Is > 0

                              IntOrdre = CInt(Left(LstChampSelectionner.Column(5, IntIndex), InStr(1, LstChampSelectionner.Column(5, IntIndex), " ") - 1))

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

               IntOrdre = IntOrdre + 1

         End Select

         IntIndex = LstChampSelectionner.ListIndex

         TxtChamp(0) = LstChampSelectionner.Column(0)

         TxtChamp(1) = LstChampSelectionner.Column(1)

         TxtChamp(2) = LstChampSelectionner.Column(2)

         Select Case CmbChampCondition1.ListIndex
            Case Is < 0

            Case Else

               Select Case InStr(1, CmbChampValeurCondition1, "#")
                  Case 0

                     TxtCondition1Valeur = CmbChampCondition1.Column(0) & " " & CmbChampValeurCondition1

                  Case Else

                     TxtCondition1Valeur = CmbChampCondition1.Column(0) & " " & Left(CmbChampValeurCondition1, InStr(1, CmbChampValeurCondition1, "#") - 1)

               End Select

         End Select

         Select Case CmbChampCondition2.ListIndex
            Case Is < 0

            Case Else

               Select Case InStr(1, CmbChampValeurCondition2, "#")
                  Case 0

                     TxtCondition2Valeur = CmbChampCondition2.Column(0) & " " & CmbChampValeurCondition2

                  Case Else

                     TxtCondition2Valeur = CmbChampCondition2.Column(0) & " " & Left(CmbChampValeurCondition2, InStr(1, CmbChampValeurCondition2, "#") - 1)

               End Select

         End Select

         Select Case CmbChampOrdre.ListIndex
            Case Is < 0

            Case Else

               TxtOrdreValeur = IntOrdre & " " & CmbChampOrdre

         End Select

         LstChampSelectionner.RemoveItem IntIndex

         Select Case LstChampSelectionner.ListCount
            Case IntIndex

               LstChampSelectionner.AddItem "'" & TxtChamp(0) & "';'" & TxtChamp(1) & "';'" & TxtChamp(2) & "';'" & TxtCondition1Valeur & "';'" & TxtCondition2Valeur & "';'" & TxtOrdreValeur & "'"

            Case Else

               LstChampSelectionner.AddItem "'" & TxtChamp(0) & "';'" & TxtChamp(1) & "';'" & TxtChamp(2) & "';'" & TxtCondition1Valeur & "';'" & TxtCondition2Valeur & "';'" & TxtOrdreValeur & "'", IntIndex

         End Select

      Case Else

   End Select

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
Private Sub LstChamp_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChamp_DblClick

   Select Case LstChamp.ListIndex
      Case Is >= 0

         Select Case 0
            Case Is < InStr(1, LstChamp.Column(0), "SocNum")

               StrSocietesCle = LstChamp.Column(0)

               CmdAfficherSocietes.Enabled = True

            Case Is < InStr(1, LstChamp.Column(0), "ConNum")

               StrContactsCle = LstChamp.Column(0)

               CmdAfficherContacts.Enabled = True

            Case Is < InStr(1, LstChamp.Column(0), "ChdRemNum")

               StrChaudieresCle = LstChamp.Column(0)

               CmdAfficherChaudieres.Enabled = True

            Case Is < InStr(1, LstChamp.Column(0), "CadAchNum")

               StrCadeauxCle = LstChamp.Column(0)

               CmdAfficherCadeaux.Enabled = True

            Case Else

         End Select

         LstChampSelectionner.AddItem "'" & LstChamp.Column(0) & "';'" & LstChamp.Column(1) & "';'" & LstChamp.Column(2) & "';;;;"

         LstChamp.RemoveItem LstChamp.ListIndex

      Case Else

   End Select

Exit_LstChamp_DblClick:

   Exit Sub

Err_LstChamp_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChamp_DblClick"

   Resume Exit_LstChamp_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChampSelectionner_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChampSelectionner_DblClick

   Select Case LstChampSelectionner.ListIndex
      Case Is >= 0

         Select Case 0
            Case Is < InStr(1, LstChamp.Column(0), "SocNum")

               StrSocietesCle = vbNullString

               CmdAfficherSocietes.Enabled = False

            Case Is < InStr(1, LstChamp.Column(0), "ConNum")

               StrContactsCle = vbNullString

               CmdAfficherContacts.Enabled = False

            Case Is < InStr(1, LstChamp.Column(0), "ChdRemNum")

               StrChaudieresCle = vbNullString

               CmdAfficherChaudieres.Enabled = False

            Case Is < InStr(1, LstChamp.Column(0), "CadAchNum")

               StrCadeauxCle = vbNullString

               CmdAfficherCadeaux.Enabled = False

            Case Else

         End Select

         LstChamp.AddItem "'" & LstChampSelectionner.Column(0) & "';'" & LstChampSelectionner.Column(1) & "';'" & LstChamp.Column(2) & "'"

         LstChampSelectionner.RemoveItem LstChampSelectionner.ListIndex

      Case Else

   End Select

Exit_LstChampSelectionner_DblClick:

   Exit Sub

Err_LstChampSelectionner_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChampSelectionner_DblClick"

   Resume Exit_LstChampSelectionner_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbTableRequete_AfterUpdate()

   On Error GoTo Err_CmbTableRequete_AfterUpdate

   CreeListeChampTableRequete

   LstChampSelectionner.RowSource = vbNullString

   CmbChampValeurCondition1.RowSource = vbNullString

   CmbChampValeurCondition1 = vbNullString

   CmbChampCondition1 = vbNullString

   CmbChampCondition1.Enabled = True

   CmbChampCondition1.Locked = False

   CmbChampValeurCondition2.RowSource = vbNullString

   CmbChampValeurCondition2 = vbNullString

   CmbChampCondition2 = vbNullString

   CmbChampCondition2.Enabled = True

   CmbChampCondition2.Locked = False

   CmbChampOrdre = vbNullString

   CmbChampOrdre.Enabled = True

   CmbChampOrdre.Locked = False

   StrSocietesCle = vbNullString

   CmdAfficherSocietes.Enabled = False

   StrContactsCle = vbNullString

   CmdAfficherContacts.Enabled = False

   StrChaudieresCle = vbNullString

   CmdAfficherChaudieres.Enabled = False

   StrCadeauxCle = vbNullString

   CmdAfficherCadeaux.Enabled = False

Exit_CmbTableRequete_AfterUpdate:

   Exit Sub

Err_CmbTableRequete_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbTableRequete_AfterUpdate"

   Resume Exit_CmbTableRequete_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChampSelectionner_AfterUpdate()

   On Error GoTo Err_LstChampSelectionner_AfterUpdate

   CmbChampCondition1.Enabled = True

   CmbChampCondition1.Locked = False

   CmbChampValeurCondition1.RowSource = "SELECT DISTINCT " & TxtChampSelectionner & " FROM " & CmbTableRequete.Column(0) & " ORDER BY " & TxtChampSelectionner & ";"

   Select Case InStr(1, LstChampSelectionner.Column(3), " ")
      Case Is > 0

         CmbChampValeurCondition1.Enabled = True

         CmbChampValeurCondition1.Locked = False

         CmbChampCondition1 = Left(LstChampSelectionner.Column(3), InStr(1, LstChampSelectionner.Column(3), " ") - 1)

         CmbChampValeurCondition1 = Mid(LstChampSelectionner.Column(3), InStr(1, LstChampSelectionner.Column(3), " ") + 1)

      Case Else

         CmbChampValeurCondition1.Enabled = False

         CmbChampValeurCondition1.Locked = True

         CmbChampCondition1 = vbNullString

         CmbChampValeurCondition1 = vbNullString

   End Select

   CmbChampCondition2.Enabled = True

   CmbChampCondition2.Locked = False

   CmbChampValeurCondition2.RowSource = "SELECT DISTINCT " & TxtChampSelectionner & " FROM " & CmbTableRequete.Column(0) & " ORDER BY " & TxtChampSelectionner & ";"

   Select Case InStr(1, LstChampSelectionner.Column(4), " ")
      Case Is > 0

         CmbChampValeurCondition2.Enabled = True

         CmbChampValeurCondition2.Locked = False

         CmbChampCondition2 = Left(LstChampSelectionner.Column(4), InStr(1, LstChampSelectionner.Column(4), " ") - 1)

         CmbChampValeurCondition2 = Mid(LstChampSelectionner.Column(4), InStr(1, LstChampSelectionner.Column(4), " ") + 1)

      Case Else

         CmbChampValeurCondition2.Enabled = False

         CmbChampValeurCondition2.Locked = True

         CmbChampCondition2 = vbNullString

         CmbChampValeurCondition2 = vbNullString

   End Select

   CmbChampOrdre.Enabled = True

   CmbChampOrdre.Locked = False

   Select Case InStr(1, LstChampSelectionner.Column(5), " ")
      Case Is > 0

         CmbChampOrdre = Mid(LstChampSelectionner.Column(5), InStr(1, LstChampSelectionner.Column(5), " ") + 1)

      Case Else

         CmbChampOrdre = vbNullString

   End Select

Exit_LstChampSelectionner_AfterUpdate:

   Exit Sub

Err_LstChampSelectionner_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChampSelectionner_AfterUpdate"

   Resume Exit_LstChampSelectionner_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampCondition1_AfterUpdate()

   On Error GoTo Err_CmbChampCondition1_AfterUpdate

   Select Case CmbChampCondition1.ListIndex
      Case Is <= 0

         CmbChampValeurCondition1 = vbNullString

         CmbChampValeurCondition1.Enabled = False

         CmbChampValeurCondition1.Locked = True

      Case Else

         CmbChampValeurCondition1.Enabled = True

         CmbChampValeurCondition1.Locked = False

   End Select

Exit_CmbChampCondition1_AfterUpdate:

   Exit Sub

Err_CmbChampCondition1_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampCondition1_AfterUpdate"

   Resume Exit_CmbChampCondition1_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampCondition2_AfterUpdate()

   On Error GoTo Err_CmbChampCondition2_AfterUpdate

   Select Case CmbChampCondition2.ListIndex
      Case Is <= 0

         CmbChampValeurCondition2 = vbNullString

         CmbChampValeurCondition2.Enabled = False

         CmbChampValeurCondition2.Locked = True

      Case Else

         CmbChampValeurCondition2.Enabled = True

         CmbChampValeurCondition2.Locked = False

   End Select

Exit_CmbChampCondition2_AfterUpdate:

   Exit Sub

Err_CmbChampCondition2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampCondition2_AfterUpdate"

   Resume Exit_CmbChampCondition2_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampValeurCondition1_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_CmbChampValeurCondition1_BeforeUpdate

   Select Case CmbChampCondition1.Column(0)
      Case "<", "<=", "=", ">=", ">", "<>"

         On Error Resume Next

         Select Case LstChampSelectionner.Column(2)
            Case dbByte

               VarValeur = CByte(CmbChampValeurCondition1)

            Case dbChar, dbText

               VarValeur = CStr(CmbChampValeurCondition1)

            Case dbCurrency

               VarValeur = CCur(CmbChampValeurCondition1)

            Case dbDate, dbTime, dbTimeStamp

               VarValeur = CDate(CmbChampValeurCondition1)

            Case dbDecimal

               VarValeur = CDec(CmbChampValeurCondition1)

            Case dbDouble

               VarValeur = CDbl(CmbChampValeurCondition1)

            Case dbSingle

               VarValeur = CSng(CmbChampValeurCondition1)

            Case dbInteger

               VarValeur = CInt(CmbChampValeurCondition1)

            Case dbLong

               VarValeur = CLng(CmbChampValeurCondition1)

            Case Else

         End Select

         Select Case Err.Number
            Case 0

            Case Else

               Err.Clear

               MsgBox "La valeur saisie est invalide"

               Cancel = True

         End Select

         On Error GoTo Err_CmbChampValeurCondition1_BeforeUpdate

      Case Else

   End Select

Exit_CmbChampValeurCondition1_BeforeUpdate:

   Exit Sub

Err_CmbChampValeurCondition1_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampValeurCondition1_BeforeUpdate"

   Resume Exit_CmbChampValeurCondition1_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampValeurCondition2_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_CmbChampValeurCondition2_BeforeUpdate

   Select Case CmbChampCondition2.Column(0)
      Case "<", "<=", "=", ">=", ">", "<>"

         On Error Resume Next

         Select Case LstChampSelectionner.Column(2)
            Case dbByte

               VarValeur = CByte(CmbChampValeurCondition2)

            Case dbChar, dbText

               VarValeur = CStr(CmbChampValeurCondition2)

            Case dbCurrency

               VarValeur = CCur(CmbChampValeurCondition2)

            Case dbDate, dbTime, dbTimeStamp

               VarValeur = CDate(CmbChampValeurCondition2)

            Case dbDecimal

               VarValeur = CDec(CmbChampValeurCondition2)

            Case dbDouble

               VarValeur = CDbl(CmbChampValeurCondition2)

            Case dbSingle

               VarValeur = CSng(CmbChampValeurCondition2)

            Case dbInteger

               VarValeur = CInt(CmbChampValeurCondition2)

            Case dbLong

               VarValeur = CLng(CmbChampValeurCondition2)

            Case Else

         End Select

         Select Case Err.Number
            Case 0

            Case Else

               Err.Clear

               MsgBox "La valeur saisie est invalide"

               Cancel = True

         End Select

         On Error GoTo Err_CmbChampValeurCondition2_BeforeUpdate

      Case Else

   End Select

Exit_CmbChampValeurCondition2_BeforeUpdate:

   Exit Sub

Err_CmbChampValeurCondition2_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampValeurCondition2_BeforeUpdate"

   Resume Exit_CmbChampValeurCondition2_BeforeUpdate
End Sub
