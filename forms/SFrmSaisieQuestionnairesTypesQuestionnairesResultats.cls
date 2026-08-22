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
'***    Declaration De Procedure Privée                                                        ***
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
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)
   Dim CtlActive As Access.Control
   Dim VarValeurTableau As Variant
   Dim VarValeur As Variant
   Dim StrValeurListe As String
   Dim LngValeurCompteur As Long
   Dim IntValeurMin As Integer
   Dim IntValeurMax As Integer

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               On Error Resume Next

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Set CtlActive = CmbSocQstResValeur

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               StrValeurListe = "'Resultats Valeur';"

               Select Case SocQstTypIDTypNum
                  Case vbBoolean

                     StrValeurListe = StrValeurListe & "'Vrai';'Faux'"

                     CmbSocQstResValeur.InputMask = vbNullString

                     CmbSocQstResValeur.Format = vbNullString

                     CmbSocQstResValeur.RowSource = StrValeurListe

                  Case vbByte, vbInteger

                     Select Case SocQstTypIDTypNum
                        Case vbByte

                           IntValeurMin = 0

                           IntValeurMax = 255

                        Case vbInteger

                           IntValeurMin = -32768

                           IntValeurMax = 32767

                     End Select

                     Select Case Nz(SocQstTypIDMax, IntValeurMax) - Nz(SocQstTypIDMin, IntValeurMin)
                        Case Is < 50

                           For LngValeurCompteur = Nz(SocQstTypIDMin, IntValeurMin) To Nz(SocQstTypIDMax, IntValeurMax)

                              StrValeurListe = StrValeurListe & LngValeurCompteur & ";"

                           Next

                           CmbSocQstResValeur.RowSource = StrValeurListe

                        Case Else

                           CmbSocQstResValeur.RowSource = vbNullString

                     End Select

                     StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                     CmbSocQstResValeur.InputMask = vbNullString

                     CmbSocQstResValeur.Format = vbNullString

                     CmbSocQstResValeur.RowSource = StrValeurListe

                  Case vbLong, vbSingle, vbDouble, vbCurrency

                     CmbSocQstResValeur.InputMask = vbNullString

                     CmbSocQstResValeur.Format = vbNullString

                     CmbSocQstResValeur.RowSource = vbNullString

                  Case vbDate

                     CmbSocQstResValeur.InputMask = "99/99/0000;;_"

                     CmbSocQstResValeur.Format = "Date, abrégé"

                     CmbSocQstResValeur.RowSource = vbNullString

                  Case vbString

                     Select Case IsNull(SocQstTypIDMin)
                        Case True

                           CmbSocQstResValeur.InputMask = Nz(SocQstTypIDMax, vbNullString)

                           CmbSocQstResValeur.Format = vbNullString

                           CmbSocQstResValeur.RowSource = vbNullString

                        Case False

                           VarValeurTableau = Split(SocQstTypIDMin, ";")

                           For Each VarValeur In VarValeurTableau

                              StrValeurListe = StrValeurListe & VarValeur & ";"

                           Next

                           StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                           CmbSocQstResValeur.InputMask = Nz(SocQstTypIDMax, vbNullString)

                           CmbSocQstResValeur.Format = vbNullString

                           CmbSocQstResValeur.RowSource = StrValeurListe

                     End Select

               End Select

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
'***    Declaration d'Evenement Privée                                                        ***
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

   On Error Resume Next

   Select Case IsObject(Me.Parent)
      Case True

      Case False

         Cancel = True

   End Select

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

         Cancel = True

   End Select

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
Private Sub Form_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant
   Dim VarMax As Variant
   Dim VarMin As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case SocQstTypIDObligatoire = False And (IsNull(CmbSocQstResValeur) Or CmbSocQstResValeur = vbNullString)
      Case True

         SocQstResValeur = Null

      Case False

         On Error Resume Next

         Select Case SocQstTypIDTypNum
            Case vbBoolean

               Select Case CmbSocQstResValeur
                  Case "Vrai"

                  Case "Faux"

                  Case Else

                     Err.Raise 13

               End Select

            Case vbByte

               VarValeur = CByte(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CByte(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CByte(SocQstTypIDMin)

               End Select

            Case vbInteger

               VarValeur = CInt(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CInt(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CInt(SocQstTypIDMin)

               End Select

            Case vbLong

               VarValeur = CLng(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CLng(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CLng(SocQstTypIDMin)

               End Select

            Case vbSingle

               VarValeur = CSng(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CSng(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CSng(SocQstTypIDMin)

               End Select

            Case vbDouble

               VarValeur = CDbl(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDbl(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDbl(SocQstTypIDMin)

               End Select

            Case vbCurrency

               VarValeur = CCur(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CCur(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CCur(SocQstTypIDMin)

               End Select

            Case vbDate

               VarValeur = CDate(CmbSocQstResValeur)

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDate(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDate(SocQstTypIDMin)

               End Select

            Case vbString

               Select Case Not IsNull(SocQstTypIDMin) And IsNull(SocQstTypIDMax)
                  Case True

                     VarValeur = UCase(CStr(CmbSocQstResValeur))

                  Case False

                     VarValeur = CStr(CmbSocQstResValeur)

               End Select

               Select Case VarValeur
                  Case vbNullString

                     SocQstResValeur = Null

                  Case Else

                     SocQstResValeur = VarValeur

               End Select

               Select Case IsNull(SocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CStr(SocQstTypIDMax)

               End Select

               Select Case IsNull(SocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CStr(SocQstTypIDMin)

               End Select

         End Select

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Cancel = True

               MsgBox "Le Champ Questionnaire Resultat Valeur Est Invalide !", , Me.Caption

               CmbSocQstResValeur.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select

         On Error GoTo Err_Form_BeforeUpdate

         Select Case SocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case VarMin <= VarValeur
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Doit Etre Superieur Ou Egal A " & VarMin & " !", , Me.Caption

                           CmbSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

               Select Case IsNull(VarMax)
                  Case True

                  Case False

                     Select Case VarMax >= VarValeur
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Doit Etre Inferieur Ou Egal A " & VarMax & " !", , Me.Caption

                           CmbSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

            Case vbString

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case CmbSocQstResValeur.ListIndex
                        Case Is >= 0

                        Case Else

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Est Invalide !", , Me.Caption

                           CmbSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

         End Select

   End Select

   SocQstResSocNum = Me.Parent.CmbSocQstSocNum

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
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

   CmbSocQstResValeur.SetFocus

   Select Case IsNull(SocQstTypIDMessage)
      Case True

      Case False

         MsgBox SocQstTypIDMessage, , Me.Caption

   End Select

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
Private Sub CmbSocQstResValeur_DblClick(Cancel As Integer)

   On Error GoTo Err_CmbSocQstResValeur_DblClick

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acDialog, "FrmSaisieQuestionnairesTypes:SFrmSaisieQuestionnairesTypesQuestionnairesResultats:CmbSocQstResValeur:"

Exit_CmbSocQstResValeur_DblClick:

   Exit Sub

Err_CmbSocQstResValeur_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocQstResValeur_DblClick"

   Resume Exit_CmbSocQstResValeur_DblClick
End Sub
