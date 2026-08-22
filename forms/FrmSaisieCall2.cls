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
'***    Declaration De Procedure Privée                                                       ***
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

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               On Error Resume Next

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = CadSocCallQst2

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               Select Case CtlActive.Name
                  Case "CmdAfficherCall1", "CmdAfficherCall3"

                     CmdImprimer.SetFocus

                     Set CtlActive = CmdImprimer

                  Case Else

               End Select
               CmdAfficherCall1.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL1", "SocCallSocNum=" & SocNum)), True, False)

               CmdAfficherCall3.Visible = IIf(IsNull(DLookup("SocCallNum", "TBLSOCIETESCALL3", "SocCallSocNum=" & SocNum)), True, False)

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

               Select Case IsNull(SocCallNum)
                  Case True

                  Case False

                     CadSocCallQst3.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst3.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst4.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst4.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst5.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst5.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst6.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst6.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst7.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst7.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst8.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst8.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst9.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst9.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst10.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst10.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst11.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst11.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst12.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst12.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst13.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst13.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst14.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst14.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst15.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst15.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst16.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst16.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst17.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst17.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst18.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst18.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst19.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst19.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst20.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst20.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst21.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst21.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst22.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst22.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst23.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst23.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst24.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst24.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst25.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst25.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst26.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst26.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst27.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst27.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst28.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst28.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst29.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst29.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst20.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst20.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst21.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst21.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst22.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst22.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst23.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst23.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst24.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst24.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst25.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst25.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst26.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst26.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst27.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst27.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst28.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst28.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst29.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst29.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst30.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst30.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst31.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst31.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst32.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst32.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst33.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst33.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst34.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst34.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst35.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst35.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst36.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst36.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst37.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst37.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst38.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst38.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst39.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst39.Locked = Not (CadSocCallQst2 = 1)

                     CadSocCallQst79.Enabled = (CadSocCallQst2 = 1)

                     CadSocCallQst79.Locked = Not (CadSocCallQst2 = 1)

                     Select Case CadSocCallQst2
                        Case 0, 1

                        Case 2

                           CadSocCallQst3 = 0

                           CadSocCallQst4 = 0

                           CadSocCallQst5 = 0

                           CadSocCallQst6 = 0

                           CadSocCallQst7 = 0

                           CadSocCallQst8 = 0

                           CadSocCallQst9 = 0

                           CadSocCallQst10 = 0

                           CadSocCallQst11 = 0

                           CadSocCallQst12 = 0

                           CadSocCallQst13 = 0

                           CadSocCallQst14 = 0

                           CadSocCallQst15 = 0

                           CadSocCallQst16 = 0

                           CadSocCallQst17 = 0

                           CadSocCallQst18 = 0

                           CadSocCallQst19 = 0

                           CadSocCallQst20 = 0

                           CadSocCallQst21 = 0

                           CadSocCallQst22 = 0

                           CadSocCallQst23 = 0

                           CadSocCallQst24 = 0

                           CadSocCallQst25 = 0

                           CadSocCallQst26 = 0

                           CadSocCallQst27 = 0

                           CadSocCallQst28 = 0

                           CadSocCallQst29 = 0

                           CadSocCallQst30 = 0

                           CadSocCallQst31 = 0

                           CadSocCallQst32 = 0

                           CadSocCallQst33 = 0

                           CadSocCallQst34 = 0

                           CadSocCallQst35 = 0

                           CadSocCallQst36 = 0

                           CadSocCallQst37 = 0

                           CadSocCallQst38 = 0

                           CadSocCallQst39 = 0

                     End Select

                     CadSocCallQst40.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst40.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst41.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst41.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst42.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst42.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst43.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst43.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst44.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst44.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst45.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst45.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst46.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst46.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst47.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst47.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst48.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst48.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst49.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst49.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst50.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst50.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst51.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst51.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst52.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst52.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst53.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst53.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst54.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst54.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst55.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst55.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst56.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst56.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst57.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst57.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst58.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst58.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst59.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst59.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst60.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst60.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst61.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst61.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst62.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst62.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst63.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst63.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst64.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst64.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst65.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst65.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst66.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst66.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst67.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst67.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst68.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst68.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst69.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst69.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst70.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst70.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst71.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst71.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst72.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst72.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst73.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst73.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst74.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst74.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst75.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst75.Locked = Not (CadSocCallQst39 = 1)

                     CadSocCallQst76.Enabled = (CadSocCallQst39 = 1)

                     CadSocCallQst76.Locked = Not (CadSocCallQst39 = 1)

                     Select Case CadSocCallQst39
                        Case 0, 1

                        Case 2

                           CadSocCallQst40 = 0

                           CadSocCallQst41 = 0

                           CadSocCallQst42 = 0

                           CadSocCallQst43 = 0

                           CadSocCallQst44 = 0

                           CadSocCallQst45 = 0

                           CadSocCallQst46 = 0

                           CadSocCallQst47 = 0

                           CadSocCallQst48 = 0

                           CadSocCallQst49 = 0

                           CadSocCallQst50 = 0

                           CadSocCallQst51 = 0

                           CadSocCallQst52 = 0

                           CadSocCallQst53 = 0

                           CadSocCallQst54 = 0

                           CadSocCallQst55 = 0

                           CadSocCallQst56 = 0

                           CadSocCallQst57 = 0

                           CadSocCallQst58 = 0

                           CadSocCallQst59 = 0

                           CadSocCallQst60 = 0

                           CadSocCallQst61 = 0

                           CadSocCallQst62 = 0

                           CadSocCallQst63 = 0

                           CadSocCallQst64 = 0

                           CadSocCallQst65 = 0

                           CadSocCallQst66 = 0

                           CadSocCallQst67 = 0

                           CadSocCallQst68 = 0

                           CadSocCallQst69 = 0

                           CadSocCallQst70 = 0

                           CadSocCallQst71 = 0

                           CadSocCallQst72 = 0

                           CadSocCallQst73 = 0

                           CadSocCallQst74 = 0

                           CadSocCallQst75 = 0

                           CadSocCallQst76 = 0

                     End Select

                     TxtSocCallQst77.Enabled = (CadSocCallQst76 = 1)

                     TxtSocCallQst77.Locked = Not (CadSocCallQst76 = 1)

                     CmbSocCallQst78.Enabled = (CadSocCallQst76 = 1)

                     CmbSocCallQst78.Locked = Not (CadSocCallQst76 = 1)

                     Select Case CadSocCallQst76
                        Case 0, 1

                        Case 2

                           TxtSocCallQst77 = Null

                           CmbSocCallQst78 = Null

                     End Select

                     CadSocCallQst80.Enabled = (CadSocCallQst79 = 2)

                     CadSocCallQst80.Locked = Not (CadSocCallQst79 = 2)

                     Select Case CadSocCallQst79
                        Case 0, 1

                           CadSocCallQst80 = 0

                           TxtSocCallQst81 = Null

                           CmbSocCallQst82 = Null

                        Case 2

                     End Select

                     TxtSocCallQst81.Enabled = (CadSocCallQst80 = 1)

                     TxtSocCallQst81.Locked = Not (CadSocCallQst80 = 1)

                     CmbSocCallQst82.Enabled = (CadSocCallQst80 = 1)

                     CmbSocCallQst82.Locked = Not (CadSocCallQst80 = 1)

                     Select Case CadSocCallQst80
                        Case 0, 1

                        Case 2

                           TxtSocCallQst81 = Null

                           CmbSocCallQst82 = Null

                     End Select

                     CadSocCallQst129.Enabled = (CadSocCallQst128 = 2)

                     CadSocCallQst129.Locked = Not (CadSocCallQst128 = 2)

                     Select Case CadSocCallQst128
                        Case 0, 1

                           CadSocCallQst129 = 0

                           TxtSocCallQst130 = Null

                           CmbSocCallQst131 = Null

                        Case 2

                     End Select

                     TxtSocCallQst130.Enabled = (CadSocCallQst129 = 1)

                     TxtSocCallQst130.Locked = Not (CadSocCallQst129 = 1)

                     CmbSocCallQst131.Enabled = (CadSocCallQst129 = 1)

                     CmbSocCallQst131.Locked = Not (CadSocCallQst129 = 1)

                     Select Case CadSocCallQst129
                        Case 0, 1

                        Case 2

                           TxtSocCallQst130 = Null

                           CmbSocCallQst131 = Null

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
'***    Declaration D'Evenement Privée                                                        ***
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

   DoCmd.RunCommand acCmdWindowCascade

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='INTERFACE'")
      Case "GESTION"

         Me.MenuBar = "Menu Coeur De Pro 2003"

         Me.Toolbar = "Outil Coeur De Pro 2003"

         CmdAfficherContacts.Visible = True

         CmdAfficherSocietes.Visible = True

      Case Else

         Me.MenuBar = "Menu Call Coeur De Pro 2003"

         Me.Toolbar = "Outil Call Coeur De Pro 2003"

         CmdAfficherContacts.Visible = False

         CmdAfficherSocietes.Visible = False

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
Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

'   Select Case CadSocCallQst1
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst1").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst1.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select

   Select Case CadSocCallQst2
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst2").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst2.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst83
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst83").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst83.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst86
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst86").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst86.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst89
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst89").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst89.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst92
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst92").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst92.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst95
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst95").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst95.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst98
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst98").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst98.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst101
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst101").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst101.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst104
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst104").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst104.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst107
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst107").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst107.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst110
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst110").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst110.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst112
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst112").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst112.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst114
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst114").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst114.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst116
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst116").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst116.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst118
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst118").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst118.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst120
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst120").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst120.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst122
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst122").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst122.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst124
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst124").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst124.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst126
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst126").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst126.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

   Select Case CadSocCallQst128
      Case 0

         Cancel = True

         MsgBox Me.Controls("EtiSocCallQst128").Caption & vbCrLf & "n'a pas ete repondu !"

         CadSocCallQst128.SetFocus

         GoTo Exit_Form_BeforeUpdate

      Case Else

   End Select

'   Select Case CadSocCallQst132
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst132").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst132.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select
'
'   Select Case CadSocCallQst133
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst133").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst133.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select
'
'   Select Case CadSocCallQst134
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst134").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst134.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select
'
'   Select Case CadSocCallQst135
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst135").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst135.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select
'
'   Select Case CadSocCallQst136
'      Case 0
'
'         Cancel = True
'
'         MsgBox Me.Controls("EtiSocCallQst136").Caption & vbCrLf & "n'a pas ete repondu !"
'
'         CadSocCallQst136.SetFocus
'
'         GoTo Exit_Form_BeforeUpdate
'
'      Case Else
'
'   End Select

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
Private Sub Form_Current()

   On Error GoTo Err_Form_Current

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Select Case Err.Number
            Case 2501, 2759

            Case Else

               MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdFermer_Click"

         End Select

         Err.Clear

         GoTo Exit_CmdFermer_Click

   End Select

   DoCmd.Close acForm, Me.Name, acSavePrompt

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

   On Error GoTo Err_CmdRechercher_Click

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.PrintOut acSelection

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
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
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherContacts_Click

   Select Case IsNull(DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum))
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheContacts", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheContacts.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "ConNum=" & DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum)

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

               DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConNum=" & DLookup("ConNum", "SelSocietesSTAVillesContactsTAInscrit", "SocNum=" & SocNum), acFormEdit

         End Select

   End Select

Exit_CmdAfficherContacts_Click:

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
Private Sub CmdAfficherSocietes_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case ETATFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               On Error Resume Next

               DoCmd.ApplyFilter , "SocNum=" & SocNum

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

               DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & SocNum, acFormEdit

         End Select

   End Select

Exit_CmdAfficherSocietes_Click:

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
Private Sub CmdAfficherCall3_Click()

   On Error GoTo Err_CmdAfficherCall3_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall3", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall3", acNormal, , "SocNum=" & SocNum, , acWindowNormal

   End Select

Exit_CmdAfficherCall3_Click:

   Exit Sub

Err_CmdAfficherCall3_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCall3_Click"

   Resume Exit_CmdAfficherCall3_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherCall1_Click()

   On Error GoTo Err_CmdAfficherCall1_Click

   Select Case IsNull(SocNum)
      Case True

      Case False

         DoCmd.Close acForm, "FrmSaisieCall1", acSavePrompt

         DoCmd.OpenForm "FrmSaisieCall1", acNormal, , "SocNum=" & SocNum, , acWindowNormal

   End Select

Exit_CmdAfficherCall1_Click:

   Exit Sub

Err_CmdAfficherCall1_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherCall1_Click"

   Resume Exit_CmdAfficherCall1_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst2_AfterUpdate()

   On Error GoTo Err_CadSocCallQst2_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst2
      Case 0

      Case 1

         CadSocCallQst3.SetFocus

      Case 2

         CadSocCallQst83.SetFocus

   End Select

Exit_CadSocCallQst2_AfterUpdate:

   Exit Sub

Err_CadSocCallQst2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst2_AfterUpdate"

   Resume Exit_CadSocCallQst2_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst39_AfterUpdate()

   On Error GoTo Err_CadSocCallQst39_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst39
      Case 0

      Case 1

         CadSocCallQst40.SetFocus

      Case 2

         CadSocCallQst83.SetFocus

   End Select

Exit_CadSocCallQst39_AfterUpdate:

   Exit Sub

Err_CadSocCallQst39_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst39_AfterUpdate"

   Resume Exit_CadSocCallQst39_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst76_AfterUpdate()

   On Error GoTo Err_CadSocCallQst76_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst76
      Case 0

      Case 1

         TxtSocCallQst77.SetFocus

      Case 2

         CadSocCallQst79.SetFocus

   End Select

Exit_CadSocCallQst76_AfterUpdate:

   Exit Sub

Err_CadSocCallQst76_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst76_AfterUpdate"

   Resume Exit_CadSocCallQst76_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst79_AfterUpdate()

   On Error GoTo Err_CadSocCallQst79_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst79
      Case 0

      Case 1

         CadSocCallQst83.SetFocus

      Case 2

         CadSocCallQst80.SetFocus

   End Select

Exit_CadSocCallQst79_AfterUpdate:

   Exit Sub

Err_CadSocCallQst79_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst79_AfterUpdate"

   Resume Exit_CadSocCallQst79_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst80_AfterUpdate()

   On Error GoTo Err_CadSocCallQst80_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst80
      Case 0

      Case 1

         TxtSocCallQst81.SetFocus

      Case 2

         CadSocCallQst83.SetFocus

   End Select

Exit_CadSocCallQst80_AfterUpdate:

   Exit Sub

Err_CadSocCallQst80_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst80_AfterUpdate"

   Resume Exit_CadSocCallQst80_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst128_AfterUpdate()

   On Error GoTo Err_CadSocCallQst128_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst128
      Case 0

      Case 1

'         CadSocCallQst132.SetFocus

      Case 2

         CadSocCallQst129.SetFocus

   End Select

Exit_CadSocCallQst128_AfterUpdate:

   Exit Sub

Err_CadSocCallQst128_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst128_AfterUpdate"

   Resume Exit_CadSocCallQst128_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CadSocCallQst129_AfterUpdate()

   On Error GoTo Err_CadSocCallQst129_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CadSocCallQst129
      Case 0

      Case 1

         TxtSocCallQst130.SetFocus

      Case 2

         'CadSocCallQst132.SetFocus

   End Select

Exit_CadSocCallQst129_AfterUpdate:

   Exit Sub

Err_CadSocCallQst129_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CadSocCallQst129_AfterUpdate"

   Resume Exit_CadSocCallQst129_AfterUpdate
End Sub
