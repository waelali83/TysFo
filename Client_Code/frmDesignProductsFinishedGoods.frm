VERSION 5.00
Begin VB.Form frmDesignProductsFinishedGoods 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Design Products Finished Goods"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDesignProductsFinishedGoods 
      Height          =   3960
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   4080
      Begin VB.TextBox txtSaturdayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "0.0"
         Top             =   2030
         Width           =   1080
      End
      Begin VB.TextBox txtMondayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "0.0"
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txtTuesdayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "0.0"
         Top             =   600
         Width           =   1080
      End
      Begin VB.TextBox txtThursdayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "0.0"
         Top             =   1320
         Width           =   1080
      End
      Begin VB.TextBox txtFridayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "0.0"
         Top             =   1680
         Width           =   1080
      End
      Begin VB.TextBox txtTotalWeightShipped 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   7
         Text            =   "0.0"
         Top             =   2400
         Width           =   1080
      End
      Begin VB.TextBox txtTotalBoxesReceived 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "0.0"
         Top             =   2760
         Width           =   1080
      End
      Begin VB.TextBox txtTotalWeightReceived 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   9
         Text            =   "0.0"
         Top             =   3120
         Width           =   1080
      End
      Begin VB.TextBox txtWednesdayEndingWeight 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "0.0"
         Top             =   960
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2970
         TabIndex        =   11
         Top             =   3510
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1890
         TabIndex        =   10
         Top             =   3510
         Width           =   1000
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   20
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   255
         TabIndex        =   19
         Top             =   645
         Width           =   1950
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   255
         TabIndex        =   18
         Top             =   975
         Width           =   2220
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   255
         TabIndex        =   17
         Top             =   1365
         Width           =   2010
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Friday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   255
         TabIndex        =   16
         Top             =   1695
         Width           =   1740
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday Ending Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   255
         TabIndex        =   15
         Top             =   2055
         Width           =   1965
      End
      Begin VB.Label lblTotalWEightShipped 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Weight Shipped"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   14
         Top             =   2445
         Width           =   1755
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Boxes Received"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   255
         TabIndex        =   13
         Top             =   2805
         Width           =   1770
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Weight Received"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   255
         TabIndex        =   12
         Top             =   3165
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmDesignProductsFinishedGoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmDesignProductsFinishedGoods.frm
'*File Description              : To get user input for printing inventory
'*                                details pertaining to plant 042 in Weekly
'*                                Sales Summary Report
'*Author                        : US Technology
'*Date Created                  : Jul-25-02
'*Date Last Modified            : Jul-01-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : CustomerStorageFunctions
'*Components Used               : USTriGrid
'*Functions Defined             : 1.ValidateTextBox
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Private fsCopyString                    As String
Private fbCanBeCopied                   As Boolean
 
Public Event DesignProducts(ByVal sDesignBeginWgt As String, _
                            ByVal sDesignEndWgt As String, _
                            ByVal sDesignTurn As String, _
                            ByVal sDesignAvgInv As String, _
                            ByVal sDesTotBoxRcd As String, _
                            ByVal sDesTotWgtRcd As String, _
                            ByVal sDesTotWgtShp As String, _
                            ByVal bOkCancelled As Boolean)

'*******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub cmdCancel_Click()

    RaiseEvent DesignProducts(0, _
                              0, _
                              0, _
                              0, _
                              0, _
                              0, _
                              0, _
                              False)
    Unload Me
    
End Sub

'*******************************************************************************
'* Functional Description   :   Assigns user input details and calculates values
'*                              to display in the Weekly Sales Summary Report
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub cmdOK_Click()

Dim dMonEndWgt As Double
Dim dTueEndWgt As Double
Dim dWedEndWgt As Double
Dim dThuEndWgt As Double
Dim dFriEndWgt As Double
Dim dSatEndWgt As Double
Dim dTotWgtShp As Double
Dim dTotBoxRcd As Double
Dim dTotWgtRcd As Double
Dim dTotEnd    As Double
Dim dGainLoss  As Double
Dim dTurn      As Double
Dim dAvgInv    As Double

    Select Case True
        Case txtMondayEndingWeight.Text = "" Or _
                Val(txtMondayEndingWeight.Text) = 0
                txtMondayEndingWeight.SetFocus
                txtMondayEndingWeight.SelStart = 0
                txtMondayEndingWeight.SelLength = Len(txtMondayEndingWeight.Text)
                GoTo ErrHandler
        Case txtTuesdayEndingWeight.Text = "" Or _
                Val(txtTuesdayEndingWeight.Text) = 0
                txtTuesdayEndingWeight.SetFocus
                txtTuesdayEndingWeight.SelStart = 0
                txtTuesdayEndingWeight.SelLength = Len(txtTuesdayEndingWeight.Text)
                GoTo ErrHandler
        Case txtWednesdayEndingWeight.Text = "" Or _
                Val(txtWednesdayEndingWeight.Text) = 0
                txtWednesdayEndingWeight.SetFocus
                txtWednesdayEndingWeight.SelStart = 0
                txtWednesdayEndingWeight.SelLength = Len _
                                                    (txtWednesdayEndingWeight.Text)
                GoTo ErrHandler
        Case txtThursdayEndingWeight.Text = "" Or _
                Val(txtThursdayEndingWeight.Text) = 0
                txtThursdayEndingWeight.SetFocus
                txtThursdayEndingWeight.SelStart = 0
                txtThursdayEndingWeight.SelLength = Len _
                                                    (txtThursdayEndingWeight.Text)
                GoTo ErrHandler
        Case txtFridayEndingWeight.Text = "" Or _
                Val(txtFridayEndingWeight.Text) = 0
                txtFridayEndingWeight.SetFocus
                txtFridayEndingWeight.SelStart = 0
                txtFridayEndingWeight.SelLength = Len(txtFridayEndingWeight.Text)
                GoTo ErrHandler
        Case txtSaturdayEndingWeight.Text = "" Or _
                Val(txtSaturdayEndingWeight.Text) = 0
                txtSaturdayEndingWeight.SetFocus
                txtSaturdayEndingWeight.SelStart = 0
                txtSaturdayEndingWeight.SelLength = Len _
                                                    (txtSaturdayEndingWeight.Text)
                GoTo ErrHandler
        Case txtTotalWeightShipped.Text = "" Or _
                Val(txtTotalWeightShipped.Text) = 0
                txtTotalWeightShipped.SetFocus
                txtTotalWeightShipped.SelStart = 0
                txtTotalWeightShipped.SelLength = Len(txtTotalWeightShipped.Text)
                GoTo ErrHandler
        Case txtTotalBoxesReceived.Text = "" Or _
                Val(txtTotalBoxesReceived.Text) = 0
                txtTotalBoxesReceived.SetFocus
                txtTotalBoxesReceived.SelStart = 0
                txtTotalBoxesReceived.SelLength = Len(txtTotalBoxesReceived.Text)
                GoTo ErrHandler
        Case txtTotalWeightReceived.Text = "" Or _
                Val(txtTotalWeightReceived.Text) = 0
                txtTotalWeightReceived.SetFocus
                txtTotalWeightReceived.SelStart = 0
                txtTotalWeightReceived.SelLength = Len(txtTotalWeightReceived.Text)
                GoTo ErrHandler
    End Select

    dMonEndWgt = Trim(txtMondayEndingWeight.Text)
    dTueEndWgt = Trim(txtTuesdayEndingWeight.Text)
    dWedEndWgt = Trim(txtWednesdayEndingWeight.Text)
    dThuEndWgt = Trim(txtThursdayEndingWeight.Text)
    dFriEndWgt = Trim(txtFridayEndingWeight.Text)
    dSatEndWgt = Trim(txtSaturdayEndingWeight.Text)
    dTotWgtShp = Trim(txtTotalWeightShipped.Text)
    dTotBoxRcd = Trim(txtTotalBoxesReceived.Text)
    dTotWgtRcd = Trim(txtTotalWeightReceived.Text)

    dTotEnd = dMonEndWgt + dTueEndWgt + dWedEndWgt + dThuEndWgt + dFriEndWgt _
              + dSatEndWgt
    dGainLoss = dSatEndWgt - dMonEndWgt
    
    If dTotWgtShp <> 0 Then
        dTurn = dTotEnd / dTotWgtShp
    Else
        dTurn = dTotEnd / 365
    End If
    
    dAvgInv = dTotEnd / 6
    
    RaiseEvent DesignProducts(dMonEndWgt, _
                              dSatEndWgt, _
                              dTurn, _
                              dAvgInv, _
                              dTotBoxRcd, _
                              dTotWgtRcd, _
                              dTotWgtShp, _
                              True)
    
    Unload Me
    
CleanUpAndExit:

    Exit Sub
    
ErrHandler:

    ShowErrorMsg "DesignPro001", Me.Caption, vbOKOnly

    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtMondayEndingWeight_GotFocus()
    
    txtMondayEndingWeight.SelStart = 0
    txtMondayEndingWeight.SelLength = Len(txtMondayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMondayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtMondayEndingWeight.Text, _
                                    txtMondayEndingWeight.SelText, _
                                    txtMondayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtMondayEndingWeight.Text = GetPastedText(txtMondayEndingWeight.Text, _
                                                txtMondayEndingWeight.SelText, _
                                                txtMondayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtMondayEndingWeight.Text, _
                                txtMondayEndingWeight.SelText, _
                                txtMondayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMondayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtMondayEndingWeight.Text)
    iSelLength = txtMondayEndingWeight.SelLength
    iSelStart = txtMondayEndingWeight.SelStart
    sSelText = txtMondayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtMondayEndingWeight.SelStart = iSelStart
            txtMondayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMondayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtMondayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMondayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMondayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtMondayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtMondayEndingWeight.Text), txtMondayEndingWeight.SelText, _
                                txtMondayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMondayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtMondayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtMondayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMondayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTuesdayEndingWeight_GotFocus()
    
    txtTuesdayEndingWeight.SelStart = 0
    txtTuesdayEndingWeight.SelLength = Len(txtTuesdayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTuesdayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtTuesdayEndingWeight.Text, _
                                    txtTuesdayEndingWeight.SelText, _
                                    txtTuesdayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtTuesdayEndingWeight.Text = GetPastedText(txtTuesdayEndingWeight.Text, _
                                                txtTuesdayEndingWeight.SelText, _
                                                txtTuesdayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtTuesdayEndingWeight.Text, _
                                txtTuesdayEndingWeight.SelText, _
                                txtTuesdayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTuesdayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtTuesdayEndingWeight.Text)
    iSelLength = txtTuesdayEndingWeight.SelLength
    iSelStart = txtTuesdayEndingWeight.SelStart
    sSelText = txtTuesdayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtTuesdayEndingWeight.SelStart = iSelStart
            txtTuesdayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTuesdayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtTuesdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtTuesdayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTuesdayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtTuesdayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtTuesdayEndingWeight.Text), txtTuesdayEndingWeight.SelText, _
                                txtTuesdayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTuesdayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtTuesdayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtTuesdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtTuesdayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtWednesdayEndingWeight_GotFocus()
    
    txtWednesdayEndingWeight.SelStart = 0
    txtWednesdayEndingWeight.SelLength = Len(txtWednesdayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWednesdayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtWednesdayEndingWeight.Text, _
                                    txtWednesdayEndingWeight.SelText, _
                                    txtWednesdayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtWednesdayEndingWeight.Text = GetPastedText(txtWednesdayEndingWeight.Text, _
                                                txtWednesdayEndingWeight.SelText, _
                                                txtWednesdayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtWednesdayEndingWeight.Text, _
                                txtWednesdayEndingWeight.SelText, _
                                txtWednesdayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWednesdayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtWednesdayEndingWeight.Text)
    iSelLength = txtWednesdayEndingWeight.SelLength
    iSelStart = txtWednesdayEndingWeight.SelStart
    sSelText = txtWednesdayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtWednesdayEndingWeight.SelStart = iSelStart
            txtWednesdayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWednesdayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtWednesdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtWednesdayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWednesdayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtWednesdayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtWednesdayEndingWeight.Text), txtWednesdayEndingWeight.SelText, _
                                txtWednesdayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWednesdayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtWednesdayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtWednesdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtWednesdayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtThursdayEndingWeight_GotFocus()
    
    txtThursdayEndingWeight.SelStart = 0
    txtThursdayEndingWeight.SelLength = Len(txtThursdayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtThursdayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtThursdayEndingWeight.Text, _
                                    txtThursdayEndingWeight.SelText, _
                                    txtThursdayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtThursdayEndingWeight.Text = GetPastedText(txtThursdayEndingWeight.Text, _
                                                txtThursdayEndingWeight.SelText, _
                                                txtThursdayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtThursdayEndingWeight.Text, _
                                txtThursdayEndingWeight.SelText, _
                                txtThursdayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtThursdayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtThursdayEndingWeight.Text)
    iSelLength = txtThursdayEndingWeight.SelLength
    iSelStart = txtThursdayEndingWeight.SelStart
    sSelText = txtThursdayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtThursdayEndingWeight.SelStart = iSelStart
            txtThursdayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtThursdayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtThursdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtThursdayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtThursdayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtThursdayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtThursdayEndingWeight.Text), txtThursdayEndingWeight.SelText, _
                                txtThursdayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtThursdayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtThursdayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtThursdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtThursdayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtFridayEndingWeight_GotFocus()
    
    txtFridayEndingWeight.SelStart = 0
    txtFridayEndingWeight.SelLength = Len(txtFridayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtFridayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtFridayEndingWeight.Text, _
                                    txtFridayEndingWeight.SelText, _
                                    txtFridayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtFridayEndingWeight.Text = GetPastedText(txtFridayEndingWeight.Text, _
                                                txtFridayEndingWeight.SelText, _
                                                txtFridayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtFridayEndingWeight.Text, _
                                txtFridayEndingWeight.SelText, _
                                txtFridayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtFridayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtFridayEndingWeight.Text)
    iSelLength = txtFridayEndingWeight.SelLength
    iSelStart = txtFridayEndingWeight.SelStart
    sSelText = txtFridayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtFridayEndingWeight.SelStart = iSelStart
            txtFridayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtFridayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtFridayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtFridayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtFridayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtFridayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtFridayEndingWeight.Text), txtFridayEndingWeight.SelText, _
                                txtFridayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtFridayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtFridayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtFridayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtFridayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtSaturdayEndingWeight_GotFocus()
    
    txtSaturdayEndingWeight.SelStart = 0
    txtSaturdayEndingWeight.SelLength = Len(txtSaturdayEndingWeight.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtSaturdayEndingWeight_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtSaturdayEndingWeight.Text, _
                                    txtSaturdayEndingWeight.SelText, _
                                    txtSaturdayEndingWeight.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtSaturdayEndingWeight.Text = GetPastedText(txtSaturdayEndingWeight.Text, _
                                                txtSaturdayEndingWeight.SelText, _
                                                txtSaturdayEndingWeight.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtSaturdayEndingWeight.Text, _
                                txtSaturdayEndingWeight.SelText, _
                                txtSaturdayEndingWeight.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtSaturdayEndingWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtSaturdayEndingWeight.Text)
    iSelLength = txtSaturdayEndingWeight.SelLength
    iSelStart = txtSaturdayEndingWeight.SelStart
    sSelText = txtSaturdayEndingWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtSaturdayEndingWeight.SelStart = iSelStart
            txtSaturdayEndingWeight.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtSaturdayEndingWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtSaturdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtSaturdayEndingWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtSaturdayEndingWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtSaturdayEndingWeight.Text

        If CheckNegPastedSelPart(Trim(txtSaturdayEndingWeight.Text), txtSaturdayEndingWeight.SelText, _
                                txtSaturdayEndingWeight.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtSaturdayEndingWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtSaturdayEndingWeight.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtSaturdayEndingWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtSaturdayEndingWeight.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTotalWeightShipped_GotFocus()
    
    txtTotalWeightShipped.SelStart = 0
    txtTotalWeightShipped.SelLength = Len(txtTotalWeightShipped.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightShipped_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtTotalWeightShipped.Text, _
                                    txtTotalWeightShipped.SelText, _
                                    txtTotalWeightShipped.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtTotalWeightShipped.Text = GetPastedText(txtTotalWeightShipped.Text, _
                                                txtTotalWeightShipped.SelText, _
                                                txtTotalWeightShipped.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtTotalWeightShipped.Text, _
                                txtTotalWeightShipped.SelText, _
                                txtTotalWeightShipped.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightShipped_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtTotalWeightShipped.Text)
    iSelLength = txtTotalWeightShipped.SelLength
    iSelStart = txtTotalWeightShipped.SelStart
    sSelText = txtTotalWeightShipped.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtTotalWeightShipped.SelStart = iSelStart
            txtTotalWeightShipped.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightShipped_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtTotalWeightShipped.Text)
    Call EvaluateNagativeNumber(sText)
    txtTotalWeightShipped.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightShipped_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtTotalWeightShipped.Text

        If CheckNegPastedSelPart(Trim(txtTotalWeightShipped.Text), txtTotalWeightShipped.SelText, _
                                txtTotalWeightShipped.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightShipped_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtTotalWeightShipped.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtTotalWeightShipped.Text)
    Call EvaluateNagativeNumber(sText)
    txtTotalWeightShipped.Text = sText

End Sub
'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTotalWeightReceived_GotFocus()
    
    txtTotalWeightReceived.SelStart = 0
    txtTotalWeightReceived.SelLength = Len(txtTotalWeightReceived.Text)

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightReceived_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtTotalWeightReceived.Text, _
                                    txtTotalWeightReceived.SelText, _
                                    txtTotalWeightReceived.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtTotalWeightReceived.Text = GetPastedText(txtTotalWeightReceived.Text, _
                                                txtTotalWeightReceived.SelText, _
                                                txtTotalWeightReceived.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtTotalWeightReceived.Text, _
                                txtTotalWeightReceived.SelText, _
                                txtTotalWeightReceived.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightReceived_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtTotalWeightReceived.Text)
    iSelLength = txtTotalWeightReceived.SelLength
    iSelStart = txtTotalWeightReceived.SelStart
    sSelText = txtTotalWeightReceived.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtTotalWeightReceived.SelStart = iSelStart
            txtTotalWeightReceived.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightReceived_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtTotalWeightReceived.Text)
    Call EvaluateNagativeNumber(sText)
    txtTotalWeightReceived.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightReceived_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtTotalWeightReceived.Text

        If CheckNegPastedSelPart(Trim(txtTotalWeightReceived.Text), txtTotalWeightReceived.SelText, _
                                txtTotalWeightReceived.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalWeightReceived_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtTotalWeightReceived.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtTotalWeightReceived.Text)
    Call EvaluateNagativeNumber(sText)
    txtTotalWeightReceived.Text = sText

End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTotalBoxesReceived_GotFocus()

    txtTotalBoxesReceived.SelStart = 0
    txtTotalBoxesReceived.SelLength = Len(txtTotalBoxesReceived.Text)
    
End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalBoxesReceived_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If InStr(Clipboard.GetText, ".") > 0 Then

            KeyCode = 0

        ElseIf Not IsNumeric(Clipboard.GetText) Then

            KeyCode = 0

        ElseIf Len(Clipboard.GetText) + Len(txtTotalBoxesReceived.Text) - txtTotalBoxesReceived.SelLength > 6 Then

            KeyCode = 0

        End If

    End If

End Sub


'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalBoxesReceived_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = True
        fsCopyString = txtTotalBoxesReceived.Text

        If InStr(Clipboard.GetText, ".") > 0 Then

            fbCanBeCopied = False

        ElseIf Not IsNumeric(Clipboard.GetText) Then

            fbCanBeCopied = False

        ElseIf Len(Clipboard.GetText) + Len(txtTotalBoxesReceived.Text) - txtTotalBoxesReceived.SelLength > 6 Then

            fbCanBeCopied = False

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalBoxesReceived_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtTotalBoxesReceived.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtTotalBoxesReceived.Text)
    Call EvaluateNagativeNumber(sText)
    txtTotalBoxesReceived.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   Allows only integers in this field.
'* Parameter Description    :   KeyAscii-Ascii value of key pressed.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTotalBoxesReceived_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> Asc("-") Then
        
        If KeyAscii = 22 Then
        
            If InStr(Clipboard.GetText, ".") > 0 Then
    
                KeyAscii = 0
    
            ElseIf Not IsNumeric(Clipboard.GetText) Then
    
                KeyAscii = 0
    
            ElseIf Len(Clipboard.GetText) + Len(txtTotalBoxesReceived.Text) _
                   - txtTotalBoxesReceived.SelLength > 6 Then
    
                KeyAscii = 0
    
            End If
        
        Else
            If InStr(txtTotalBoxesReceived.Text, "-") > 0 Then
        
                ValidateForIntegers KeyAscii
                
            ElseIf Len(txtTotalBoxesReceived.Text) = 6 And KeyAscii > 31 Then
            
                KeyAscii = 0
                
            Else
                
                ValidateForIntegers KeyAscii
            
            End If
                
        End If
        
    ElseIf txtTotalBoxesReceived.SelStart <> 0 Then
        
        KeyAscii = 0
        
    End If

End Sub
