VERSION 5.00
Begin VB.Form frmDailyProduction 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Production Report"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDailyProduction 
      Height          =   2520
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   3675
      Begin VB.ComboBox cboCustomerType 
         Height          =   315
         ItemData        =   "frmDailyProduction.frx":0000
         Left            =   1800
         List            =   "frmDailyProduction.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCustomerNo 
         Height          =   285
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   2
         Top             =   600
         Width           =   1305
      End
      Begin VB.ComboBox cboProdBreakout 
         Height          =   315
         ItemData        =   "frmDailyProduction.frx":0004
         Left            =   1800
         List            =   "frmDailyProduction.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox cboReportShift 
         Height          =   315
         ItemData        =   "frmDailyProduction.frx":0008
         Left            =   1800
         List            =   "frmDailyProduction.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboDivisionCode 
         Height          =   315
         ItemData        =   "frmDailyProduction.frx":000C
         Left            =   1800
         List            =   "frmDailyProduction.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1500
         TabIndex        =   6
         Top             =   2055
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2565
         TabIndex        =   7
         Top             =   2055
         Width           =   1000
      End
      Begin VB.Label lblCustomerType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         TabIndex        =   12
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label lblReportShift 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Shift"
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
         TabIndex        =   11
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label lblDivisionCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Division Code"
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
         TabIndex        =   10
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label lblCustomerNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer No."
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
         TabIndex        =   9
         Top             =   645
         Width           =   1155
      End
      Begin VB.Label lblProductBreakout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Breakout"
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
         TabIndex        =   8
         Top             =   1725
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDailyProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmDailyProduction.frm
'*File Description              : Get the  user input for printing Daily
'*                                Production Report
'*Author                        : US Technology
'*Date Created                  : Sep-03-02
'*Date Last Modified            : Dec-11-02
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : Inventory Functions
'*Components Used               : None
'*Functions Defined             : LoadDivisionCodes
'*                                ValidateCustomer
'*                                PrintDailyProductionRpt
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)    Change Description    Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Private fsPetDate           As String

'Not to unload the form if customer is not valid

Private fbIsValid           As Boolean

'To check whether cancel button is clicked or not

Private fbCancelled         As Boolean

'to check the successful loading of the form

Private flLoadSuccess       As Long

'******************************************************************************
'* Functional Description   :   The entry point to this form from modPrint.
'*                              Invokes Form_Load().
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean to check whether form loading is
'*                              successful
'******************************************************************************

Public Function LoadDailyProductionReport() As Boolean
    
    On Error GoTo ErrHandler
    
    LoadDailyProductionReport = True
    
    'call the form load
    
    Me.Show vbModal
    
    'if there is any problem in loading the form
    
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    
    Exit Function

ErrHandler:
    
    'Status text message reset.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    LoadDailyProductionReport = False

End Function

'******************************************************************************
'* Functional Description   :   Loads the Daily Production Report Form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()

Dim objWeeklyFrzRecapRpt        As Object
Dim lResult                     As Long
Dim lDivisionResult             As Long
Dim objDailyProductionRpt       As Object
Dim rsDivisionCodes             As ADODB.Recordset

'Added by TCS to get MultiShiftInd of the plant
Dim fMultishiftInd As String

    On Error GoTo ErrHandler
    
    fbCancelled = True
    flLoadSuccess = 0
    lResult = 0
    
    Set objWeeklyFrzRecapRpt = _
            CreateObject("InventoryFunctions.WeeklyFrzRecapRpt")

    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    
    'Change of status bar message in case of Loading from the database.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    'Get the Pet Date of the plant.
    
    lResult = objWeeklyFrzRecapRpt.GetPetDate(gsPlantCode, fsPetDate)
    
    'resets the mouse pointer
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'If there is any error in getting the pet date
    
    If lResult <> 0 Then GoTo ErrHandler
    
    'Load all the Division Codes from the Types table.
    
    Set objDailyProductionRpt = _
            CreateObject("InventoryFunctions.DailyProductionRpt")
    
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    
    lDivisionResult = objDailyProductionRpt.GetAllDivisionCodes _
                                                        (rsDivisionCodes)
    
    '******************************************************
    '                       Added by TCS
    ' *****************************************************
    lResult = objDailyProductionRpt.GetMultiShiftInd(gsPlantCode, fMultishiftInd)
    If lResult <> 0 Then GoTo ErrHandler
    
    '******************************************************
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'If there is any problem in getting division codes
    
    If lDivisionResult <> 0 Then GoTo ErrHandler
    If lResult <> 0 Then GoTo ErrHandler
    
    'resets the status text
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    'Loads the division code combo with division codes retrieved from database
    
    If Not rsDivisionCodes.EOF Then
    
        cboDivisionCode.AddItem ""
        cboDivisionCode.AddItem "A - Combine All"
        rsDivisionCodes.MoveFirst
        
        While Not rsDivisionCodes.EOF
        
            cboDivisionCode.AddItem _
                    rsDivisionCodes.Fields("TYPE_CODE").Value & " - " & _
                    rsDivisionCodes.Fields("TYPE_SHORT_DESC").Value
            rsDivisionCodes.MoveNext
        
        Wend
        
        cboDivisionCode.ListIndex = 1
    
    End If
    
    
    
    ' Load the Report Shift combo
    
    cboReportShift.AddItem "A"
    cboReportShift.AddItem "B"
    cboReportShift.ListIndex = 0
    
    'Load the Prod Breakout combo
    
    cboProdBreakout.Clear
    cboProdBreakout.AddItem "Y"
    cboProdBreakout.AddItem "N"
    cboProdBreakout.ListIndex = 0
    
        
    If fMultishiftInd = "N" Then
          cboProdBreakout.ListIndex = 1
    Else
          cboProdBreakout.ListIndex = 0
    End If
    
    ' Load the Customer Type Combo
    
    cboCustomerType.AddItem "S - Ship-to"
    cboCustomerType.AddItem "B - Bill-to"
    cboCustomerType.AddItem "C - Corporate"
    cboCustomerType.ListIndex = -1
    
CleanUpAndExit:
    
    'Cleaning up of all the objects and recordset
    Set objWeeklyFrzRecapRpt = Nothing
    Set objDailyProductionRpt = Nothing
    Set rsDivisionCodes = Nothing
    Exit Sub
    
ErrHandler:
    
    If lDivisionResult <> 0 Then
        
        'Error message display in case of Server side errors
        'Error in retrieving division codes
        
        gcolErrMsg.Add lDivisionResult
        ShowErrorMsg "DailyPro001", Me.Caption, vbOKOnly, gcolErrMsg
        flLoadSuccess = lDivisionResult
        fbCancelled = True
           
    ElseIf lResult <> 0 Then
        
        'Error message display in case of Server side errors
        'Error in retrieving Pet Date
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "DailyPro002", Me.Caption, vbOKOnly, gcolErrMsg
        flLoadSuccess = lResult
        
    Else
        
        'Error message display in case of VB errors in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "DailyPro003", Me.Caption, vbOKOnly, gcolErrMsg
        flLoadSuccess = Err.Number
        
    End If
    
    'Cleaning up of all the objects and recordset
    
    Set objWeeklyFrzRecapRpt = Nothing
    Set objDailyProductionRpt = Nothing
    Set rsDivisionCodes = Nothing
    
    'calls the form terminate when there is any problem in loading form
    
    Form_Terminate
    
End Sub

'******************************************************************************
'* Functional Description   :   Allow the Customer No input based on the
'*                              Division Code selected.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cboDivisionCode_Click()
    
    On Error GoTo ErrHandler
    
    If cboDivisionCode.ListIndex <> 0 Then
    
        txtCustomerNo.Text = "0"
        txtCustomerNo.Enabled = False
        cboCustomerType.ListIndex = -1
        cboCustomerType.Enabled = False
    
    Else
    
        txtCustomerNo.Enabled = True
        
    End If

CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:

    'Error message display in case of VB errors in this function.
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyPro011", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Unload Pop up Window
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Terminate()

    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the status bar of main menu as false is
'*                              forms other than the present form is open
'* Parameter Description    :   Status of Cancel button.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Unload(Cancel As Integer)

    If Forms.Count > 3 Then
    
        frmMainMenu.sbMainMenu.Visible = False
        frmMainMenu.mnuViewStatusBar.Checked = False
        
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Calls ReturnKeyForCapsAndNumerics Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtCustomerNo, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Set the default value of Customer No as zero.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_LostFocus()
    
    On Error GoTo ErrHandler
    
    If Me.ActiveControl Is cmdCancel Then Exit Sub
    
    'if customer no. text box is empty text box = 0
    
    If Trim(txtCustomerNo.Text) = "" Then txtCustomerNo.Text = "0"
    
    If Me.ActiveControl Is cboDivisionCode Then Exit Sub
    
    ' Set the Customer No field to zero if no Division Code is
    ' selected and if no input is given for the field.
    
    If cboDivisionCode.ListIndex = 0 Then
    
        If Trim(txtCustomerNo.Text) = "" Then
        
            txtCustomerNo.Text = "0"
            
        End If
        
    End If
    
    ' Enable the Customer Type field if the Customer No entered is not Zero.
    
    If Trim(txtCustomerNo.Text) <> "0" Then
    
        cboCustomerType.Enabled = True
        
        If cboCustomerType.ListIndex = -1 Then cboCustomerType.ListIndex = 0
    
    Else
        
        cboCustomerType.Enabled = False
        cboCustomerType.ListIndex = -1
    
    End If
 
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:
    
    'Error message display in case of VB errors in this function.
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyPro012", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Set the default value of Customer Type as S
'*                              if the Customer No is entered and no Customer
'*                              Type is selected.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    On Error GoTo ErrHandler
    
    ReturnValueForCapsAndNumerics txtCustomerNo, sOrigValue
    txtCustomerNo = sOrigValue
    
    If Trim(txtCustomerNo.Text) <> "" Then
    
        If Trim(txtCustomerNo.Text) <> "0" Then
        
            cboCustomerType.Enabled = True
            
            If cboCustomerType.ListIndex = -1 Then
                
                cboCustomerType.ListIndex = 0
             
            End If
            
        Else
        
            cboCustomerType.Enabled = False
            cboCustomerType.ListIndex = -1
        
        End If
        
    Else
        
        cboCustomerType.Enabled = False
        cboCustomerType.ListIndex = -1
        txtCustomerNo.SetFocus
        
    End If
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:
    
    'Error message display in case of VB errors in this function.
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyPro013", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Select the text in the Textbox
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_GotFocus()

    txtCustomerNo.SelStart = 0
    txtCustomerNo.SelLength = Len(txtCustomerNo.Text)

End Sub

'******************************************************************************
'* Functional Description   :   Unloads the Daily Production Report form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()
    
    fbCancelled = True
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates the customer and Prints the Report.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()

Dim bIsValidCustomer    As Boolean
Dim bNoError            As Boolean

    On Error GoTo ErrHandler
    
    fbCancelled = False
    txtCustomerNo_LostFocus
    
    ' Validating the customer information entered.
    
    If Trim(txtCustomerNo.Text) <> "0" Then
    
        bNoError = ValidateCustomer(Trim(txtCustomerNo.Text), _
                                        Left(cboCustomerType.Text, 1), _
                                        bIsValidCustomer)
        If bNoError Then
        
            fbIsValid = bIsValidCustomer
            
            If Not bIsValidCustomer Then
                
                'Displays the message when customer is invalide
                
                ShowErrorMsg "DailyPro004", Me.Caption, vbOKOnly
                txtCustomerNo.SetFocus
                Exit Sub
            
            End If
            
            Unload Me
        
        Else
            
            'setting the focus to the customer textbox
            'if the customer is invalid
            
            txtCustomerNo.SetFocus
            fbCancelled = True
        
        End If
        
    Else
        
        fbIsValid = True
        Unload Me
        
    End If
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:

    'Error message display in case of VB errors in this function.
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyPro014", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Prints the Report if the customer is valid.
'* Parameter Description    :   Cancel - Integer that determines whether the
'*                              form is removed from the screen
'*                              UnLoadMode - Integer
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim sDivisionCode       As String
Dim lPrintRpt           As Long
Dim lResult             As Long
    
    On Error GoTo ErrHandler
    
    If fbCancelled Then
    
        Unload Me
    
    Else
        If Not fbIsValid Then
            Cancel = 1
        Else
            If cboDivisionCode.ListIndex = 1 Then
                sDivisionCode = Left(cboDivisionCode.Text, 1)
            Else
                sDivisionCode = Left(cboDivisionCode.Text, 2)
            End If
            
            'when all the details has been selected
            
            If cboDivisionCode.ListIndex = 1 Or _
                (cboDivisionCode.ListIndex = 0 And _
                 Trim$(txtCustomerNo.Text) = "0") Then
                
                'A confirmation message to print all details is shown.
                
                lPrintRpt = ShowErrorMsg("DailyPro005", Me.Caption, vbYesNo)
        
                If lPrintRpt = vbYes Then
                    
                    'calls print function for printing the report
                    
                    lResult = PrintDailyProductionRpt(sDivisionCode, _
                                            Trim$(txtCustomerNo.Text), _
                                            Left(cboCustomerType.Text, 1), _
                                            cboReportShift.Text, _
                                            cboProdBreakout.Text)
                                            
                    'If there is any problem in printing the report
                    
                    If lResult <> 0 Then GoTo ErrHandler
                    
                ElseIf lPrintRpt = vbNo Then
                
                    Unload Me
                    
                End If
                
            'when specific division code or customer is selected
            
            Else
                
                'calls print function for printing the report
                
                lResult = PrintDailyProductionRpt(sDivisionCode, _
                                        Trim$(txtCustomerNo.Text), _
                                        Left(cboCustomerType.Text, 1), _
                                        cboReportShift.Text, _
                                        cboProdBreakout.Text)
                                                        
                'If there is any problem in printing the report
                
                If lResult <> 0 Then GoTo ErrHandler
                
            End If
            
        End If
        
        fbCancelled = True
        
    End If
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:
    
    fbCancelled = True
    Cancel = 1
    cboDivisionCode.SetFocus
    
    If lResult <> 0 Then

        GoTo CleanUpAndExit

    End If
    
    'Error message display in case of VB errors in this function.
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyPro015", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Validate the Customer information entered.
'* Parameter Description    :   sCustomerNo   - Customer No entered.
'*                              sCustomerType - Customer Type selected.
'*                              bIsValid      - Indicates whether the Customer
'*                              is valid for the report or not.
'* Return Type Description  :   Parameter to check the error.
'******************************************************************************

Private Function ValidateCustomer(ByVal sCustomerNo As String, _
                                  ByVal sCustomerType As String, _
                                  ByRef bIsValid As Boolean) As Boolean

Dim objDailyProductionRpt       As Object
Dim lResult                     As Long

    On Error GoTo ErrHandler
    
    Set objDailyProductionRpt = _
            CreateObject("InventoryFunctions.DailyProductionRpt")
                
    'validating the customer using the function ValidateCustomer
    'of the object DailProductionRpt
    
    lResult = objDailyProductionRpt.ValidateCustomer(sCustomerNo, _
                                                     sCustomerType, _
                                                     bIsValid)
    
    'If there is any error in validating the customer
    
    If lResult <> 0 Then GoTo ErrHandler
    ValidateCustomer = True
    
CleanUpAndExit:

    Set objDailyProductionRpt = Nothing
    Exit Function
    
ErrHandler:
    
    If lResult <> 0 Then
        
        'Error message display in case of Server side errors
        'Error in validating Customer
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "DailyPro006", Me.Caption, vbOKOnly, gcolErrMsg
    Else
        
        'Error message display in case of VB errors in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "DailyPro007", Me.Caption, vbOKOnly, gcolErrMsg
    End If
    
    ValidateCustomer = False
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Prints Daily Production Report.
'* Parameter Description    :   sDivisionCode - Division Code selected.
'*                              sCustomerNo   - Customer No selected.
'*                              sCustomerType - Customer Type selected.
'*                              sReportShift  - Shift selected.
'*                              sProdBreakout - Parameter to check whether a
'*                              a detailed report is needed or not.
'* Return Type Description  :   Error as Long.
'******************************************************************************

Private Function PrintDailyProductionRpt(ByVal sDivisionCode As String, _
                                         ByVal sCustomerNo As String, _
                                         ByVal sCustomerType As String, _
                                         ByVal sReportShift As String, _
                                         ByVal sProdBreakout As String) As Long

Dim objDailyProductionRpt       As Object
Dim rsDailyProductionRpt        As ADODB.Recordset
Dim colParameters               As Collection
Dim lResult                     As Long
Dim sRptFilename                As String
Dim lPrintResult                As Long

    On Error GoTo ErrHandler
    
    PrintDailyProductionRpt = 0
    
    Set objDailyProductionRpt = _
            CreateObject("InventoryFunctions.DailyProductionRpt")
            
    'Change of status bar message in case of Loading from the database.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    'Getting daily production report details using the method
    'GetDailyProductionRptData of the object DailyProductionRpt
    
    lResult = objDailyProductionRpt.GetDailyProductionRptData(gsPlantCode, _
                                                        sDivisionCode, _
                                                        sCustomerNo, _
                                                        sCustomerType, _
                                                        sReportShift, _
                                                        sProdBreakout, _
                                                        rsDailyProductionRpt)
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault

    If lResult <> 0 Then GoTo ErrHandler
    
    If sProdBreakout = "Y" Then
    
        sRptFilename = "DailyProductionBreakout.rpt"
        
    Else
    
        sRptFilename = "DailyProductionWithoutBreakout.rpt"
        
    End If
    
    'Replacing the "'" in the plantcode inorder to
    'avoid the error in Crystal report
    
    gsPlantName = Replace(gsPlantName, "'", "@")
    
    If Not rsDailyProductionRpt.EOF Then
        
        'Passing the parameters to report
        
        Set colParameters = New Collection
        colParameters.Add "PlantName"
        colParameters.Add gsPlantName
        colParameters.Add "ReportShift"
        colParameters.Add cboReportShift.Text
        colParameters.Add "PetDate"
        colParameters.Add fsPetDate
        
        Me.MousePointer = vbHourglass
            
        'Call the function in the modPrint for printing the report
        
        lPrintResult = PrintReportWithParameters("Daily Production Report", _
                                                 sRptFilename, _
                                                 rsDailyProductionRpt, _
                                                 False, _
                                                 colParameters)
        
        If lPrintResult <> 0 Then GoTo ErrHandler
        
    Else
        
        'When there is no data for the report
        
        ShowErrorMsg "DailyPro008", Me.Caption, vbOKOnly
    
    End If
    
CleanUpAndExit:
    
    'resets the status text
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    Me.MousePointer = vbDefault
    
    'cleaning up of objects,recordswet and collection
    
    Set objDailyProductionRpt = Nothing
    Set rsDailyProductionRpt = Nothing
    Set colParameters = Nothing
    Exit Function
    
ErrHandler:
     
    If lPrintResult <> 0 Then
    
        PrintDailyProductionRpt = lPrintResult
        GoTo CleanUpAndExit
    
    End If
     
    If lResult <> 0 Then
        
        'Error message display in case of Server side errors
        'Error in retreiving daily production report details
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "DailyPro009", Me.Caption, vbOKOnly, gcolErrMsg
        
    Else
        lResult = Err.Number
        'Error message display in case of VB errors in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "DailyPro010", Me.Caption, vbOKOnly, gcolErrMsg
        
        
    End If
    
    PrintDailyProductionRpt = lResult
    
    GoTo CleanUpAndExit
    
End Function

