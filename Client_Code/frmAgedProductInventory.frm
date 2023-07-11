VERSION 5.00
Begin VB.Form frmAgedProductInventory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aged Product Inventory Report"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAgedProductInventory 
      Height          =   2595
      Left            =   80
      TabIndex        =   7
      Top             =   0
      Width           =   3675
      Begin VB.ComboBox cboSummary 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboDivisionCode 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboCustomerType 
         Height          =   315
         ItemData        =   "frmAgedProductInventory.frx":0000
         Left            =   1680
         List            =   "frmAgedProductInventory.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   " &OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1500
         TabIndex        =   5
         Top             =   2145
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2565
         TabIndex        =   6
         Top             =   2145
         Width           =   1000
      End
      Begin VB.TextBox txtCustomerNo 
         Height          =   315
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox CboOriginPlantCode 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAgedProductInventory.frx":0004
         Left            =   1680
         List            =   "frmAgedProductInventory.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   1320
         Width           =   795
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
         Height          =   315
         Left            =   135
         TabIndex        =   11
         Top             =   600
         Width           =   1155
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
         Height          =   315
         Left            =   135
         TabIndex        =   10
         Top             =   240
         Width           =   1155
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
         Height          =   315
         Left            =   135
         TabIndex        =   9
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label LblOriginPlantcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origin Plant Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAgedProductInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmAgedProductInventory.frm
'*File Description              : Get the input for Aged Product
'*                                Inventory Report
'*Author                        : US Technology
'*Date Created                  : Aug-26-02
'*Date Last Modified            : Mar-17-02
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : Inventory Functions
'*Components Used               : None
'*Functions Defined             : 1) LoadDivisionCodes
'*                                2) ValidateCustomer
'*                                3) LoadFormAgedProduct
'*                                4) SetMainMenu
'*                                5) PrintAgedProductInvRpt
'*Copyright                     : US Technology
'-------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

'Variable to check for validity of Customer.

Private fbIsValid             As Boolean

'Variable to check for unloading the form.

Private fbCancelled           As Boolean

'Variable to check for Print Error.

Private fbNoPrintError        As Boolean

'Variable to check success of form load

Private flLoadSuccess           As Long

Private Sub CboOriginPlantCode_Click()
    'Enable Customer No only if Division Code is null.
    
'    If cboDivisionCode.ListIndex <> 0 Or CboOriginPlantCode.ListIndex <> 0 Then
'
'        txtCustomerNo.Text = "0"
'        txtCustomerNo.Enabled = False
'        If cboCustomerType.ListCount > 0 Then
'            cboCustomerType.ListIndex = -1
'        End If
'        cboCustomerType.Enabled = False
'
'    Else
'
'        txtCustomerNo.Enabled = True
'
'    End If

End Sub

Private Sub cboSummary_Click()

    CboOriginPlantCode.Enabled = IIf(Trim(cboSummary.Text) = "Y", False, True)
    CboOriginPlantCode.ListIndex = 0

End Sub

'******************************************************************************
'* Functional Description   :   Loads the Aged Product Inventory Report Form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()
    
    fbCancelled = True
    fbNoPrintError = True
    
    ' Load all the Division Codes from the Types table.
    
    flLoadSuccess = 0
    ' Load the Customer Type Combo
        cboCustomerType.AddItem "S - Ship-to"
        cboCustomerType.AddItem "B - Bill-to"
        cboCustomerType.AddItem "C - Corporate"
'Added by TCS
   'Load Summary
        cboSummary.AddItem "Y"
        cboSummary.AddItem "N"
        
    If LoadDivisionCodes Then
    

        cboCustomerType.ListIndex = -1
        
    Else
    
        GoTo ErrHandler
        
    End If
    cboSummary.ListIndex = 0
    
    Exit Sub
    
ErrHandler:

    flLoadSuccess = -1
    fbCancelled = True
    Form_Terminate
                   
End Sub

'******************************************************************************
'* Functional Description   :   The entry point to this mdi from modMain.
'*                              Invokes Form_Load().
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Success of loading the form.
'******************************************************************************

Public Function LoadFormAgedProduct() As Boolean
    
    On Error GoTo ErrHandler
    
    LoadFormAgedProduct = True
    
   'Call of Form_Load

    Me.Show vbModal
    
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    
    Exit Function

ErrHandler:
    
    LoadFormAgedProduct = False

End Function

'******************************************************************************
'* Functional Description   :   Set the properties of the Main Menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub SetMainMenu()

    frmMainMenu.sbMainMenu.Visible = False
    frmMainMenu.mnuViewStatusBar.Checked = False
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
 
End Sub

'******************************************************************************
'* Functional Description   :   Allow the Customer No input based on the
'*                              Division Code selected.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cboDivisionCode_Click()
'Changed by TCS
    'Enable Customer No only if Division Code is null.
    
    If cboDivisionCode.ListIndex <> 0 Then
    
        TxtCustomerNo.Text = "0"
        TxtCustomerNo.Enabled = False
        If cboCustomerType.ListCount > 0 Then
            cboCustomerType.ListIndex = -1
        End If
        cboCustomerType.Enabled = False
        
    Else
    
        TxtCustomerNo.Enabled = True
        
    End If

End Sub

'******************************************************************************
'* Functional Description   :   Unload MDI Window
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
'* Functional Description   :   Calls ValidateForIntegers Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics TxtCustomerNo, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles Customer No change event
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics TxtCustomerNo, sOrigValue
    TxtCustomerNo = sOrigValue
    
    'Enable Customer Type only when Customer No is not null and not zero
    
    If Trim(TxtCustomerNo.Text) <> "" Then
    
        If Trim(TxtCustomerNo.Text) <> "0" Then
        
            cboCustomerType.Enabled = True
            
            If cboCustomerType.ListIndex = -1 Then _
                                        cboCustomerType.ListIndex = 0
            
        Else
        
            cboCustomerType.Enabled = False
            cboCustomerType.ListIndex = -1
            
        End If
        
    Else
    
        cboCustomerType.Enabled = False
        cboCustomerType.ListIndex = -1
        TxtCustomerNo.SetFocus
            
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Set the default value of Customer No as zero.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_LostFocus()

    If Me.ActiveControl Is cmdCancel Then Exit Sub
     
    If Trim(TxtCustomerNo.Text) = "" Then TxtCustomerNo.Text = "0"
      
    If Me.ActiveControl Is cboDivisionCode Then Exit Sub
    
    ' Set the Customer No field to zero if no Division Code is
    ' selected and if no input is given for the field.
    
    If cboDivisionCode.ListIndex = 0 Then
    
        If Trim(TxtCustomerNo.Text) = "" Then
            
            TxtCustomerNo.Text = "0"
              
        End If
        
    End If
    
    'Enable the Customer Type field if the Customer No
    'entered is not Zero.
    
    If Trim(TxtCustomerNo.Text) <> "0" Then
            
        cboCustomerType.Enabled = True
       
        If cboCustomerType.ListIndex = -1 Then cboCustomerType.ListIndex = 0
                
        If Me.ActiveControl <> cmdOK Then cboCustomerType.SetFocus
       
    Else
   
        cboCustomerType.Enabled = False
        cboCustomerType.ListIndex = -1
     
    End If
 
End Sub

'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCustomerNo_GotFocus()
    
    'Select the Customer No when focus is set.
    
    TxtCustomerNo.SelStart = 0
    TxtCustomerNo.SelLength = Len(TxtCustomerNo.Text)
    
End Sub


'******************************************************************************
'* Functional Description   :   Unloads the Aged Product Inventory form.
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
    
    fbNoPrintError = True
    
    txtCustomerNo_LostFocus
    
    ' Validating the customer information entered.
    
    If Trim(TxtCustomerNo.Text) <> "0" Then
    
        bNoError = ValidateCustomer(Trim(TxtCustomerNo.Text), _
                                    Left(cboCustomerType.Text, 1), _
                                    bIsValidCustomer)
                                    
                            
        If bNoError Then
        
            fbIsValid = bIsValidCustomer
            
            If Not bIsValidCustomer Then
                
               'Displaying message in case of Invalid Customer.
               
               ShowErrorMsg "AgedProInv001", Me.Caption, vbOKOnly
               
               TxtCustomerNo.SetFocus
              
            End If
        
        Else
            
            'Setting the validity of Customer to False in case of
            'error in validating the customer.
            
            fbIsValid = False
            TxtCustomerNo.SetFocus
        
        End If
        
    Else
        
        'Setting validity of Customer to True when all Customers are selected.
        
        fbIsValid = True
        
    End If
    
    fbCancelled = False
    Unload Me

End Sub

'******************************************************************************
'* Functional Description   :   Prints the Report if the customer is valid.
'* Parameter Description    :   Cancel-Integer that determines whether the form
'*                              is removed from the screen.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim sDivisionCode       As String
Dim lPrintRpt           As Long
Dim lResult             As Long
Dim sOriginPlantCode    As String ' Added by TCS
Dim sCustomerType       As String

    lResult = 0
    
    If Not fbCancelled Then
             
        If fbNoPrintError Then
      
            If Not fbIsValid Then
            
                Cancel = 1
            
            Else
            
            'Taking input value for Division Code.
'Changed by TCS
                If cboDivisionCode.ListIndex = 1 Then
            
                    sDivisionCode = "ALL"  'Left(cboDivisionCode.Text, 1)
                
                ElseIf cboDivisionCode.ListIndex = 0 Then
                    
                    sDivisionCode = "ALL" ' NULL Previous Value
                
                Else
            
                    sDivisionCode = Left(cboDivisionCode.Text, 2)
                
                End If
            
'Added by TCS
                'Taking input value for Origin Plant Code.
            
                If CboOriginPlantCode.ListIndex = 0 Then
            
                    sOriginPlantCode = "ALL" 'Left(CboOriginPlantCode.Text, 1)
                
                Else
            
                    sOriginPlantCode = Trim(CboOriginPlantCode.Text)
                
                End If
            
                If cboCustomerType.ListIndex = -1 Then
                    
                    sCustomerType = "A"
                
                Else
                    
                    sCustomerType = Left(cboCustomerType.Text, 1)
                
                End If
                
                'Confirmation to print details for all Customers.
'Changed by TCS
                If (cboDivisionCode.ListIndex = 1 And Trim(TxtCustomerNo.Text) = "0") And _
                    cboSummary.Text = "Y" Then
                
'                If (cboDivisionCode.ListIndex = 1 And Trim(txtCustomerNo.Text) = "0") And _
'                (cboSummary.Text = "N" And CboOriginPlantCode.ListIndex = 0) Then
                
'                If (cboDivisionCode.ListIndex = 1 And CboOriginPlantCode.ListIndex = 0) Or _
'                   (cboDivisionCode.ListIndex = 0 And Trim(txtCustomerNo.Text) = "0") Then
                
                        lPrintRpt = ShowErrorMsg("AgedProInv002", _
                                             Me.Caption, _
                                             vbYesNo)
                       
        
                    If lPrintRpt = vbYes Then
                
                        'Prints the report for all Customers.
                    
                        lResult = PrintAgedProductInvRpt(sDivisionCode, _
                                                Trim(TxtCustomerNo.Text), _
                                                sCustomerType, _
                                                sOriginPlantCode, Trim(cboSummary.Text))
                                            
                        If lResult <> 0 Then
                    
                            fbNoPrintError = False
                            Cancel = 1
                            cboDivisionCode.SetFocus
                            Exit Sub
                        
                        End If
                    
                    ElseIf lPrintRpt = vbNo Then
                
                        Unload Me
                    
                    End If
                
                Else
                
                    'Prints the report for the selected Customer.
                
                    lResult = PrintAgedProductInvRpt(sDivisionCode, _
                                            Trim(TxtCustomerNo.Text), _
                                            sCustomerType, _
                                            sOriginPlantCode, Trim(cboSummary.Text))
                
                    If lResult <> 0 Then
                    
                        fbNoPrintError = False
                        Cancel = 1
                        cboDivisionCode.SetFocus
                        Exit Sub
                     
                    End If
                                            
                End If
            
           End If
             
        End If
        
        fbCancelled = True
        
    End If
        
End Sub

'******************************************************************************
'* Functional Description   :   Fill the Division Code combo with all the
'*                              division codes from the Types table.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

Private Function LoadDivisionCodes() As Boolean
'Change by TCS

Dim objAgedProductInvRpt        As Object
Dim rsCodes                     As ADODB.Recordset
Dim lResult                     As Long
'Added by TCS
Dim lResult2                    As Long
    On Error GoTo ErrHandler
    
    Set objAgedProductInvRpt = _
            CreateObject("InventoryFunctions.AgedProductInvRpt")
            
    'Setting the message in the status bar.
    
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
                                    
    'Setting the mouse pointers.
    
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
        
    'Calling on the GetAllDivisionCodes method in AgedProductInvRpt component
    'to retrieve the Division Codes.
        
    lResult = objAgedProductInvRpt.GetAllDivisionCodes(rsCodes)
    
    'Resetting the mouse pointers to default.
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Moving the control over to error handler in case of error in
    'retrieving the Division Codes.
    
    If lResult <> 0 Then GoTo ErrHandler
    
    'Loading the Division Code combo with values retrieved from database.
    
    If Not rsCodes.EOF Then
    
        cboDivisionCode.AddItem ""
        cboDivisionCode.AddItem "A  - Combine All"
        rsCodes.MoveFirst
        
        While Not rsCodes.EOF
        
            cboDivisionCode.AddItem _
                    rsCodes.Fields("TYPE_CODE").Value & " - " & _
                    rsCodes.Fields("TYPE_SHORT_DESC").Value
            rsCodes.MoveNext
            
        Wend
        
        cboDivisionCode.ListIndex = 1
        
    End If
    
    Set rsCodes = Nothing
    
    
    'Added by TCS
    
        'Calling on the GetAllDivisionCodes method in AgedProductInvRpt component
    'to retrieve the Division Codes.
        
    lResult2 = objAgedProductInvRpt.GetOriginPlantCodes(gsPlantCode, rsCodes)
    
    'Resetting the mouse pointers to default.
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Moving the control over to error handler in case of error in
    'retrieving the Division Codes.
    
    If lResult2 <> 0 Then GoTo ErrHandler
    
    'Loading the Division Code combo with values retrieved from database.
    
    If Not rsCodes.EOF Then
    
        
        CboOriginPlantCode.AddItem "ALL"
        rsCodes.MoveFirst
        
        While Not rsCodes.EOF
        
            CboOriginPlantCode.AddItem _
                    rsCodes.Fields("ORIGIN_PLANT_CODE").Value
                    
            rsCodes.MoveNext
            
        Wend
        
        CboOriginPlantCode.ListIndex = 0
        
    End If

    
    LoadDivisionCodes = True
    
CleanUpAndExit:
    
    'Setting the message in status bar and cleans up all objects.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    Set objAgedProductInvRpt = Nothing
    Set rsCodes = Nothing
    Exit Function
    
ErrHandler:
    
    If lResult <> 0 Or lResult2 <> 0 Then
        
        'Error message shown in case of server side validation failure.
         
        If lResult <> 0 Then
            gcolErrMsg.Add lResult
            ShowErrorMsg "AgedProInv011", Me.Caption, vbOKOnly, gcolErrMsg
        Else
            gcolErrMsg.Add lResult
            ShowErrorMsg "AgedProInv014", Me.Caption, vbOKOnly, gcolErrMsg
        End If

    Else
        
        'Error in case of VB error occuring in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "AgedProInv004", Me.Caption, vbOKOnly, gcolErrMsg

    End If
    
    LoadDivisionCodes = False
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Validate the Customer information entered.
'* Parameter Description    :   sCustomerNo   - Customer No entered.
'*                              sCustomerType - Customer Type selected.
'*                              bIsValid      - Indicates whether the Customer
'*                                              is valid for the report or not.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

Private Function ValidateCustomer(ByVal sCustomerNo As String, _
                                  ByVal sCustomerType As String, _
                                  ByRef bIsValid As Boolean) As Boolean

Dim objAgedProductInvRpt        As Object
Dim lResult                     As Long

    On Error GoTo ErrHandler
    
    Set objAgedProductInvRpt = _
            CreateObject("InventoryFunctions.AgedProductInvRpt")
            
    'Calling on the ValidateCustomer method in AgedProductInvRpt component
    'to validate the Customer.
    
    lResult = objAgedProductInvRpt.ValidateCustomer(sCustomerNo, _
                                                    sCustomerType, bIsValid)
                                                    
    'Moving the control over to error handler in case of error in
    'validating the Customer.
    
    If lResult <> 0 Then GoTo ErrHandler
    
    ValidateCustomer = True
    
CleanUpAndExit:
    
    'Cleans up the objects.
    
    Set objAgedProductInvRpt = Nothing
    Exit Function
    
ErrHandler:
    
    If lResult <> 0 Then
        
        'Error message shown in case of server side validation failure.
         
        gcolErrMsg.Add lResult
        ShowErrorMsg "AgedProInv012", Me.Caption, vbOKOnly, gcolErrMsg

    Else
    
        'Error in case of VB error occuring in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "AgedProInv006", Me.Caption, vbOKOnly, gcolErrMsg

    End If
    
    ValidateCustomer = False
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Prints Aged Product Inventory Report
'* Parameter Description    :   sDivisionCode - Division Code selected.
'*                              sCustomerNo   - Customer No entered.
'*                              sCustomerType - Customer Type selected.
'* Return Type Description  :   Long - Error code in case of failure.
'******************************************************************************

Private Function PrintAgedProductInvRpt(ByVal sDivisionCode As String, _
                                        ByVal sCustomerNo As String, _
                                        ByVal sCustomerType As String, _
                                        ByVal sOriginPlantCode As String, _
                                        ByVal sSummary As String) As Long
Dim objAgedProductInvRpt        As Object
Dim rsAgedProductInvRpt         As ADODB.Recordset
Dim lResult                     As Long
Dim lPrintResult                As Long

    On Error GoTo ErrHandler

    PrintAgedProductInvRpt = 0
    
    Set objAgedProductInvRpt = _
            CreateObject("InventoryFunctions.AgedProductInvRpt")
            
    'Setting the staus bar message and mouse pointers.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    'Calling on the GetAgedProductInvRptData method in AgedProductInvRpt
    'component to retrieve the report details.
    
    If sSummary = "Y" Then
        lResult = objAgedProductInvRpt.GetAgedProductInvRptData(gsPlantCode, _
                                                        sDivisionCode, _
                                                        sCustomerNo, _
                                                        sCustomerType, _
                                                        sOriginPlantCode, _
                                                        rsAgedProductInvRpt)
                                                        
    Else
        lResult = objAgedProductInvRpt.GetAgedProductInvRptDetailData(gsPlantCode, _
                                                        sDivisionCode, _
                                                        sCustomerNo, _
                                                        sCustomerType, _
                                                        sOriginPlantCode, _
                                                        rsAgedProductInvRpt)
    
    End If
    'Resetting the mouse pointers to default.
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
        
    'Moving the control over to error handler in case of error in
    'retrieving the report details.
    
    If lResult <> 0 Then GoTo ErrHandler
        
    If Not rsAgedProductInvRpt.EOF Then
        
        Me.MousePointer = vbHourglass
            
        'Prints the report.
        
        If sSummary = "Y" Then
        
            lPrintResult = PrintReport("Aged Product Inventory Report", _
                                   "AgedProductInventory.rpt", _
                                   rsAgedProductInvRpt, _
                                   False)
        Else
            lPrintResult = PrintReport("Aged Product Inventory Report", _
                                   "AgedProductInventoryDetail.rpt", _
                                   rsAgedProductInvRpt, _
                                   False)
        End If
        
        If lPrintResult <> 0 Then GoTo ErrHandler
        
         
    ElseIf rsAgedProductInvRpt.EOF Then
        
        'Error message to be shown in case of no data being present for the
        'report.
    
        ShowErrorMsg "AgedProInv008", Me.Caption, vbOKOnly

    End If
    
CleanUpAndExit:
    
    'Resetting the mouse pointers and status bar message.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    Me.MousePointer = vbDefault
    
    'Cleans up all objects.
    
    Set objAgedProductInvRpt = Nothing
    Set rsAgedProductInvRpt = Nothing
    Exit Function
    
ErrHandler:
    
    If lPrintResult <> 0 Then
    
       'Error message shown in case of server side validation failure.
       
        PrintAgedProductInvRpt = lPrintResult
        GoTo CleanUpAndExit
        
    End If
    
    If lResult <> 0 Then
        
       'Error message shown in case of server side validation failure.
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "AgedProInv013", Me.Caption, vbOKOnly, gcolErrMsg

    Else
        
        'Error message shown in case of VB error in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "AgedProInv010", Me.Caption, vbOKOnly, gcolErrMsg

    End If
    
    PrintAgedProductInvRpt = lResult
    
    GoTo CleanUpAndExit
    
End Function


