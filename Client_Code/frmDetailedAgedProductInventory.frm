VERSION 5.00
Begin VB.Form frmDetailedAgedProductInventory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Global Aged Product Inventory Report"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAgedProductInventory 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtAgedProductDays 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "90"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtProductCode 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Text            =   "ALL"
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cboFreezerLocation 
         Height          =   315
         ItemData        =   "frmDetailedAgedProductInventory.frx":0000
         Left            =   1800
         List            =   "frmDetailedAgedProductInventory.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   3165
         TabIndex        =   3
         Top             =   1920
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   " &OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   2100
         TabIndex        =   2
         Top             =   1920
         Width           =   1000
      End
      Begin VB.ComboBox cboDivisionCode 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblAgedProductDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aged Product Days"
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
         Left            =   135
         TabIndex        =   9
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label lblProductCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Left            =   135
         TabIndex        =   7
         Top             =   960
         Width           =   1125
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
         TabIndex        =   6
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblFreezerLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freezer Location"
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
         Left            =   135
         TabIndex        =   5
         Top             =   600
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmDetailedAgedProductInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmDetailedAgedProductInventory.frm
'*File Description              : Get the input for Detailed Aged Product
'*                                Inventory Report
'*Author                        : US Technology - Ravisankar & Sangeetha TCS for SR 4152
'*Date Created                  : Sep-18-06 & Oct-16-06
'*Project Referenced            : Inventory Functions
'*Components Used               : None
'*Functions Defined             : 1) LoadDivisionCodes
'*                                2) LoadPlantCodes
'*                                3) LoadProductCodes
'*                                4) SetMainMenu
'*                                5) PrintDetailedAgedProductInvRpt
'*Copyright                     : US Technology
'-------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'
'******************************************************************************

Option Explicit

Private fbCancelled           As Boolean    'Variable to check for unloading the form.
Private fbNoPrintError        As Boolean    'Variable to check for Print Error.
Private flLoadSuccess           As Long     'Variable to check success of form load

Private Sub cboFreezerLocation_Click()

    fbCancelled = True
    fbNoPrintError = True
    
    ' Load all the Product Codes from the Types table.
    flLoadSuccess = 0

'   If Not (LoadProductCodes(Left(cboFreezerLocation.Text, 3))) Then
'        GoTo ErrHandler
'   End If
   Exit Sub

ErrHandler:
    flLoadSuccess = -1
    fbCancelled = True
End Sub


Private Sub Form_Load()
    
    fbCancelled = True
    fbNoPrintError = True
    
    ' Load all the Division Codes from the Types table.
    flLoadSuccess = 0
   
   If Not (LoadDivisionCodes) Then
        GoTo ErrHandler
   ElseIf Not (LoadPlantCodes) Then
        GoTo ErrHandler
   'ElseIf Not (LoadProductCodes(Left(cboFreezerLocation.Text, 3))) Then
   '     GoTo ErrHandler
   End If
   
   Exit Sub
ErrHandler:
    flLoadSuccess = -1
    fbCancelled = True
    Form_Terminate
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
'* Functional Description   :   The entry point to this mdi from modMain.
'*                              Invokes Form_Load().
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Success of loading the form.
'******************************************************************************

Public Function LoadFormDetailedAgedProduct() As Boolean
    On Error GoTo ErrHandler
    LoadFormDetailedAgedProduct = True
   'Call of Form_Load
    Me.Show vbModal
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    Exit Function
ErrHandler:
    LoadFormDetailedAgedProduct = False
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
'* Functional Description   :   Fill the Division Code combo with all the
'*                              division codes from the Types table.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

Private Function LoadDivisionCodes() As Boolean
    Dim objDetailedAgedProductInvRpt    As Object
    Dim rsCodes                         As ADODB.Recordset
    Dim lResult                         As Long
    Dim lResult2                        As Long
    On Error GoTo ErrHandler
    
    Set objDetailedAgedProductInvRpt = _
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
    lResult = objDetailedAgedProductInvRpt.GetAllDivisionCodes(rsCodes)
    
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
        
    'Calling on the GetAllDivisionCodes method in AgedProductInvRpt component
    'to retrieve the Division Codes.
    lResult2 = objDetailedAgedProductInvRpt.GetOriginPlantCodes(gsPlantCode, rsCodes)
    
    'Resetting the mouse pointers to default.
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Moving the control over to error handler in case of error in
    'retrieving the Division Codes.
    If lResult2 <> 0 Then GoTo ErrHandler
        
    LoadDivisionCodes = True

CleanUpAndExit:
    
    'Setting the message in status bar and cleans up all objects.
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    Set objDetailedAgedProductInvRpt = Nothing
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
'* Functional Description   :   Fill the Plant Code combo with all the
'*                              palnt codes from the Types table.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

Private Function LoadPlantCodes() As Boolean
    Dim objPlant    As Object
    Dim rsCodes                         As ADODB.Recordset
    Dim lResult                         As Long
    On Error GoTo ErrHandler
    
    Set objPlant = _
            CreateObject("MasterFileFunctions.PlantMaster")
            
    'Setting the message in the status bar.
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
                                    
    'Setting the mouse pointers.
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
        
    'Calling on the GetAllPlants method in PlantMaster component
    'to retrieve the Plant Codes.
    lResult = objPlant.GetAllPlants(rsCodes)
    
    'Resetting the mouse pointers to default.
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Moving the control over to error handler in case of error in
    'retrieving the Plant Codes.
    If lResult <> 0 Then GoTo ErrHandler
    
    'Loading the Plant Code combo with values retrieved from database.
    If Not rsCodes.EOF Then
        cboFreezerLocation.AddItem ""
        cboFreezerLocation.AddItem "ALL"

        rsCodes.MoveFirst

        While Not rsCodes.EOF
            cboFreezerLocation.AddItem _
                rsCodes.Fields("PLANT_CODE").Value & " - " & _
                rsCodes.Fields("NAME").Value
            rsCodes.MoveNext
        Wend
        
        cboFreezerLocation.ListIndex = 1
    End If
    Set rsCodes = Nothing
        
    LoadPlantCodes = True

CleanUpAndExit:
    
    'Setting the message in status bar and cleans up all objects.
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    Set objPlant = Nothing
    Set rsCodes = Nothing
    Exit Function
    
ErrHandler:
    
    If lResult <> 0 Then
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
 
    LoadPlantCodes = False
    GoTo CleanUpAndExit
    
End Function


'******************************************************************************
'* Functional Description   :   Fill the Plant Code combo with all the
'*                              palnt codes from the Types table.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

'Private Function LoadProductCodes(ByVal sFreezerLocation As String) As Boolean
'    Dim objProduct                      As Object
'    Dim rsCodes                         As ADODB.Recordset
'    Dim lResult                         As Long
'    On Error GoTo ErrHandler
'
'    Set objProduct = _
'            CreateObject("MasterFileFunctions.ProductMaster")
'
'    'Setting the message in the status bar.
'    frmMainMenu.sbMainMenu.Visible = True
'    frmMainMenu.mnuViewStatusBar.Checked = True
'    frmMainMenu.sbMainMenu.Panels(1).Text = _
'                                    "Loading From Database - Please Wait..."
'
'    'Setting the mouse pointers.
'    frmMainMenu.MousePointer = vbHourglass
'    Me.MousePointer = vbHourglass
'    DoEvents
'
'    'Calling on the GetProductDetailsForPlant method in ProductMaster component
'    'to retrieve the Product Codes.
'    lResult = objProduct.GetProductDetailsForPlant(sFreezerLocation, rsCodes)
'
'    'Resetting the mouse pointers to default.
'    frmMainMenu.MousePointer = vbDefault
'    Me.MousePointer = vbDefault
'
'    'Moving the control over to error handler in case of error in
'    'retrieving the Plant Codes.
'    If lResult <> 0 Then GoTo ErrHandler
'    '
'   ' cboProductCode.Clear
'    rsCodes.MoveFirst
'    '
'    'Loading the Product Code combo with values retrieved from database.
'    If Not rsCodes.EOF Then
'
'        txtProductCode.AddItem ""
'        txtProductCode.AddItem "ALL"
'        rsCodes.MoveFirst
'
'        While Not rsCodes.EOF
'            txtProductCode.AddItem _
'                rsCodes.Fields("PRODUCT_CODE").Value & " - " & _
'                rsCodes.Fields("PRODUCT_DESC").Value
'            rsCodes.MoveNext
'        Wend
'
'        cboProductCode.ListIndex = 1
'    Else
'
'        cboProductCode.AddItem ""
'        cboProductCode.AddItem "ALL"
'        cboProductCode.ListIndex = 1
'
'    End If
'    Set rsCodes = Nothing
'
'    LoadProductCodes = True
'
'CleanUpAndExit:
'
'    'Setting the message in status bar and cleans up all objects.
'    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
'
'    Set objProduct = Nothing
'    Set rsCodes = Nothing
'    Exit Function
'
'ErrHandler:
'
'    If lResult <> 0 Then
'        'Error message shown in case of server side validation failure.
'        If lResult <> 0 Then
'            gcolErrMsg.Add lResult
'            ShowErrorMsg "AgedProInv011", Me.Caption, vbOKOnly, gcolErrMsg
'        Else
'            gcolErrMsg.Add lResult
'            ShowErrorMsg "AgedProInv014", Me.Caption, vbOKOnly, gcolErrMsg
'        End If
'    Else
'        'Error in case of VB error occuring in this function.
'        gcolErrMsg.Add Err.Description
'        ShowErrorMsg "AgedProInv004", Me.Caption, vbOKOnly, gcolErrMsg
'    End If
'
'    LoadProductCodes = False
'    GoTo CleanUpAndExit
'
'End Function



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
'* Functional Description   :   Prints the Report.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()
    Dim bNoError            As Boolean
    fbNoPrintError = True
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
    Dim sFreezerLocation    As String
    Dim sProductCode        As String
    Dim sAgedProductDays    As String
    
    

    lResult = 0
    
    If Not fbCancelled Then
        If fbNoPrintError Then
            'Taking input value for Division Code.
            If cboDivisionCode.ListIndex = 1 Then
                sDivisionCode = "ALL"  'Left(cboDivisionCode.Text, 1)
            ElseIf cboDivisionCode.ListIndex = 0 Then
                sDivisionCode = "ALL" ' NULL Previous Value
            Else
                sDivisionCode = Left(cboDivisionCode.Text, 2)
            End If
            '
            'Taking input value for Freezer Location.
            If cboFreezerLocation.ListIndex = 1 Then
                sFreezerLocation = "ALL"
            ElseIf cboFreezerLocation.ListIndex = 0 Then
                sFreezerLocation = "ALL" ' NULL Previous Value
            Else
                sFreezerLocation = Left(cboFreezerLocation.Text, 3)
            End If
            '
            'Taking input value for Product Code.
            If txtProductCode.Text = "ALL" Then
                sProductCode = "ALL"
            Else
                sProductCode = txtProductCode.Text
            End If
            'If cboProductCode.ListIndex = 1 Then
            '    sProductCode = "ALL"
            'ElseIf cboProductCode.ListIndex = 0 Then
            '    sProductCode = "ALL"  ' NULL Previous Value
            'Else
            '    sProductCode = Mid(cboProductCode.Text, 1, InStr(cboProductCode.Text, "-") - 2)
            'End If
            '
            'Taking input value for Aged Product Days.
            sAgedProductDays = txtAgedProductDays.Text
            '
            'Confirmation to print details for all Customers.
            If cboDivisionCode.ListIndex = 1 And cboFreezerLocation.ListIndex = 1 And txtProductCode.Text = "ALL" Then
            
                lPrintRpt = ShowErrorMsg("AgedProInv002", _
                    Me.Caption, _
                    vbYesNo)
                    
                If lPrintRpt = vbYes Then
                    'Prints the report for all Customers. WR13228
                    
                    'If Report is not the AllStatus Report
                    If gsTypeReport <> "AllStatus" Then
                    
                       lResult = PrintDetailedAgedProductInvRpt(sDivisionCode, _
                                 sFreezerLocation, _
                                 sProductCode, _
                                 sAgedProductDays)
                    
                    'If the Report is the AllStatus Report
                    ElseIf gsTypeReport = "AllStatus" Then
                              
                              
                        lResult = PrintDetailedAgedProductInvRptAllStatus(sDivisionCode, _
                                  sFreezerLocation, _
                                  sProductCode, _
                                  sAgedProductDays)
                              
                              
                    End If
                                            
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
        
                'If Report is not the AllStatus Report WR13228
                If gsTypeReport <> "AllStatus" Then
             
                  lResult = PrintDetailedAgedProductInvRpt(sDivisionCode, _
                          sFreezerLocation, _
                          sProductCode, _
                          sAgedProductDays)
                          
                'If Report is the AllStatus Report
                ElseIf gsTypeReport = "AllStatus" Then
                              
                              
                  lResult = PrintDetailedAgedProductInvRptAllStatus(sDivisionCode, _
                              sFreezerLocation, _
                              sProductCode, _
                              sAgedProductDays)
                              
                End If
                          
                If lResult <> 0 Then
                
                    fbNoPrintError = False
                    Cancel = 1
                    cboDivisionCode.SetFocus
                    Exit Sub
                 
                End If
               
            End If
        End If
        fbCancelled = True
    End If
End Sub

'******************************************************************************
'* Functional Description   :   Prints Detail Aged Product Inventory Report
'* Parameter Description    :   sDivisionCode    - Division Code selected.
'*                              sFreezerLocation - Freezer Location selected.
'*                              sProductCode     - Product Code selected.
'* Return Type Description  :   Long - Error code in case of failure.
'******************************************************************************

Private Function PrintDetailedAgedProductInvRpt(ByVal sDivisionCode As String, _
                                                ByVal sFreezerLocation As String, _
                                                ByVal sProductCode As String, _
                                                ByVal sAgedProductDays As String) As Long
                                                
Dim objDetAgedProductInvRpt        As Object
Dim rsDetAgedProductInvRpt         As ADODB.Recordset
Dim lResult                     As Long
Dim lPrintResult                As Long

    On Error GoTo ErrHandler

    PrintDetailedAgedProductInvRpt = 0
    
    Set objDetAgedProductInvRpt = _
            CreateObject("InventoryFunctions.AgedProductInvRpt")
            
    'Setting the staus bar message and mouse pointers.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    'Calling on the GetDetailAgedProductInvRptData method in AgedProductInvRpt
    'component to retrieve the report details.
    
     lResult = objDetAgedProductInvRpt.GetDetailAgedProductInvRptData2(sDivisionCode, _
                                                        sFreezerLocation, _
                                                        sProductCode, _
                                                        sAgedProductDays, _
                                                        rsDetAgedProductInvRpt)
                                                        
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
        
    'Moving the control over to error handler in case of error in
    'retrieving the report details.
    
    If lResult <> 0 Then GoTo ErrHandler
        
    If Not rsDetAgedProductInvRpt.EOF Then
        
        Me.MousePointer = vbHourglass
            
        'Prints the report.
        
        lPrintResult = PrintReport("Global Aged Product Inventory Report", _
                                   "DetailedAgedProductInventory.rpt", _
                                   rsDetAgedProductInvRpt, _
                                   False)
        
        If lPrintResult <> 0 Then GoTo ErrHandler
        
         
    ElseIf rsDetAgedProductInvRpt.EOF Then
        
        'Error message to be shown in case of no data being present for the
        'report.
    
        ShowErrorMsg "AgedProInv008", Me.Caption, vbOKOnly

    End If
    
CleanUpAndExit:
    
    'Resetting the mouse pointers and status bar message.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    Me.MousePointer = vbDefault
    
    'Cleans up all objects.
    
    Set objDetAgedProductInvRpt = Nothing
    Set rsDetAgedProductInvRpt = Nothing
    Exit Function
    
ErrHandler:
    
    If lPrintResult <> 0 Then
    
       'Error message shown in case of server side validation failure.
       
        PrintDetailedAgedProductInvRpt = lPrintResult
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
    
    PrintDetailedAgedProductInvRpt = lResult
    
    GoTo CleanUpAndExit
    
End Function
'*********************************************************************************************
'* Functional Description   :   Prints Detail Aged Product Inventory Report All Status WR13228
'* Parameter Description    :   sDivisionCode    - Division Code selected.
'*                              sFreezerLocation - Freezer Location selected.
'*                              sProductCode     - Product Code selected.
'* Return Type Description  :   Long - Error code in case of failure.
'**********************************************************************************************
Private Function PrintDetailedAgedProductInvRptAllStatus(ByVal sDivisionCode As String, _
                                                ByVal sFreezerLocation As String, _
                                                ByVal sProductCode As String, _
                                                ByVal sAgedProductDays As String) As Long
                                                
Dim objDetAgedProductInvRpt        As Object
Dim rsDetAgedProductInvRpt         As ADODB.Recordset
Dim lResult                     As Long
Dim lPrintResult                As Long
Dim intThis                     As Integer



    On Error GoTo ErrHandler

    PrintDetailedAgedProductInvRptAllStatus = 0
    
    Set objDetAgedProductInvRpt = _
            CreateObject("InventoryFunctions.AgedProductInvRpt")
            
    'Setting the staus bar message and mouse pointers.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    'Calling on the GetDetailAgedProductInvRptData method in AgedProductInvRpt
    'component to retrieve the report details.
    
     lResult = objDetAgedProductInvRpt.GetDetailAgedProductInvRptAllStatus(sDivisionCode, _
                                                        sFreezerLocation, _
                                                        sProductCode, _
                                                        sAgedProductDays, _
                                                        rsDetAgedProductInvRpt)
                                                        
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
        
    'Moving the control over to error handler in case of error in
    'retrieving the report details.
    
    If lResult <> 0 Then GoTo ErrHandler
        
    If Not rsDetAgedProductInvRpt.EOF Then
        
        Me.MousePointer = vbHourglass
            
    
       
      ' Prints the report.
        
       lPrintResult = PrintReport("Global Aged Product Inventory Report All Status", _
                                  "DetailedAgedProductInventoryStatus.rpt", _
                                  rsDetAgedProductInvRpt, _
                                  False)
        
       ' If lPrintResult <> 0 Then GoTo ErrHandler
        
         
    ElseIf rsDetAgedProductInvRpt.EOF Then
        
        'Error message to be shown in case of no data being present for the
        'report.
    
        ShowErrorMsg "AgedProInv008", Me.Caption, vbOKOnly

    End If
    
CleanUpAndExit:
    
    'Resetting the mouse pointers and status bar message.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    Me.MousePointer = vbDefault
    
    'Cleans up all objects.
    
    Set objDetAgedProductInvRpt = Nothing
    Set rsDetAgedProductInvRpt = Nothing
    Exit Function
    
    
ErrHandler:
    
    If lPrintResult <> 0 Then
    
       'Error message shown in case of server side validation failure.
       
        PrintDetailedAgedProductInvRptAllStatus = lPrintResult
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
    
    PrintDetailedAgedProductInvRptAllStatus = lResult
    
    GoTo CleanUpAndExit
    
End Function






Private Sub txtAgedProductDays_Change()
    Static sOrigValue   As String
    
    ReturnValueForNumerics txtAgedProductDays, sOrigValue
    txtAgedProductDays = sOrigValue
End Sub

Private Sub txtAgedProductDays_GotFocus()
    txtAgedProductDays.SelStart = 0
    txtAgedProductDays.SelLength = Len(txtProductCode.Text)
End Sub

Private Sub txtAgedProductDays_KeyPress(KeyAscii As Integer)
    ReturnKeyForNumerics txtAgedProductDays, KeyAscii
End Sub

Private Sub txtAgedProductDays_Validate(KeepFocus As Boolean)
    ' If the value is a number larger than 999, keep the focus.
    If Not IsNumeric(txtAgedProductDays.Text) Or Val(txtAgedProductDays.Text) > 999 Or Val(txtAgedProductDays.Text) = 0 Then
        KeepFocus = True
        MsgBox _
        "Please insert a positive number greater than 0 and less than 999.", , "Aged Product Days"
    End If
End Sub

Private Sub txtProductCode_Change()
    Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtProductCode, sOrigValue
    txtProductCode = sOrigValue
    
End Sub

Private Sub txtProductCode_GotFocus()
    txtProductCode.SelStart = 0
    txtProductCode.SelLength = Len(txtProductCode.Text)
End Sub

Private Sub txtProductCode_KeyPress(KeyAscii As Integer)
    ReturnKeyForCapsAndNumerics txtProductCode, KeyAscii
End Sub

Private Sub txtProductCode_Validate(Cancel As Boolean)
    If Len(Trim(txtProductCode.Text)) = 0 Then
        txtProductCode.Text = "ALL"
    End If
End Sub
