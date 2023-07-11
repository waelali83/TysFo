VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProductRecallReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Recall Report"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProductRecallReport 
      Height          =   3795
      Left            =   65
      TabIndex        =   12
      Top             =   0
      Width           =   5565
      Begin MSComCtl2.DTPicker dtFromShipped 
         Height          =   300
         Left            =   1920
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "M/d/yyyy"
         Format          =   59965443
         CurrentDate     =   38335
      End
      Begin VB.TextBox txtToSlNo 
         Height          =   285
         Left            =   3480
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2355
         Width           =   1215
      End
      Begin VB.TextBox txtFromSlNo 
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2355
         Width           =   1215
      End
      Begin VB.ComboBox cboOrgPlantCode 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1923
         Width           =   855
      End
      Begin VB.TextBox txtCustNo 
         Height          =   285
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1521
         Width           =   1455
      End
      Begin VB.ComboBox cboProdShift 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1089
         Width           =   735
      End
      Begin VB.TextBox txtPN 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Top             =   270
         Width           =   1770
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   4365
         TabIndex        =   11
         Top             =   3330
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   3285
         TabIndex        =   10
         Top             =   3330
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker dtpFromMFGDate 
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59965441
         CurrentDate     =   37469
      End
      Begin MSComCtl2.DTPicker dtpToMFGDate 
         Height          =   300
         Left            =   3600
         TabIndex        =   2
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59965441
         CurrentDate     =   37469
      End
      Begin MSComCtl2.DTPicker dtToShipped 
         Height          =   300
         Left            =   3840
         TabIndex        =   9
         Top             =   2760
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "M/d/yyyy"
         Format          =   59965443
         CurrentDate     =   38335
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Index           =   9
         Left            =   3600
         TabIndex        =   22
         Top             =   2760
         Width           =   150
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Shipped. From"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1650
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   2355
         Width           =   150
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial No. From"
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
         Index           =   6
         Left            =   480
         TabIndex        =   19
         Top             =   2340
         Width           =   1290
      End
      Begin VB.Label lblRS 
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
         Height          =   225
         Index           =   5
         Left            =   315
         TabIndex        =   18
         Top             =   1930
         Width           =   1455
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Id"
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
         Left            =   750
         TabIndex        =   17
         Top             =   1515
         Width           =   1020
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Shift"
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
         Left            =   735
         TabIndex        =   16
         Top             =   1100
         Width           =   1035
      End
      Begin VB.Label lblRS 
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
         Index           =   0
         Left            =   645
         TabIndex        =   15
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Mfg. Date"
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
         Left            =   525
         TabIndex        =   14
         Top             =   685
         Width           =   1245
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   690
         Width           =   150
      End
   End
End
Attribute VB_Name = "frmProductRecallReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmProductRecallReport.frm
'*File Description              : To get user input for printing
'*                                Product Recall Report
'*Author                        : US Technology
'*Date Created                  : Nov-11-02
'*Date Last Modified            : Mar-07-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : LoadOutFunctions.vbp
'*Components Used               : None
'*Functions Defined             : ValidateProduct
'*                                GenerateReports
'*                                ValidateDates
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit


Private Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) _
                As Long

Private Declare Function GetWindowRect Lib "user32" _
                        (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    
End Type

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Calls method to Generates the Reports.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()
'Added By TCS
Dim colParam        As Collection

    Set colParam = New Collection

    colParam.Add gsPlantCode, "Plant_Code"
    
    If Trim(txtPN.Text) = "" Then

        ShowErrorMsg "ProdRecRpt001", Me.Caption, vbOKOnly
        txtPN.SetFocus
        Exit Sub
    Else
        colParam.Add Trim(txtPN.Text), "Product_Code"
    End If
    
    If ValidateProduct = "0" Then Exit Sub
    If Not ValidateDates Then Exit Sub
    
    
' Form date for MFG date

    colParam.Add dtpFromMFGDate.Value, "From_Date"
    colParam.Add dtpToMFGDate.Value, "To_Date"

' Product Shift
    If cboProdShift.ListIndex = -1 Then
        colParam.Add "ALL", "Prod_Shift"
    Else
        colParam.Add Trim(cboProdShift.Text), "Prod_Shift"
    End If
    
'Customer ID
    If Len(txtCustNo.Text) = 0 Then
        colParam.Add "ALL", "Customer_ID"
    Else
        colParam.Add Trim(txtCustNo.Text), "Customer_ID"
    End If
    
'Origin PlantCode
    If cboOrgPlantCode.ListIndex = -1 Then
        colParam.Add "ALL", "Origin_Plant_Code"
    Else
        colParam.Add Trim(cboOrgPlantCode.Text), "Origin_Plant_Code"
    End If
    
' Serial No - Start from & to
    If Len(txtFromSlNo.Text) = 0 Then
        colParam.Add "ALL", "From_Serial"
        colParam.Add "ALL", "To_Serial"
    Else
        colParam.Add Trim(txtFromSlNo.Text), "From_Serial"
        
        If Len(txtToSlNo.Text) = 0 Then
            colParam.Add Trim(txtFromSlNo.Text), "To_Serial"
        Else
            colParam.Add Trim(txtToSlNo.Text), "To_Serial"
        End If
    End If
    
    
' Date Shipped - from & To
    If IsNull(dtFromShipped.Value) Then
        colParam.Add "ALL", "FromShipDate"
        colParam.Add "ALL", "ToShipDate"
    Else
        colParam.Add dtFromShipped.Value, "FromShipDate"
        If IsNull(dtToShipped.Value) Then
            colParam.Add dtFromShipped.Value, "ToShipDate"
        Else
            colParam.Add dtToShipped.Value, "ToShipDate"
        End If
    End If
    
    Dim I
    
    


    
    If GenerateReports(colParam) = True Then
        
        Unload Me
    
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets default values on screen
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()
    
    dtpFromMFGDate.Value = Date - 365
    dtpToMFGDate.Value = Date
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    
'Added by TCS
' For REQ 17 & 20 - Dec-14-04
    cboProdShift.AddItem "ALL"
    cboProdShift.AddItem "A"
    cboProdShift.AddItem "B"
    cboProdShift.ListIndex = 0
' Customer No
    txtCustNo.Text = "ALL"
    
' Populate Origin Plant Code
    cboOrgPlantCode.AddItem "ALL"
    ' Loading Orgin Plants Code from DB
    LoadOrgPlantCodes
    cboOrgPlantCode.ListIndex = 0
    
' Default Value for Serial no
    txtFromSlNo.Text = "ALL"
    txtToSlNo.Text = "ALL"

'Default Value for  Shipped Date
    dtFromShipped.Value = Date
    dtFromShipped.Value = Null
    dtToShipped.Value = Date
    dtToShipped.Value = Null
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates the Product Code
'* Parameter Description    :   None.
'* Return Type Description  :   Returns true - success, false - failure
'******************************************************************************

Private Function ValidateProduct() As String

Dim objLoadOut  As Object
Dim sValid      As String
Dim lResult     As Long

On Error GoTo ErrHandler

    Set objLoadOut = CreateObject("LoadOutFunctions.ProductRecallRpt")
    lResult = objLoadOut.VerifyProduct(gsPlantCode, txtPN.Text, sValid)
    If lResult <> 0 Then GoTo ErrHandler
    
    If sValid = "0" Then
    
        ' Show the invalid product code message
        
        ValidateProduct = sValid
        ShowErrorMsg "ProdRecRpt002", Me.Caption, vbOKOnly
        txtPN.SetFocus
    
    End If
    
CleanUpandExit:
    
    Set objLoadOut = Nothing
    Exit Function

ErrHandler:

    ValidateProduct = "0"
    If lResult <> 0 Then
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "ProdRecRpt006", _
                     Me.Caption, _
                     vbOKOnly, gcolErrMsg
    
    Else
        ' VB error
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "ProdRecRpt008", _
                     Me.Caption, _
                     vbOKOnly, gcolErrMsg
    
    End If
    GoTo CleanUpandExit

End Function

'******************************************************************************
'* Functional Description   :   Generates the reports
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - True - Success, False - Failure
'******************************************************************************

Private Function GenerateReports(ByVal clParam As Collection) As Boolean

Dim objLoadOut                  As Object
Dim lResult                     As Long
Dim lHistResult                 As Long




' this is set to true only if both detail *AND* history report return zero rows

Dim rsDetails                   As ADODB.Recordset
Dim rsHistory                   As ADODB.Recordset

'Holds the name of the crystal database object

Dim CrDatabase                  As CRPEAuto.Database

'Holds the collection of the crystal database tables

Dim crTables                    As CRPEAuto.DatabaseTables

'Holds a table of the crystal databasetables collection

Dim crTable                     As CRPEAuto.DatabaseTable

'Holds a table of the crystal database tables collection

Dim crTableReport               As CRPEAuto.DatabaseTable

'Holds a table of the crystal database tables collection

Dim crTableGraph                As CRPEAuto.DatabaseTable

'Holds a collection of the crystal database formula fields collection

Dim crFields                    As CRPEAuto.FormulaFieldDefinitions

'Holds the file name

Dim sCrTempFileName             As String

'Holds the name for the sub report

Dim sCrTempSubFileName          As String

'Various crystal report objects.

Dim crReport                    As CRPEAuto.Report
Dim crSubRep                    As CRPEAuto.Report
Dim crDat                       As CRPEAuto.Database
Dim crTabs                      As CRPEAuto.DatabaseTables
Dim crTab                       As CRPEAuto.DatabaseTable
Dim mcrOption                   As CRPEAuto.ReportOptions
Dim view                        As CRPEAuto.view

'Holds a collection of the crystal database formula fields collection

Dim crField1                    As CRPEAuto.FormulaFieldDefinition
Dim crField2                    As CRPEAuto.FormulaFieldDefinition
Dim crField3                    As CRPEAuto.FormulaFieldDefinition
Dim crField4                    As CRPEAuto.FormulaFieldDefinition
'Added by TCS
Dim crField5                    As CRPEAuto.FormulaFieldDefinition
Dim crField6                    As CRPEAuto.FormulaFieldDefinition
Dim crField7                    As CRPEAuto.FormulaFieldDefinition
Dim crField8                    As CRPEAuto.FormulaFieldDefinition
Dim crField9                    As CRPEAuto.FormulaFieldDefinition
Dim crField10                   As CRPEAuto.FormulaFieldDefinition
Dim crField11                   As CRPEAuto.FormulaFieldDefinition

'Dimension variables to be used during the sizing of the preview window.

Dim ipixX                       As Integer
Dim ipixY                       As Integer
Dim lWnd                        As Long
Dim TaskbarDim                  As RECT
Dim iTaskbarHt                  As Integer
Dim mcrPrintOption              As CRPEAuto.PrintWindowOptions

Dim CrSections                  As CRPEAuto.Sections
Dim CrSection                   As CRPEAuto.Section
Dim CrReportObjs                As CRPEAuto.ReportObjects
Dim CrSubreportObj              As CRPEAuto.SubreportObject
Dim CrSubreport                 As CRPEAuto.Report
Dim CrDatabaseTables            As CRPEAuto.DatabaseTables
Dim CrDatabaseTable             As CRPEAuto.DatabaseTable

'Holds a table of the crystal databasetables collection

Dim crTable1                    As CRPEAuto.DatabaseTable
Dim iCheckResult                As Integer

On Error GoTo ErrHandler

    GenerateReports = False
    Set objLoadOut = CreateObject("LoadOutFunctions.ProductRecallRpt")

    Set rsDetails = New ADODB.Recordset
    Set rsHistory = New ADODB.Recordset
    
    'changing the status bar text when details are loading from the database
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    iCheckResult = 0
    
    'get the data for detail report
    If clParam.Item("FromShipDate") = "ALL" Then
        
        lResult = objLoadOut.PrintDetailReport(clParam, rsDetails)
    
    End If
    'get the data for history report
    
    lHistResult = objLoadOut.PrintHistoryReport(clParam, rsHistory)
                                            
    
    'If there is any error in retrieving both the data
    
    If lResult <> 0 And lHistResult <> 0 Then GoTo ErrHandler

'Changed by TCS Dec-14-04

    
    If rsDetails Is Nothing And rsHistory Is Nothing Then
        'Error message displaying "No data for report" condition.
        ShowErrorMsg "OutCustBOL009", Me.Caption, vbOKOnly
        GenerateReports = True
        GoTo CleanUpandExit
        
    Else
'        If rsDetails.EOF And rsHistory.EOF Then
'            'Error message displaying "No data for report" condition.
'            ShowErrorMsg "OutCustBOL009", Me.Caption, vbOKOnly
'            GenerateReports = True
'            GoTo CleanUpandExit
'        End If
    End If



        'Setting of message in status bar during the process of
        'printing.

        frmMainMenu.sbMainMenu.Panels(1).Text = _
                            "Printing Report - Please Wait..."
        
        sCrTempFileName = gsReportPath & "\" & "ProductRecall.rpt"

        Set crReport = CrystalApplication.OpenReport(sCrTempFileName)
        
        'Passing data to the main report to print order summary details
        
        Set CrDatabase = crReport.Database
        Set crTables = CrDatabase.Tables
        Set crTable1 = crTables(1)
        
        If Not rsDetails Is Nothing Then
            Call crTable1.SetPrivateData(3, rsDetails)
        End If
        
        Set CrSections = crReport.Sections
        
        'passing data to the duplicate subreport

        Set CrSection = CrSections.Item(1)
        Set CrReportObjs = CrSection.ReportObjects

        Set CrSubreportObj = CrReportObjs.Item(2)
        Set CrSubreport = crReport.OpenSubreport(CrReportObjs(2).Name)

        Set CrDatabase = CrSubreport.Database
        Set CrDatabaseTables = CrDatabase.Tables
        Set CrDatabaseTable = CrDatabaseTables.Item(1)

        If Not rsHistory Is Nothing Then
            CrDatabaseTable.SetPrivateData 3, rsHistory
        End If
        
        'passing data to the sub report
        
        Set CrSection = CrSections.Item(14)
        Set CrReportObjs = CrSection.ReportObjects
    
        Set CrSubreportObj = CrReportObjs.Item(1)
        Set CrSubreport = crReport.OpenSubreport(CrReportObjs(1).Name)
    
        Set CrDatabase = CrSubreport.Database
        Set CrDatabaseTables = CrDatabase.Tables
        Set CrDatabaseTable = CrDatabaseTables.Item(1)
        
        If Not rsHistory Is Nothing Then
            CrDatabaseTable.SetPrivateData 3, rsHistory
        End If
        
        'checking which data is there
        
'Changed by TCS
        If Not rsDetails Is Nothing Then
            If Not rsHistory Is Nothing Then
                If rsDetails.State = adStateOpen And rsHistory.State = adStateOpen Then
                    If Not rsDetails.EOF And rsHistory.EOF Then
                        iCheckResult = 1
                    ElseIf rsDetails.EOF And Not rsHistory.EOF Then
                        iCheckResult = 2
                    ElseIf rsDetails.EOF And rsHistory.EOF Then
                        iCheckResult = 3
                    End If
                Else
                    iCheckResult = 3
                End If
            End If
        End If
        
        Set crFields = crReport.FormulaFields

        Set crField1 = crFields("FromDate")
        crField1.Text = "'" & Format(dtpFromMFGDate.Value, _
                                     "MM/DD/YYYY") & "'"

        Set crField2 = crFields("ToDate")
        crField2.Text = "'" & Format(dtpToMFGDate.Value, _
                                     "MM/DD/YYYY") & "'"

        Set crField3 = crFields("ProductCode")
        crField3.Text = "'" & txtPN.Text & "'"
        
        Set crField4 = crFields("Flag")
        crField4.Text = "'" & iCheckResult & "'"
        
  ' Added by TCS
                
        'Production Shift
         Set crField5 = crFields("ProdShift")
                crField5.Text = "'" & clParam.Item("Prod_Shift") & "'"
        
        
        'Customer Id
        Set crField6 = crFields("Customer")
                crField6.Text = "'" & clParam.Item("Customer_ID") & "'"
        
        
        'Origin Plant Code
        Set crField7 = crFields("OriginPlant")
                crField7.Text = "'" & clParam.Item("Origin_Plant_Code") & "'"
        
        
        'From Serial Number
        Set crField8 = crFields("FromSerial")
                crField8.Text = "'" & clParam.Item("From_Serial") & "'"
        
        
        'To Serial Number
        Set crField9 = crFields("ToSerial")
                crField9.Text = "'" & clParam.Item("To_Serial") & "'"
        
        
        'From Date Shipped
        Set crField10 = crFields("FromDateShipped")
                crField10.Text = "'" & Format(clParam.Item("FromShipDate"), "MM/DD/YYYY") & "'"
        
        
        'To Date Shipped
        Set crField11 = crFields("ToDateShipped")
                crField11.Text = "'" & Format(clParam.Item("ToShipDate"), "MM/DD/YYYY") & "'"
        
                
                
        
        Set mcrOption = crReport.Options
        Set mcrPrintOption = crReport.PrintWindowOptions
        
        mcrOption.ZoomMode = crFullSize
        mcrPrintOption.HasPrintSetupButton = True

        ipixX = Screen.Width / Screen.TwipsPerPixelX
        ipixY = Screen.Height / Screen.TwipsPerPixelY

        'Get the hWnd of the taskbar

        lWnd = FindWindow("Shell_TrayWnd", vbNullString)

        'Fill the structure Rect

        GetWindowRect lWnd, TaskbarDim

        'Height of the taskbar

        iTaskbarHt = TaskbarDim.Bottom - TaskbarDim.Top
        
        'Code for creating a print preview for the report.
        
        Set view = crReport.Preview("Product Recall Report", 0, 0, ipixX, _
                                     ipixY - iTaskbarHt)
        glReportHandle = GetActiveWindow()
        SetWindowPos glReportHandle, _
                     HWND_TOPMOST, _
                     0, 0, 0, 0, _
                     SWP_NOSIZE Or SWP_NOMOVE
    
    GenerateReports = True
    
CleanUpandExit:
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    Set objLoadOut = Nothing
    Set rsDetails = Nothing
    Set rsHistory = Nothing
    Exit Function
    
ErrHandler:

   GenerateReports = False
   If lResult <> 0 Then
        
        gcolErrMsg.Add lResult
        ShowErrorMsg "ProdRecRpt005", _
                     Me.Caption, _
                     vbOKOnly, gcolErrMsg
    
    Else
    
        'VB Error
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "ProdRecRpt007", _
                     Me.Caption, _
                     vbOKOnly, gcolErrMsg
    
    End If
    GoTo CleanUpandExit

End Function

'******************************************************************************
'* Functional Description   :   Checks whether the From date < To Date
'* Parameter Description    :   None.
'* Return Type Description  :   True on success, False on failure
'******************************************************************************

Private Function ValidateDates() As Boolean
    
    ValidateDates = True
    If dtpFromMFGDate.Value > dtpToMFGDate.Value Then
        
        ValidateDates = False
        ShowErrorMsg "ProdRecRpt004", Me.Caption, vbOKOnly
        dtpFromMFGDate.SetFocus
    
    End If

End Function

Private Sub txtCustNo_Change()
    
Static sOrigValue       As String 'Holds the previous value in the text
    
    ReturnValueForCapsAndNumerics txtCustNo, sOrigValue
    txtCustNo = sOrigValue

End Sub

Private Sub txtCustNo_GotFocus()
    txtCustNo.SelStart = 0
    txtCustNo.SelLength = Len(txtCustNo.Text)
End Sub

Private Sub txtCustNo_KeyPress(KeyAscii As Integer)
    ReturnKeyForCapsAndNumerics txtCustNo, KeyAscii
End Sub

Private Sub txtFromSlNo_Change()
    
Static sOrigValue       As String 'Holds the previous value in the text
    
    ReturnValueForCapsAndNumerics txtFromSlNo, sOrigValue
    txtFromSlNo = sOrigValue
End Sub

Private Sub txtFromSlNo_GotFocus()
    txtFromSlNo.SelStart = 0
    txtFromSlNo.SelLength = Len(txtFromSlNo.Text)
End Sub

Private Sub txtFromSlNo_KeyPress(KeyAscii As Integer)
    ReturnKeyForCapsAndNumerics txtFromSlNo, KeyAscii
End Sub

'******************************************************************************
'* Functional Description     : Handles Change Event of the text control
'* Parameter Description      : None
'* Return Type Description    : None
'******************************************************************************

Private Sub txtPN_Change()
    
Static sOrigValue       As String 'Holds the previous value in the text
    
    ReturnValueForCapsAndNumerics txtPN, sOrigValue
    txtPN = sOrigValue

End Sub

'******************************************************************************
'* Functional Description     : Selects the text when the focus shifts to
'*                              text box for product code
'* Parameter Description      : None
'* Return Type Description    : None
'******************************************************************************

Private Sub txtPN_GotFocus()

    txtPN.SelStart = 0
    txtPN.SelLength = Len(txtPN.Text)

End Sub

'******************************************************************************
'* Functional Description   :   Makes UCase and Numeric entry
'* Parameter Description    :   Keyascii - Integer
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtPN_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtPN, KeyAscii
    
End Sub


'******************************************************************************
'* Functional Description   :   Fill the Origin Plant Code combo
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Parameter to check the error.
'******************************************************************************

Private Function LoadOrgPlantCodes() As Boolean

Dim ObjOrgPlantCode        As Object
Dim rsCodes                     As ADODB.Recordset
Dim lResult                     As Long
    
    On Error GoTo ErrHandler
    
    Set ObjOrgPlantCode = _
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
        
        
    lResult = ObjOrgPlantCode.GetOriginPlantCodes(gsPlantCode, rsCodes)
    
    'Resetting the mouse pointers to default.
    
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    If lResult <> 0 Then GoTo ErrHandler
    
    'Loading the Origin Plant Code combo with values retrieved from database.
    
    If Not rsCodes.EOF Then
        rsCodes.MoveFirst
        While Not rsCodes.EOF
            cboOrgPlantCode.AddItem _
                    rsCodes.Fields("ORIGIN_PLANT_CODE").Value
            rsCodes.MoveNext
        Wend
    End If
    LoadOrgPlantCodes = True
    
CleanUpandExit:
    
    'Setting the message in status bar and cleans up all objects.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    Set ObjOrgPlantCode = Nothing
    Set rsCodes = Nothing
    Exit Function
    
ErrHandler:
    
    If lResult <> 0 Then
        
        'Error message shown in case of server side validation failure.
         
        
            gcolErrMsg.Add lResult
            ShowErrorMsg "AgedProInv011", Me.Caption, vbOKOnly, gcolErrMsg
        

    Else
        
        'Error in case of VB error occuring in this function.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "AgedProInv004", Me.Caption, vbOKOnly, gcolErrMsg

    End If
    
    LoadOrgPlantCodes = False
    GoTo CleanUpandExit
    
End Function

Private Sub txtToSlNo_Change()
Static sOrigValue       As String 'Holds the previous value in the text
    
    ReturnValueForCapsAndNumerics txtToSlNo, sOrigValue
    txtToSlNo = sOrigValue
End Sub

Private Sub txtToSlNo_GotFocus()
    txtToSlNo.SelStart = 0
    txtToSlNo.SelLength = Len(txtToSlNo.Text)
End Sub

Private Sub txtToSlNo_KeyPress(KeyAscii As Integer)
    ReturnKeyForCapsAndNumerics txtToSlNo, KeyAscii
End Sub
