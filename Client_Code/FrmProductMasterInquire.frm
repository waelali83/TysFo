VERSION 5.00
Begin VB.Form FrmProductMasterInquire 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Master - Inquire"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProductMasterInquire 
      Height          =   2775
      Left            =   100
      TabIndex        =   6
      Top             =   0
      Width           =   3615
      Begin VB.TextBox TxtProductCode 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "ALL"
         Top             =   240
         Width           =   1650
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2505
         TabIndex        =   5
         Top             =   2280
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   1000
      End
      Begin VB.ComboBox cboDivisionCode 
         Height          =   315
         ItemData        =   "FrmProductMasterInquire.frx":0000
         Left            =   1560
         List            =   "FrmProductMasterInquire.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   1935
      End
      Begin VB.TextBox txtLabelNo 
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "ALL"
         Top             =   1680
         Width           =   1290
      End
      Begin VB.ComboBox cboProductType 
         Height          =   315
         ItemData        =   "FrmProductMasterInquire.frx":0004
         Left            =   1560
         List            =   "FrmProductMasterInquire.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1190
         Width           =   1935
      End
      Begin VB.Label lblLabelNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label No."
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
         Left            =   120
         TabIndex        =   10
         Top             =   1710
         Width           =   795
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
         Left            =   120
         TabIndex        =   9
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblProdType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Type"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label lblProdCode 
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
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmProductMasterInquire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : FrmProductMasterInquire.frm
'*File Description              : To get user input for Product Master Browse/Update
'*Author                        : US Technology
'*Date Created                  : Nov-01-04
'*Date Last Modified            : Nov-02-04
'*Version                       : 1.0
'*Layer                         : Client
'*Project Referenced            : InventoryFunctions
'*                                MasterFileFunctions
'*Components Used               : None
'*Functions Defined             : 1. SetMainMenu
'*                                2. LoadFormProductMasterBrowseUpdate
'*                                3. GetProductMasterDetails
'*Copyright                     : TCS
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                        Nov-02-04  TCS
'******************************************************************************

Option Explicit

Private fbIsValid           As Boolean
Private fbCancelled         As Boolean

'Variable to check success of form load

Private flLoadSuccess       As Long

Public Event PopulateProductMaster(ByVal sProductCode As String, _
                                     ByVal sDivisionCode As String, _
                                     ByVal sProductType As String, _
                                     ByVal sLabelNo As String)

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    fbCancelled = False
    fbIsValid = False
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Validate the entries and populate details,
'                           :   unload the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()


Dim bNoError                As Boolean
Dim bProductNoError         As Boolean
Dim sProductCode            As String
Dim sProdDivisionCode       As String
Dim sProductType            As String
Dim sLabelNo                As String
    
    
    
    'Fetch the values from the inquire form into variables

    sProductCode = Trim(TxtProductCode.Text)
    sProductType = SplitValue(Trim(cboProductType.Text), 0)
    sProdDivisionCode = SplitValue(Trim(cboDivisionCode.Text), 0)
    sLabelNo = Trim(txtLabelNo.Text)
    
     
    If (sProductCode = "ALL" Or sProductCode = "") Then
        sProductCode = ""
    End If
             
    If (sProductType = "A" Or sProductType = "") Then
        sProductType = ""
    End If
    
    If (sProdDivisionCode = "A" Or sProdDivisionCode = "") Then
        sProdDivisionCode = ""
    End If
             
    If (sLabelNo = "ALL" Or sLabelNo = "") Then
        sLabelNo = ""
    End If
    
    
    
'    If sProductCode = "" And sProductTpye = "" And sProdDivisionCode = "" And sLabelNo = "" Then
'        If ShowErrorMsg("", Me.Caption, vbInformation) = vbYes Then
'
'        End If
'    End If
    RaiseEvent PopulateProductMaster(sProductCode, _
                                      sProdDivisionCode, _
                                      sProductType, _
                                      sLabelNo)
    
    fbCancelled = True
    fbIsValid = True
   Unload Me
    Exit Sub
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    fbIsValid = False
    Unload Me
End If
End Sub

'*******************************************************************************
'* Functional Description   :   Loads Detail Inventory Browse/Update Inquire Form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub Form_Load()

    fbCancelled = True
    
    flLoadSuccess = 0
    
    On Error GoTo ErrHandler
    
    'Setting default value for Customer Number, City, and State
        frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.sbMainMenu.Panels(1).Text = _
        "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    
    
    
    
    Exit Sub

ErrHandler:

    flLoadSuccess = -1
    fbCancelled = True
    
        'Reset Main Menu and status bar
    
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    
    Form_Terminate
    
End Sub

'*******************************************************************************
'* Functional Description   :   The entry point to this mdi from modPrint.
'*                              Invokes Form_Load().
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - Success of form loading.
'*******************************************************************************

Public Function LoadFormProductMasterBrowseUpdate() As Boolean
    
    
    On Error GoTo ErrHandler
    
    Load Me
    
    LoadFormProductMasterBrowseUpdate = False
     fbIsValid = False
    'Initailize Division Code , Product Types
    
    GetProductMasterDetails
    
    'Call of Form_Load

    Me.Show vbModal
     LoadFormProductMasterBrowseUpdate = fbIsValid
    Unload Me
    
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    
    Exit Function

ErrHandler:
    
    LoadFormProductMasterBrowseUpdate = False

End Function

'******************************************************************************
'* Functional Description   :   Set the properties of the Main Menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub SetMainMenu()

    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
 
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
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
End Sub

Private Sub txtLabelNo_Validate(Cancel As Boolean)
    If Len(Trim(txtLabelNo.Text)) = 0 Then
        txtLabelNo.Text = "ALL"
    End If
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProductCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics TxtProductCode, sOrigValue
    TxtProductCode = sOrigValue

End Sub

'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProductCode_GotFocus()
   
    'Select the BayLocation when focus is set.
    
    TxtProductCode.SelStart = 0
    TxtProductCode.SelLength = Len(TxtProductCode.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Calls ReturnKeyForCapsAndNumerics Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProductCode_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics TxtProductCode, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Set the default value of Customer No as zero.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub TxtProductCode_LostFocus()

    ' Set the Customer No field to zero if no Division Code is
    ' selected and if no input is given for the field.
    
    If Me.ActiveControl Is cmdCancel Then Exit Sub
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelNo_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtLabelNo, sOrigValue
    txtLabelNo = sOrigValue

End Sub

'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelNo_GotFocus()
   
    'Select the BayLocation when focus is set.
    
    txtLabelNo.SelStart = 0
    txtLabelNo.SelLength = Len(txtLabelNo.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Calls ReturnKeyForCapsAndNumerics Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelNo_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtLabelNo, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Set the default value of Customer No as zero.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelNo_LostFocus()

    ' Set the Customer No field to zero if no Division Code is
    ' selected and if no input is given for the field.
    
    If Me.ActiveControl Is cmdCancel Then Exit Sub
    
End Sub

'******************************************************************************
'* Functional Description   :   Get Product Type & Division Details
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Function GetProductMasterDetails() As Boolean
    Dim objProduct                  As Object
    Dim frsDivCode                  As ADODB.Recordset
'    Dim objDetailInventory          As InventoryFunctions.CodeByCodeRpt
    Dim lResult                     As Long

    On Error GoTo ErrHandler
    
'1. Get Division Code
    Set objProduct = _
            CreateObject("InventoryFunctions.CodebyCodeRpt")
    lResult = objProduct.GetAllDivisionCodes(frsDivCode)
    
    If lResult <> 0 Then GoTo ErrHandler
    
    If Not frsDivCode Is Nothing Then
        If Not frsDivCode.EOF Then
            
            cboDivisionCode.AddItem "A - Combine All"
            frsDivCode.MoveFirst
            
            While Not frsDivCode.EOF
            
                cboDivisionCode.AddItem _
                        frsDivCode.Fields("TYPE_CODE").Value & " - " & _
                        frsDivCode.Fields("TYPE_SHORT_DESC").Value
                frsDivCode.MoveNext
            
            Wend
            
            cboDivisionCode.ListIndex = 0
            
        End If
     End If
     
    Set objProduct = Nothing
    Set frsDivCode = Nothing

'2. Get Product Types
    Set objProduct = Nothing
    Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
                                                    
    lResult = objProduct.GetProductTypes(frsDivCode)
    
    If lResult <> 0 Then GoTo ErrHandler
    If Not frsDivCode Is Nothing Then
        If Not frsDivCode.EOF Then
            cboProductType.AddItem "A - Combine All"
            frsDivCode.MoveFirst
            While Not frsDivCode.EOF
            
                cboProductType.AddItem _
                        frsDivCode.Fields("TYPE_CODE").Value & " - " & _
                        frsDivCode.Fields("TYPE_SHORT_DESC").Value
                frsDivCode.MoveNext
            
            Wend
            cboProductType.ListIndex = 0
        End If
    End If

    '
    GetProductMasterDetails = True

CleanUpAndExit:

    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT


    Set objProduct = Nothing
    Set frsDivCode = Nothing
    Exit Function


ErrHandler:

    If lResult <> 0 Then
         
         'Display Server error
         
         gcolErrMsg.Add lResult
         giErrMsg = ShowErrorMsg("ProMstr002", _
                                 Me.Caption, _
                                 vbOKOnly, _
                                 gcolErrMsg)
    Else
         
         'Display VB error
         
         lResult = Err.Number
         gcolErrMsg.Add Err.Description
         giErrMsg = ShowErrorMsg("ProMstr012", _
                                 Me.Caption, _
                                 vbOKOnly, _
                                 gcolErrMsg)
    End If

GoTo CleanUpAndExit


End Function
'Used to get Code for the ProductType , Divisioncode
Private Function SplitValue(StVal, IndexNo As Integer)
Dim var
var = Split(StVal, "-")
SplitValue = Trim(var(IndexNo))
End Function

Private Sub TxtProductCode_Validate(Cancel As Boolean)
    If Len(Trim(TxtProductCode.Text)) = 0 Then
        TxtProductCode.Text = "ALL"
    End If
End Sub
