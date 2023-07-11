VERSION 5.00
Begin VB.Form frmAddProduct 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Production - Add / Update"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3675
   Icon            =   "frmAddProduct.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInquire 
      Height          =   1785
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1365
         TabIndex        =   4
         Top             =   1335
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2445
         TabIndex        =   5
         Top             =   1335
         Width           =   1000
      End
      Begin VB.TextBox txtWgt 
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
         Left            =   1550
         MaxLength       =   12
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtProdCode 
         Height          =   285
         Left            =   1550
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtBoxes 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1550
         MaxLength       =   7
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Weight"
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
         Left            =   205
         TabIndex        =   8
         Top             =   1005
         Width           =   1200
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
         Left            =   205
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Boxes"
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
         Left            =   205
         TabIndex        =   6
         Top             =   630
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmAddProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmAddProduct.frm
'*File Description              : To get user input for adding a product
'*Author                        : US Technology
'*Date Created                  : Sep-03-02
'*Date Last Modified            : Mar-18-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : InventoryFunctions
'*Components Used               : None
'*Functions Defined             : 1. ValidateProductCode
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Public gfrmGeneral                      As Form

Private fbCancelled                     As Boolean

Private fsCopyString                    As String
Private fbCanBeCopied                   As Boolean

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    fbCancelled = True
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates the user input.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()
    
    On Error GoTo ErrHandler
    
    txtProdCode.Text = Trim(txtProdCode.Text)
    txtBoxes.Text = Trim(txtBoxes.Text)
    txtWgt.Text = Trim(txtWgt.Text)
      
    'Prompt the user for a Product Code.
    
    If txtProdCode.Text = "" Then

        giErrMsg = ShowErrorMsg("AddPro012", Me.Caption, vbOKOnly)
        
        txtProdCode.SetFocus
        txtProdCode_GotFocus
        
        Exit Sub
        
    End If
    
    'Validate the Product Code with the database.
    
    If Not ValidateProductCode Then Exit Sub
    
    'Prompt the user for Ending Boxes value.
    
    If txtBoxes.Text = "" Then

        giErrMsg = ShowErrorMsg("AddPro002", Me.Caption, vbOKOnly)
        
        txtBoxes.SetFocus
        txtBoxes_GotFocus
    
    'Prompt the user for a valid Ending Boxes value.
    
    ElseIf txtBoxes.Text <> "" And Val(txtBoxes.Text) = 0 Then
    
        txtBoxes.Text = 0
        giErrMsg = ShowErrorMsg("AddPro013", Me.Caption, vbOKOnly)
        
        txtBoxes.SetFocus
        txtBoxes_GotFocus
        
    'Prompt the user for Ending Weight value.
    
    ElseIf txtWgt.Text = "" Then

        giErrMsg = ShowErrorMsg("AddPro003", Me.Caption, vbOKOnly)
        
        txtWgt.SetFocus
        txtWgt_GotFocus
        
    'Prompt the user for a valid Ending Weight value.
        
    ElseIf txtWgt.Text <> "" And Val(txtWgt.Text) = 0 Then
        
        txtWgt.Text = 0
        txtWgt.Text = Format(txtWgt.Text, "########0.00")
        
        giErrMsg = ShowErrorMsg("AddPro014", Me.Caption, vbOKOnly)
        
        txtWgt.SetFocus
'        txtWgt_LostFocus
    
    'Validate the Ending Weight value.
    
    ElseIf CDec(txtWgt.Text) > 99999999.99 Then

        giErrMsg = ShowErrorMsg("AddPro004", Me.Caption, vbOKOnly)
        txtWgt.SetFocus
        txtWgt_GotFocus
    
    Else
    
        txtWgt.Text = Format(txtWgt.Text, "########0.00")
        
        If Not gfrmGeneral Is Nothing Then
        
            'Check for the duplicate entry of the Product Code in the Daily
            'Production Grid. If not present, add the record to the Grid.
            
            If Not gfrmGeneral.AddRow(txtProdCode.Text, _
                                      txtBoxes.Text, _
                                      txtWgt.Text) Then
                                      
                txtProdCode.SetFocus
                txtProdCode_GotFocus
                Exit Sub
                
            End If
            
        End If
        
        Unload Me
        
    End If
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:
    
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("AddPro005", Me.Caption, vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Loads the form and sets status of cancel
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()

    fbCancelled = False
    
    'Set the default values for the Ending Boxes and Weight fields.
    
    txtBoxes.Text = "0"
    txtWgt.Text = "0.00"
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtBoxes_GotFocus()

    txtBoxes.SelStart = 0
    txtBoxes.SelLength = Len(txtBoxes.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtProdCode, sOrigValue
    txtProdCode = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_GotFocus()

    txtProdCode.SelStart = 0
    txtProdCode.SelLength = Len(txtProdCode.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the text in selected mode.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtWgt_GotFocus()

    txtWgt.SelStart = 0
    txtWgt.SelLength = Len(txtWgt.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the KeyPress event of "ProductCode"
'* Parameter Description    :   Keyascii value of the key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_KeyPress(KeyAscii As Integer)
    
    On Error GoTo ErrHandler
    
    ReturnKeyForCapsAndNumerics txtProdCode, KeyAscii
   
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:
    
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("AddPro007", Me.Caption, vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit

End Sub

'******************************************************************************
'* Functional Description   :   Checks with the database if the entered product
'*                              code is valid.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean value to return if valid or not.
'******************************************************************************

Private Function ValidateProductCode() As Boolean

Dim objProduct              As Object
Dim sProductExistsInMaster  As String
Dim lResult                 As Long
Dim sProductID              As String

    If fbCancelled Then Exit Function
    
    On Error GoTo ErrHandler
    sProductID = Trim(txtProdCode.Text)
    
    'Check whether the Product Code exists in the database.
    
    Set objProduct = CreateObject("InventoryFunctions.DailyProduction")
    lResult = objProduct.ProductCodeExistsInPlant(sProductID, _
                                                  gsPlantCode, _
                                                  sProductExistsInMaster)
    If lResult <> 0 Then GoTo ErrHandler

    'Prompt the user for a valid Product Code.
    
    If CLng(Mid(sProductExistsInMaster, 1, 1)) = 0 Then
    
        giErrMsg = ShowErrorMsg("AddPro001", Me.Caption, vbOKOnly)
        
        txtProdCode.SetFocus
        txtProdCode_GotFocus
        
        ValidateProductCode = False
        GoTo CleanUpAndExit
        
    End If
    
    ValidateProductCode = True
    
CleanUpAndExit:
    
    Set objProduct = Nothing
    Exit Function

ErrHandler:
    
    If lResult <> 0 Then
    
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("AddPro011", Me.Caption, vbOKOnly, gcolErrMsg)
        
    Else
    
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("AddPro008", Me.Caption, vbOKOnly, gcolErrMsg)
        
    End If
    
    GoTo CleanUpAndExit

End Function

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If CheckNegPastedSelPart(txtWgt.Text, _
                                    txtWgt.SelText, _
                                    txtWgt.SelStart, _
                                    Clipboard.GetText, 8, 2) Then

            txtWgt.Text = GetPastedText(txtWgt.Text, _
                                                txtWgt.SelText, _
                                                txtWgt.SelStart, _
                                                Clipboard.GetText)

        End If

    ElseIf KeyCode = 46 Then

        ' Check if delete key is pressed.
        ' This will not be got in the keypress event

        If Not CheckNegSelectedPart(txtWgt.Text, _
                                txtWgt.SelText, _
                                txtWgt.SelStart, 0, 8, 2) Then

            KeyCode = 0

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtWgt.Text)
    iSelLength = txtWgt.SelLength
    iSelStart = txtWgt.SelStart
    sSelText = txtWgt.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 8, 2, iSelLength) Then

        KeyAscii = 0

        If sSelText = "" Then

            txtWgt.SelStart = iSelStart
            txtWgt.SelLength = iSelLength

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String

    sText = Trim(txtWgt.Text)
    Call EvaluateNagativeNumber(sText)
    txtWgt.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Format the field to ####0.00.
'* Parameter Description    :   None
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_LostFocus()

Dim sText As String

    sText = Trim(txtWgt.Text)

    If RoundOffDecimalNumber(sText, 8, 2, True) Then

        txtWgt.Text = sText

    Else

        txtWgt.SetFocus
        Exit Sub

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = False
        fsCopyString = txtWgt.Text

        If CheckNegPastedSelPart(Trim(txtWgt.Text), txtWgt.SelText, _
                                txtWgt.SelStart, Clipboard.GetText, 8, 2) Then

            fbCanBeCopied = True

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtWgt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtWgt.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtWgt.Text)
    Call EvaluateNagativeNumber(sText)
    txtWgt.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtBoxes_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyV And Shift = 2 Then

        If InStr(Clipboard.GetText, ".") > 0 Then

            KeyCode = 0

        ElseIf Not IsNumeric(Clipboard.GetText) Then

            KeyCode = 0

        ElseIf Len(Clipboard.GetText) + Len(txtBoxes.Text) - txtBoxes.SelLength > 6 Then

            KeyCode = 0

        End If

    End If

End Sub


'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtBoxes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = MouseButtonConstants.vbRightButton Then

        fbCanBeCopied = True
        fsCopyString = txtBoxes.Text

        If InStr(Clipboard.GetText, ".") > 0 Then

            fbCanBeCopied = False

        ElseIf Not IsNumeric(Clipboard.GetText) Then

            fbCanBeCopied = False

        ElseIf Len(Clipboard.GetText) + Len(txtBoxes.Text) - txtBoxes.SelLength > 6 Then

            fbCanBeCopied = False

        End If

    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtBoxes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then

        If Not fbCanBeCopied Then

             txtBoxes.Text = fsCopyString
             fbCanBeCopied = True

        End If
    End If

    sText = Trim(txtBoxes.Text)
    Call EvaluateNagativeNumber(sText)
    txtBoxes.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   Allows only integers in this field.
'* Parameter Description    :   KeyAscii-Ascii value of key pressed.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtBoxes_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> Asc("-") Then
        
        If KeyAscii = 22 Then
        
            If InStr(Clipboard.GetText, ".") > 0 Then
    
                KeyAscii = 0
    
            ElseIf Not IsNumeric(Clipboard.GetText) Then
    
                KeyAscii = 0
    
            ElseIf Len(Clipboard.GetText) + Len(txtBoxes.Text) - txtBoxes.SelLength > 6 Then
    
                KeyAscii = 0
    
            End If
        
        Else
            If InStr(txtBoxes.Text, "-") > 0 Then
        
                ValidateForIntegers KeyAscii
                
            ElseIf Len(txtBoxes.Text) = 6 And KeyAscii > 31 Then
            
                KeyAscii = 0
                
            Else
                
                ValidateForIntegers KeyAscii
            
            End If
                
        End If
        
    ElseIf txtBoxes.SelStart <> 0 Then
        
        KeyAscii = 0
        
    End If

End Sub
