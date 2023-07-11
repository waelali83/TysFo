VERSION 5.00
Begin VB.Form frmProductBayTransfer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detail Inventory Transfer -  Product/Bay Transfer"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInquire 
      Height          =   2000
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   4410
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   2235
         TabIndex        =   5
         Top             =   1500
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   3315
         TabIndex        =   3
         Top             =   1500
         Width           =   1000
      End
      Begin VB.TextBox txtFrom 
         Height          =   285
         Left            =   1930
         MaxLength       =   6
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtProdCode 
         Height          =   285
         Left            =   1930
         MaxLength       =   10
         TabIndex        =   2
         Top             =   255
         Width           =   1815
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1930
         MaxLength       =   6
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Bay Loc."
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
         Left            =   510
         TabIndex        =   8
         Top             =   630
         Width           =   1170
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
         Left            =   505
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Bay Loc."
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
         Left            =   510
         TabIndex        =   6
         Top             =   990
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmProductBayTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmProductBayTransfer.frm
'*File Description              : To get the user input for transferring a
'*                                Product's Bay Location.
'*Author                        : US Technology
'*Date Created                  : Aug-28-02
'*Date Last Modified            : Mar-22-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : InventoryFunctions
'*Components Used               : None
'*Functions Defined             : 1. ValidateEntries
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Private fsProdCode              As String
Private fsFromBay               As String
Private fsToBay                 As String
Private fsShiftCode             As String

Private fbCancelClicked         As Boolean

'******************************************************************************
'* Functional Description   :   Sets the fbCancelClicked variable and unloads
'*                              the form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    fbCancelClicked = True
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates the entries and unloads the form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()

    txtProdCode.Text = Trim(txtProdCode.Text)
    txtFrom.Text = Trim(txtFrom.Text)
    txtTo.Text = Trim(txtTo.Text)
            
    If ValidateEntries Then
    
        fbCancelClicked = False
        Unload Me
        
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Returns True, if Cancel Button is pressed.
'                               Otherwise False.
'* Parameter Description    :   None.
'* Return Type Description  :   CancelClicked - Boolean.
'******************************************************************************

Public Property Get CancelClicked() As Boolean

    CancelClicked = fbCancelClicked
    
End Property

'******************************************************************************
'* Functional Description   :   Returns the Product Code entered by the user.
'* Parameter Description    :   None.
'* Return Type Description  :   Product Code - String.
'******************************************************************************

Public Property Get ProductCode() As String

    ProductCode = fsProdCode
    
End Property

'******************************************************************************
'* Functional Description   :   Sets the Product Code.
'* Parameter Description    :   Product Code - String.
'* Return Type Description  :   None.
'******************************************************************************

Public Property Let ProductCode(ByVal sProductCode As String)

    fsProdCode = sProductCode
    
End Property

'******************************************************************************
'* Functional Description   :   Returns the FromBayLoc entered by the user.
'* Parameter Description    :   None.
'* Return Type Description  :   FromBayLoc - String.
'******************************************************************************

Public Property Get FromBayLoc() As String

    FromBayLoc = fsFromBay
    
End Property

'******************************************************************************
'* Functional Description   :   Sets the FromBayLoc.
'* Parameter Description    :   FromBayLoc - String.
'* Return Type Description  :   None.
'******************************************************************************

Public Property Let FromBayLoc(ByVal sFromBayLoc As String)

    fsFromBay = sFromBayLoc
    
End Property

'******************************************************************************
'* Functional Description   :   Returns the ToBayLoc entered by the user.
'* Parameter Description    :   None.
'* Return Type Description  :   ToBayLoc - String.
'******************************************************************************

Public Property Get ToBayLoc() As String

    ToBayLoc = fsToBay
    
End Property

'******************************************************************************
'* Functional Description   :   Sets the ToBayLoc.
'* Parameter Description    :   ToBayLoc - String.
'* Return Type Description  :   None.
'******************************************************************************

Public Property Let ToBayLoc(ByVal sToBayLoc As String)

    fsToBay = sToBayLoc
    
End Property

'******************************************************************************
'* Functional Description   :   Sets status of Cancel as true and initialises
'*                              the textboxes values
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()

    fbCancelClicked = True
    txtProdCode.Text = fsProdCode
    txtFrom.Text = fsFromBay
    
End Sub

'******************************************************************************
'* Functional Description   :   Valides the ProductCode and BayLoc
'* Parameter Description    :   None
'* Return Type Description  :   Boolean.  True  - All are valid entries.
'*                                        False - Otherwise.
'******************************************************************************

Private Function ValidateEntries() As Boolean

Dim objPalletMoveReLocate   As Object
Dim rsDetails               As ADODB.Recordset
Dim lErrorCode              As Long
Dim iValid                  As Integer
Dim bValidBay               As Boolean
Dim bValidProduct           As Boolean

    On Error GoTo ErrHandler

    ValidateEntries = False
    Set objPalletMoveReLocate = CreateObject _
                                    ("InventoryFunctions.PalletMoveReLocate")
    
    'Check for valid Product Code entry.
    
    If txtProdCode.Text <> "" Then
    
        fsProdCode = txtProdCode.Text
        
        lErrorCode = objPalletMoveReLocate.ValidateProductCode _
                                    (gsPlantCode, fsProdCode, bValidProduct)
        
        'Server Error while validating Product Code.
    
        If lErrorCode <> 0 Then
        
            gcolErrMsg.Add lErrorCode
            giErrMsg = ShowErrorMsg("ProBayTrans007", _
                                    Me.Caption, _
                                    vbOKOnly, _
                                    gcolErrMsg)
            GoTo CleanUpAndExit
        End If
        
        'Validating Product Code.
    
        If Not bValidProduct Then
        
            giErrMsg = ShowErrorMsg("ProBayTrans004", Me.Caption, vbOKOnly)
            txtProdCode.SetFocus
            GoTo CleanUpAndExit
            
        Else
        
            If txtFrom.Text <> "" Then
            
                fsFromBay = txtFrom.Text
                
               'Validate the From Bay Location.

                lErrorCode = objPalletMoveReLocate.ChkBayLoc _
                                                        (gsPlantCode, _
                                                        fsFromBay, _
                                                        bValidBay)
                
               'Server Error in validating From Bay Location.

                If lErrorCode <> 0 Then
                    gcolErrMsg.Add lErrorCode
                    giErrMsg = ShowErrorMsg("ProBayTrans008", _
                                            Me.Caption, _
                                            vbOKOnly, _
                                            gcolErrMsg)
                    GoTo CleanUpAndExit
                End If
                 
               'Invalid Bay Location

                If Not bValidBay Then
                
                    giErrMsg = ShowErrorMsg("ProBayTrans005", _
                                             Me.Caption, _
                                             vbOKOnly)
                    txtFrom.SetFocus
                    GoTo CleanUpAndExit
                    
                Else
                
                    If txtTo.Text <> "" Then
                    
                        fsToBay = txtTo.Text

                        'Validate the To Bay Location

                        lErrorCode = objPalletMoveReLocate.ChkBayLoc _
                                                            (gsPlantCode, _
                                                             fsToBay, _
                                                             bValidBay)
                         
                        'Server Error in validating Bay Location.

                        If lErrorCode <> 0 Then
                        
                            gcolErrMsg.Add lErrorCode
                            giErrMsg = ShowErrorMsg("ProBayTrans009", _
                                                    Me.Caption, _
                                                    vbOKOnly, _
                                                    gcolErrMsg)
                            GoTo CleanUpAndExit
                            
                        End If
                           
                        'In Valid Bay Location

                        If Not bValidBay Then
                        
                            giErrMsg = ShowErrorMsg("ProBayTrans006", _
                                        Me.Caption, vbOKOnly)
                            txtTo.SetFocus
                            GoTo CleanUpAndExit
                            
                        End If
                        
                    Else
                    
                        giErrMsg = ShowErrorMsg("ProBayTrans001", _
                                                Me.Caption, _
                                                vbOKOnly)
                        txtTo.SetFocus
                        GoTo CleanUpAndExit
                        
                    End If
                    
                End If
                
            Else
            
                giErrMsg = ShowErrorMsg("ProBayTrans002", Me.Caption, vbOKOnly)
                txtFrom.SetFocus
                GoTo CleanUpAndExit
                
            End If
            
        End If
        
    Else
    
        giErrMsg = ShowErrorMsg("ProBayTrans003", Me.Caption, vbOKOnly)
        txtProdCode.SetFocus
        GoTo CleanUpAndExit
        
    End If
    
    ValidateEntries = True
    
CleanUpAndExit:
    
    'Clean up all objects
    
    Set objPalletMoveReLocate = Nothing
    Set rsDetails = Nothing
    Exit Function
    
ErrHandler:
    
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("ProBayTrans010", _
                            Me.Caption, _
                            vbOKOnly, _
                            gcolErrMsg)
    GoTo CleanUpAndExit

End Function

'******************************************************************************
'* Functional Description   :   Validates the changes in this text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtFrom_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtFrom, sOrigValue
    txtFrom = sOrigValue

End Sub

'******************************************************************************
'* Functional Description   :   Validates the changes in this text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTo_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtTo, sOrigValue
    txtTo = sOrigValue

End Sub

'******************************************************************************
'* Functional Description   :   Validates the changes in this text box.
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
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtFrom_GotFocus()
    
    txtFrom.SelStart = 0
    txtFrom.SelLength = Len(txtFrom.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Valides the Key entered.
'* Parameter Description    :   KeyAscii - Ascii value of key entered.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtFrom_KeyPress(KeyAscii As Integer)

    
    'Check if the first three characters are alphanumeric
    'and the remaining entries numeric.
    
    If Len(txtFrom.Text) < 3 Then
    
        ReturnKeyForCapsAndNumerics txtFrom, KeyAscii
        
    Else
    
        ReturnKeyForNumerics txtFrom, KeyAscii
        
    End If
   
End Sub

'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_GotFocus()
    
    txtProdCode.SelStart = 0
    txtProdCode.SelLength = Len(txtProdCode.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Valides the Key entered.
'* Parameter Description    :   KeyAscii - Ascii value of key entered.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_KeyPress(KeyAscii As Integer)
        
    ReturnKeyForCapsAndNumerics txtProdCode, KeyAscii
        
End Sub

'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTo_GotFocus()
    
    txtTo.SelStart = 0
    txtTo.SelLength = Len(txtTo.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Valides the Key entered.
'* Parameter Description    :   KeyAscii - Ascii value of key entered.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtTo_KeyPress(KeyAscii As Integer)

    'Check if the first three caracters are aplhanumeric
    'and the remaining entries numeric.
    
    If Len(txtTo.Text) < 3 Then
    
        ReturnKeyForCapsAndNumerics txtTo, KeyAscii
    
    Else
    
        ReturnKeyForNumerics txtTo, KeyAscii
    
    End If
    
       
End Sub
