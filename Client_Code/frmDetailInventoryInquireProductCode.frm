VERSION 5.00
Begin VB.Form frmDetailInventoryInquireProductCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detail Inventory - Inquire"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPalletID 
      Height          =   1110
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      Begin VB.TextBox txtProdCode 
         Height          =   285
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1845
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2475
         TabIndex        =   3
         Top             =   660
         Width           =   1000
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1395
         TabIndex        =   2
         Top             =   660
         Width           =   1000
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
         Left            =   255
         TabIndex        =   4
         Top             =   270
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmDetailInventoryInquireProductCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'* File Name                     : frmDetailInventoryInquireProductCode.frm
'* File Description              : To get user input for product code
'* Author                        : US Technology
'* Date Created                  : Sep-03-02
'* Date Last Modified            : Mar-28-03
'* Version                       : 2.0
'* Layer                         : Client
'* Project Referenced            : None
'* Components Used               : None
'* Functions Defined             : None
'* Copyright                     : US Technology
'-------------------------------------------------------------------------------
'* Change History:
'* Change Code   Source(defect ID)    Change Description    Date      Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Public Event DetailInventoryInquire(ByVal sProductCode As String)

'*******************************************************************************
'* Functional Description   :   Unloads the form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'*******************************************************************************
'* Functional Description   :   Passes the entry Bayloc back to the calling
'*                              form if valid.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub cmdOK_Click()

    If Trim(txtProdCode.Text) <> "" Then
    
        'Pass the Product Code to the parent form.
        
        RaiseEvent DetailInventoryInquire(Trim(txtProdCode.Text))
        Unload Me
        
    Else
        
        'Prompt the user to enter a Product Code.
        
         ShowErrorMsg "DtlInvInqPrd001", Me.Caption, vbOKOnly
        txtProdCode.SetFocus
                
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   : Validates for alphanumerics.
'* Parameter Description    : None.
'* Return Type Description  : None.
'*******************************************************************************

Private Sub txtProdCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtProdCode, sOrigValue
    txtProdCode = sOrigValue
    
End Sub

'*******************************************************************************
'* Functional Description   :   Highlight the entry in the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtProdCode_GotFocus()
    
    txtProdCode.SelStart = 0
    txtProdCode.SelLength = Len(txtProdCode.Text)
    
End Sub

'*******************************************************************************
'* Functional Description   : Validates for alphanumerics.
'* Parameter Description    : Keyascii value of key pressed.
'* Return Type Description  : None.
'*******************************************************************************

Private Sub txtProdCode_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtProdCode, KeyAscii

End Sub
