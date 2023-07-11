VERSION 5.00
Begin VB.Form frmProductCodeInquire 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Code - Inquire"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraInquire 
      Height          =   1065
      Left            =   65
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   1620
         TabIndex        =   2
         Top             =   615
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2700
         TabIndex        =   3
         Top             =   615
         Width           =   1000
      End
      Begin VB.TextBox txtProdCode 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblInquire 
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
         Index           =   1
         Left            =   255
         TabIndex        =   4
         Top             =   270
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmProductCodeInquire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmProductCodeInquire.frm
'*File Description              : To get Product Code as user input
'*Author                        : US Technology - Kavitha
'*Date Created                  : Aug-26-02
'*Date Last Modified            : Nov-28-02
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : None
'*Components Used               : USTriGrid
'*Functions Defined             : None
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Public Event ProductMasterInquire(ByVal sProductCode As String)

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Passes the user entry to the parent form and
'*                              Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()

    If Trim(txtProdCode.Text) <> "" Then
        RaiseEvent ProductMasterInquire(Trim(txtProdCode.Text))
    Else
        giErrMsg = ShowErrorMsg("ProCodInq001", Me.Caption, vbOKOnly)
        txtProdCode.SetFocus
        Exit Sub
    End If
    
CleanUpAndExit:
    
    Unload Me
    Exit Sub

ErrHandler:

   gcolErrMsg.Add Err.Description
   giErrMsg = ShowErrorMsg("ProCodInq002", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
   GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Converts any character entered to Upper Case
'* Parameter Description    :   Ascii value of the key pressed.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtProdCode, KeyAscii

End Sub

'******************************************************************************
'* Functional Description   :   Validates data on their change.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtProdCode, sOrigValue
    txtProdCode = sOrigValue
    
End Sub


