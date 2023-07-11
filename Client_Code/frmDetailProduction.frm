VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDetailProduction 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detail Production Report"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAgedProductInventory 
      Height          =   1710
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   3210
      Begin VB.TextBox txtLoadId 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1695
         TabIndex        =   2
         Top             =   660
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpReportDate 
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   72810497
         CurrentDate     =   37498
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   825
         TabIndex        =   3
         Top             =   1140
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   1920
         TabIndex        =   4
         Top             =   1140
         Width           =   1000
      End
      Begin VB.Label lblLoadId 
         Caption         =   "Inbound Load ID"
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
         Left            =   240
         TabIndex        =   6
         Top             =   705
         Width           =   1515
      End
      Begin VB.Label lblReportDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Date"
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
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmDetailProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmDetailProduction.frm
'*File Description              : To get user input for printing report
'*Author                        : US Technology
'*Date Created                  : Sep-03-02
'*Date Last Modified            : Dec-11-02
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : Inventory Functions
'*Components Used               : USTriGrid
'*Functions Defined             : PrintDetailProductionRpt
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)   Change Description     Date       Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'*                                                     Oct-30-03  ibdkkam
'*Changes to add load id to inquire and pass it as a parameter to class that
'*calls the SP.
'******************************************************************************

Option Explicit

Private fbUnload As Boolean

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Call the PrintDetailProductionRpt function,
'*                              Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()
    
    fbUnload = True
    
    If (Trim(txtLoadId.Text) = "ALL" Or Trim(txtLoadId.Text) = "") Then
        txtLoadId.Text = "A"
    End If
    PrintDetailProductionRpt Format(dtpReportDate.Value, "MM/DD/YYYY"), txtLoadId.Text
    
    If fbUnload Then Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets default values on screen
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()
    
    fbUnload = True
    dtpReportDate.Value = Now
    txtLoadId.Text = "ALL"

End Sub
'******************************************************************************
'* Functional Description   :   Calls ReturnKeyForCapsAndNumerics Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLoadId_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtLoadId, KeyAscii
    
End Sub
'******************************************************************************
'* Functional Description   :   Selection of entry in text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLoadId_GotFocus()
   
    'Select the LoadId when focus is set.
    
    txtLoadId.SelStart = 0
    txtLoadId.SelLength = Len(txtLoadId.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Prints Detail Production Report
'* Parameter Description    :   sReportDate - Report Date selected.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub PrintDetailProductionRpt(ByVal sReportDate As String, _
                                     ByVal sLoadId As String)

Dim objDetailProductionRpt      As Object
Dim rsDtlProdInvData            As ADODB.Recordset
Dim lResult                     As Long

    On Error GoTo ErrHandler
    
    Set objDetailProductionRpt = _
            CreateObject("InventoryFunctions.DetailProductionRpt")

    ' Status bar

    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
        
    ' Get the production details
        
    lResult = objDetailProductionRpt.GetDetailProductionData(gsPlantCode, _
                                                        sReportDate, _
                                                        sLoadId, _
                                                        rsDtlProdInvData)
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
                                                        
    ' Handle the error
                                                        
    If lResult <> 0 Then GoTo ErrHandler
        Me.MousePointer = vbHourglass
        
    If Not rsDtlProdInvData.EOF Then
        PrintReport "Detail Production Report", "DetailProduction.rpt", _
                     rsDtlProdInvData, False
    Else
        ' No Data
        ShowErrorMsg "DetailPro003", Me.Caption, vbOKOnly
    End If
      
CleanUpAndExit:
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    Me.MousePointer = vbDefault
    
    Set objDetailProductionRpt = Nothing
    Set rsDtlProdInvData = Nothing
    Exit Sub
    
ErrHandler:
    
    If lResult <> 0 Then
       gcolErrMsg.Add lResult
       ShowErrorMsg "DetailPro001", Me.Caption, vbOKOnly, gcolErrMsg

    Else

       gcolErrMsg.Add Err.Description
       ShowErrorMsg "DetailPro002", Me.Caption, vbOKOnly, gcolErrMsg

    End If
    
    fbUnload = False
    GoTo CleanUpAndExit
    
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
