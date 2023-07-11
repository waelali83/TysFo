VERSION 5.00
Object = "{C0CF0B8C-9B38-40B4-A604-BB740046617B}#58.3#0"; "UstriGrid.ocx"
Begin VB.Form frmProductBreak 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Production - Prod. Breakout"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProductBreakOut 
      Height          =   2250
      Left            =   65
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   2640
         TabIndex        =   3
         Top             =   1800
         Width           =   1000
      End
      Begin USTriSuperGrid.USTriGrid ustgrdProdBreakOut 
         Height          =   1335
         Left            =   380
         TabIndex        =   2
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2355
         Columns         =   3
         FixedColumns    =   1
         FixedRows       =   0
         Rows            =   5
         TopRow          =   0
         Appearance      =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         BackColorFixed  =   -2147483633
         BackColorSel    =   -2147483635
         BorderStyle     =   0
         ForeColor       =   -2147483640
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         GridColorFixed  =   12632256
         GridLines       =   3
         GridLinesFixed  =   3
         GridLineWidth   =   1
         MergeCells      =   0
         TextStyle       =   0
         TextStyleFixed  =   0
         WordWrap        =   0   'False
         AllowUserResizing=   0
         ScrollBars      =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BandDisplay     =   0
         BackColorUnpopulated=   -2147483633
         AllowBigSelection=   -1  'True
         Object.WhatsThisHelpID =   0
         version         =   "1.0"
      End
      Begin VB.Label lblProdCodeValue 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblProdCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   1380
      TabIndex        =   0
      Top             =   600
      Width           =   1305
   End
End
Attribute VB_Name = "frmProductBreak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmProductBreak.frm
'*File Description              : To display the breakout details of a product.
'*Author                        : US Technology
'*Date Created                  : Aug-28-02
'*Date Last Modified            : Mar-18-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : None
'*Components Used               : USTriGrid
'*Functions Defined             : 1. FormatGrid
'*                                2. LoadProductBreakUp
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date        Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Private Enum ColumnName
    cnTransport = 0
    cnBox = 1
    cnWeight = 2
End Enum

Private Const GRID_ROW_HEIGHT = 300

Public fsProductCode As String
Public flTruckBoxes  As Long
Public fdblTruckWgt  As Double
Public flTotalBoxes  As Long
Public fdblTotalWgt  As Double

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub cmdOK_Click()
    
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Format the grid on form load.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()

    FormatGrid
    LoadProductBreakUp
    
End Sub

'******************************************************************************
'* Functional Description   :   Format the look and feel of the grid.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub FormatGrid()
    
Dim sTempData       As String
Dim iRowCount       As Integer
Dim iColIndex       As Integer
Const COLUMN_COUNT  As Integer = 3
        
    On Error GoTo ErrHandler
    
    With ustgrdProdBreakOut
        
        'Position the grid
        
        .Width = fraProductBreakOut.Width - 200
        .Left = 120
        
        .FixedRows = 1
        .FixedColumns = 1
        .RowHeight(0) = GRID_ROW_HEIGHT
        .RowHeight(1) = GRID_ROW_HEIGHT
        .RowHeight(2) = GRID_ROW_HEIGHT
        .RowHeight(3) = GRID_ROW_HEIGHT
        .WordWrap = True
        .Columns = COLUMN_COUNT
        
        'Format the Header Row
        
        .CellValue(0, cnTransport) = " "
        .CellValue(0, cnBox) = "Box"
        .CellValue(0, cnWeight) = "Weight"
        .CellValue(1, cnTransport) = "Truck"
        .CellValue(2, cnTransport) = "Conveyor"
        .CellValue(3, cnTransport) = "Total"
        
        'Set the Column Types
        
        .ColumnType(cnTransport) = Normal
        .ColumnType(cnBox) = Normal
        .ColumnType(cnWeight) = Normal
        
        .ColWidth(cnTransport) = .Width / 3.2
        .ColWidth(cnBox) = .Width / 3.8
        .ColWidth(cnWeight) = .Width / 2.4

        .ColAlignment(cnBox) = flexAlignRightCenter
        .ColAlignment(cnWeight) = flexAlignRightCenter
        .Row = 0
        
        For iColIndex = 0 To COLUMN_COUNT - 1
            .ColAlignmentFixed(iColIndex) = flexAlignCenterCenter
        Next
               
        For iColIndex = 0 To COLUMN_COUNT - 1
            .Column = iColIndex
            .CellFontBold = True
        Next
        
        For iColIndex = 1 To COLUMN_COUNT - 1
            .Column = iColIndex
            .Row = 3
            .CellFontBold = True
        Next
        
        For iRowCount = 1 To 3
            .Row = iRowCount
            .ColAlignmentFixed(cnTransport) = flexAlignLeftCenter
            .Column = 0
            .CellFontBold = True
        Next
        
        .Rows = 4
        .Column = 0
        
    End With
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:

    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("ProdBrk001", Me.Caption, vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Uploads data into the grid
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub LoadProductBreakUp()
    
    On Error GoTo ErrHandler
    
    'Display the breakout details of the selected Product Code.
    
    lblProdCodeValue.Caption = fsProductCode
    
    With ustgrdProdBreakOut
    
        .CellValue(1, cnBox) = flTruckBoxes
        .CellValue(1, cnWeight) = Format(fdblTruckWgt, "#0.00")
        .CellValue(2, cnBox) = flTotalBoxes - flTruckBoxes
        .CellValue(2, cnWeight) = Format(fdblTotalWgt - fdblTruckWgt, "#0.00")
        .CellValue(3, cnBox) = flTotalBoxes
        .CellValue(3, cnWeight) = Format(fdblTotalWgt, "#0.00")
        
    End With
    
CleanUpAndExit:
    
    Exit Sub
    
ErrHandler:

    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("ProdBrk002", Me.Caption, vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit
    
End Sub
