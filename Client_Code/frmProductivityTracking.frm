VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProductivityTracking 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Productivity Tracking"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraProdTracking 
      Height          =   2055
      Left            =   85
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   350
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   1845
         TabIndex        =   5
         Top             =   1560
         Width           =   1000
      End
      Begin VB.ComboBox cboShift 
         Height          =   345
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   53608451
         CurrentDate     =   38350
      End
      Begin VB.Label lblShift 
         Caption         =   "Shift"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDate 
         Caption         =   "Date"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmProductivityTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmInBoundOutBoundPalletReport.frm
'*File Description              : To get user input for printing report
'*Author                        : TCS
'*Date Created                  : Dec-29-04
'*Date Last Modified            : Dec-29-04
'*Version                       : 1.0
'*Layer                         : Client
'*Project Referenced            : LoadOutFunctions
'*Components Used               : None
'*Functions Defined             : 1. LoadPalletCount
'*                                2. GetPetDate
'*Copyright                     : TCS
'-------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)       Change Description    Date     Author
'*Initial Release                                      Dec-29-04     TCS

'******************************************************************************
Option Explicit

Dim sDate   As String
Dim sShift  As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

sDate = Format(dtpDate.Value, "MM/DD/YYYY")
sShift = Trim(cboShift.Text)

If CreateExcelSheet Then
    Unload Me

End If

End Sub

Private Sub Form_Load()

On Error GoTo ErrHandler
   ' Set default date
    dtpDate.Value = Date
    
    GetPetDate

    ' Load Shift Details
    
    cboShift.Clear
    cboShift.AddItem "A"
    cboShift.AddItem "B"
    cboShift.AddItem "C"
    
    cboShift.ListIndex = 0
    Exit Sub
ErrHandler:
  
  

End Sub




'******************************************************************************
'* Functional Description   :   Get the pet date of the plant.
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean value checking pet date is retrieved
'*                              or not.
'******************************************************************************

Private Function GetPetDate() As Boolean

Dim objPallet               As Object
Dim lResult                 As Long
Dim dtPetDate               As Date

    On Error GoTo ErrHandler
    
    lResult = 0
    GetPetDate = True
    
    Set objPallet = CreateObject("LoadOutFunctions.InboundOutBoundRpt")
    
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    
    'Change of status bar message in case of Loading from the database.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    ' Get the Pet Date of the selected plant.
    
    lResult = objPallet.GetPetDate(gsPlantCode, dtPetDate)
    
    'If there is any error goto error handler
    
    If lResult <> 0 Then GoTo ErrHandler
    
    dtpDate.Value = dtPetDate
    
CleanUpandExit:
    
    'Status text message reset.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Cleaning of the object
    
    Set objPallet = Nothing
    Exit Function

ErrHandler:
    
    If lResult = 0 Then
        
        'Displaying the error in case of VB Error.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "InOutPltRpt005", _
                     Me.Caption, _
                     vbOKOnly, _
                     gcolErrMsg
        
    Else
    
        'Displaying the error in case of Server side error.
        'Error in retrieving pet date
    
        gcolErrMsg.Add lResult
        ShowErrorMsg "InOutPltRpt003", _
                     Me.Caption, _
                     vbOKOnly, _
                     gcolErrMsg

    End If
    
    GetPetDate = False
    
    GoTo CleanUpandExit

End Function


Private Function CreateExcelSheet() As Boolean
Dim xlApp       As Object
Dim xlwBook     As Object
Dim xlSheet     As Object
Dim sFilename   As String
Dim fsoFileList As FileSystemObject
Dim fFile       As File

Dim rsProdDet   As ADODB.Recordset
Dim rsZoneDet   As ADODB.Recordset

Dim xlRow       As Long
Dim xlCol       As Long
Dim iVal        As Integer
Dim sFlag       As String

On Error GoTo ErrHandler


If LoadProdTrackingDetails(rsProdDet, rsZoneDet) = 0 Then


    
    If Not rsProdDet.EOF Then
    
        sFilename = rsProdDet("ColValue") & "_" & Format(dtpDate.Value, "MMDD") & "_" & cboShift.Text
        
        rsProdDet.MoveNext
        
        
    Else
    
        GoTo MsgHead
  
    End If
' File system Object
    Set fsoFileList = New FileSystemObject

' Chechking Mydocumnet folder is available or not
    If Not fsoFileList.FolderExists("C:\Documents and Settings\" & UCase(gsWinUserName) & "\My Documents\WMS") Then
    
        fsoFileList.CreateFolder "C:\Documents and Settings\" & UCase(gsWinUserName) & "\My Documents\WMS"
    End If
    
        sFilename = "C:\Documents and Settings\" & UCase(gsWinUserName) & "\My Documents\WMS\" & sFilename & ".Xls"
        
        

' copy template file
    If fsoFileList.FileExists(sFilename) Then
        fsoFileList.DeleteFile sFilename, True
        
    End If




    fsoFileList.CopyFile App.Path & "\Template_ProductivityTracking.xls", sFilename, True
    
    
    ' if sorcue file is a Read only file
    
    Set fFile = fsoFileList.GetFile(sFilename)
        fFile.Attributes = Normal
            
    
' Create Excel Application object
    Set xlApp = CreateObject("Excel.Application")

' Get current work details
    Set xlwBook = xlApp.Workbooks.Open(sFilename)

' Get Work sheet
    Set xlSheet = xlwBook.Worksheets.Item(1)

' Set Values to appropriate cells
 

 

    While Not rsProdDet.EOF
    
        'xlwBook.workSheets("INPUT Reqd").Range(rsProdDet("ColLoc")).Value = rsProdDet("ColValue")
        If IsNumeric(rsProdDet("ColValue")) Then
            xlSheet.Range(rsProdDet("ColLoc")).Value = IIf(rsProdDet("ColValue") = "0", "", Val(rsProdDet("ColValue")))
        Else
            xlSheet.Range(rsProdDet("ColLoc")).Value = rsProdDet("ColValue")
        End If
        rsProdDet.MoveNext
        
    Wend
    
    iVal = 66
    
    
    While Not rsZoneDet.EOF
        
    Select Case rsZoneDet("GroupName")
        Case Is = "A"
        iVal = iVal + 1
        xlSheet.Range(rsZoneDet("GroupName") & rsZoneDet("GroupID")).Value = rsZoneDet("ColValue")
        sFlag = "A"
        Case Else
        
          If sFlag <> rsZoneDet("GroupName") Then
            xlRow = xlRow
          End If
        
            xlRow = 0
            
            xlRow = xlSheet.Range("A67:A" & Trim(Str(iVal))).Find(rsZoneDet("Zone")).Cells.Row
            
            If xlRow > 0 Then
                xlSheet.Range(rsZoneDet("GroupName") & Trim(Str(xlRow))).Value = IIf(rsZoneDet("ColValue") = "0", "", Val(rsZoneDet("ColValue")))
            
            End If
            
            
            If sFlag <> rsZoneDet("GroupName") Then
                xlSheet.Range(rsZoneDet("GroupName") & Trim(Str(iVal + 1))).Value = "=Sum(" & rsZoneDet("GroupName") & "67:" & rsZoneDet("GroupName") & Trim(Str(iVal)) & ")"
                sFlag = rsZoneDet("GroupName")
            End If
            
    End Select
        rsZoneDet.MoveNext
    
    Wend
    
         
    
     xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders.Item(5).LineStyle = xlNone
     xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
     
    With xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
      
    With xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A68:A" & Trim(Str(iVal + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    
'
'    xlSheet.Range("A67:F" & Trim(Str(iVal))).Borders.Item(2).LineStyle = 7 ' xlSheet.Range("A66").Borders.Item(2).LineStyle
'    xlSheet.Range("A67:F" & Trim(Str(iVal))).Borders.Item(3).LineStyle = 0
    
' Color & Border style
    
    With xlSheet.Range("B68:F" & Trim(Str(iVal)))
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
    End With
    
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("B68:F" & Trim(Str(iVal))).Interior
        .ColorIndex = 34
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
    
    
  ' Total row Border
    
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlSheet.Range("A" & Trim(Str(iVal + 1)) & ":F" & Trim(Str(iVal + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    xlSheet.Range("A" & Trim(Str(iVal + 1))).Value = "=Sum(A67:A" & Trim(Str(iVal)) & ")"
    xlSheet.Range("B" & Trim(Str(iVal + 1))).Value = "=Sum(B67:B" & Trim(Str(iVal)) & ")"
    xlSheet.Range("C" & Trim(Str(iVal + 1))).Value = "=Sum(C67:C" & Trim(Str(iVal)) & ")"
    xlSheet.Range("D" & Trim(Str(iVal + 1))).Value = "=Sum(D67:D" & Trim(Str(iVal)) & ")"
    xlSheet.Range("E" & Trim(Str(iVal + 1))).Value = "=Sum(E67:E" & Trim(Str(iVal)) & ")"
    xlSheet.Range("F" & Trim(Str(iVal + 1))).Value = "=Sum(F67:F" & Trim(Str(iVal)) & ")"

    xlSheet.Range("A" & Trim(Str(iVal + 1))).Value = "CHECK"


' save Xls file
    
    xlwBook.Save
    xlwBook.Close SaveChanges:=True

' close the excel Object
    xlApp.Quit

' Intimate File name and path
    MsgBox "WMS input excel document has been created." & vbCrLf & "The path is '" & sFilename & "'", vbInformation, Me.Caption

Else

MsgHead:
    MsgBox "No data available", vbInformation, Me.Caption

End If


CleanUpandExit:
'Relase Objects
    Set xlApp = Nothing
    Set xlwBook = Nothing
    
CreateExcelSheet = True

Exit Function

ErrHandler:
    MsgBox "Error in client side report generation", vbInformation, Me.Caption
    CreateExcelSheet = False
    GoTo CleanUpandExit
End Function

Private Function LoadProdTrackingDetails(ByRef rsProdDet _
                                        As ADODB.Recordset, _
                                        ByRef rsZoneDetail _
                                        As ADODB.Recordset)
                                        
Dim objPallet               As Object
Dim lResult                 As Long
Dim dtPetDate               As Date

    On Error GoTo ErrHandler
    
    lResult = 0
    LoadProdTrackingDetails = True
    
    Set objPallet = CreateObject("Productivity.ProductivityTrack")
    
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    
    'Change of status bar message in case of Loading from the database.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
                                    "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    Me.MousePointer = vbHourglass
    DoEvents
    
    ' Get the Pet Date of the selected plant.
    
    lResult = objPallet.GetProdTracking(gsPlantCode, _
                                        sDate, _
                                        sShift, _
                                        rsProdDet, _
                                        rsZoneDetail)
                                        
    
    'If there is any error goto error handler
    LoadProdTrackingDetails = lResult
    If lResult <> 0 Then GoTo ErrHandler
    
    
    LoadProdTrackingDetails = lResult
CleanUpandExit:
    
    'Status text message reset.
    
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    frmMainMenu.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    
    'Cleaning of the object
    
    Set objPallet = Nothing
    Exit Function

ErrHandler:
    
    If lResult = 0 Then
        
        'Displaying the error in case of VB Error.
        
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "InOutPltRpt005", _
                     Me.Caption, _
                     vbOKOnly, _
                     gcolErrMsg
        
    Else
    
        'Displaying the error in case of Server side error.
        'Error in retrieving pet date
    
        gcolErrMsg.Add lResult
        ShowErrorMsg "InOutPltRpt003", _
                     Me.Caption, _
                     vbOKOnly, _
                     gcolErrMsg

    End If
    
    LoadProdTrackingDetails = False
    
    GoTo CleanUpandExit


End Function
