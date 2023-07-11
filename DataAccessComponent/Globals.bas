Attribute VB_Name = "Globals"
Option Explicit

Public gsUser                   As String
Public gsArchUser               As String

Public Const GC_NotLoggedIn = -99
Public Const GC_DateNullMask = "01/01/1753"
Public Const GC_EncryptionKey = "USTRCODE"
Public Const GC_LogFileName = "USTRLogFile.log"
Public Const GC_ArchSettingsFileName = "USTRSettings.dat"

' For adding comment in the configuration files

Public Const GC_COMMENT_SEPERATOR = "'"
Public Const GC_PROVIDER = "MSDAORA.1"

Public Function SetADOParamDate(sDate As String) As Date

    ' set an empty date to a null mask
    
    If IsDate(sDate) Then
        If sDate = "12:00:00 AM" Then
            sDate = GC_DateNullMask
        End If
    Else
        sDate = GC_DateNullMask
    End If
    
    SetADOParamDate = CDate(sDate)


End Function

Public Sub SetADOParamString(ByRef sString As String, ByRef lLength As Integer)

    ' we need to pad empty strings because ADO 2.0 does not allow empty strings as varchar parameters
    ' the stored procedure being called should trim any strings being passed
                               
    If sString = "" Then
        sString = " "
    End If
    
    lLength = Len(sString)

End Sub

Public Function WriteLog(ByVal sLogDescrip As String, _
                         Optional ByVal sCallingProc As String, _
                         Optional ByVal sUserName As String, _
                         Optional ByVal lErrorCode As Long) As Long
    
Dim FileNum As Integer
   
    'get free file number
    
    FileNum = FreeFile
      
    Open App.Path & "\" & GC_LogFileName For Append As FileNum
   
   
    Write #FileNum, sLogDescrip, sCallingProc, sUserName, _
          lErrorCode, Format(Now, "mm-dd-yyyy-hh:mm:ss")
    Close #FileNum
    
    Exit Function
    
errorhandle:

    WriteLog = Err
    
    'close the file if it's still open
    
    On Error Resume Next
    Close #FileNum
    
End Function

Public Sub GetGlobalArchSettings(ByRef iUseTransactions As Integer, _
                                 ByRef iLogging As Integer, _
                                 ByRef sArchUser As String, _
                                 ByRef sArchUserPass As String, _
                                 ByRef sDefaultDSN As String, _
                                 ByRef sFilePath As String, _
                                 Optional ByRef iUseDSN As Integer, _
                                 Optional ByRef sDataSource As String)
'Local dat file variables

    On Error Resume Next
    
    'For and from the dat file
    
    Dim sDatUseTrans             As String
    Dim sDatLogging              As String
    Dim sDatUser                 As String
    Dim sDatUserPass             As String
    Dim sDatDefaultDSN           As String
    Dim sDatUseDSN               As String
    Dim sDatDataSource           As String
    Dim sDatFilePath             As String

    'Open the dat file for the values of the admin tool...
    
    Open App.Path & "\USTRSettings.dat" For Input As #1
    Line Input #1, sDatUseTrans
    Line Input #1, sDatLogging
    Line Input #1, sDatUser
    Line Input #1, sDatUserPass
    Line Input #1, sDatDefaultDSN
    Line Input #1, sDatFilePath
    Line Input #1, sDatUseDSN
    Line Input #1, sDatDataSource
    Close #1
    
    If (InStr(sDatLogging, GC_COMMENT_SEPERATOR)) = 0 Then
        sDatLogging = Trim(sDatLogging)
    Else
        sDatLogging = Trim(Mid(sDatLogging, 1, InStr(sDatLogging, GC_COMMENT_SEPERATOR) - 1))
    End If
    If (InStr(sDatUseTrans, GC_COMMENT_SEPERATOR)) = 0 Then
        sDatUseTrans = Trim(sDatUseTrans)
    Else
        sDatUseTrans = Trim(Mid(sDatUseTrans, 1, InStr(sDatUseTrans, GC_COMMENT_SEPERATOR) - 1))
    End If
    If (InStr(sDatUseDSN, GC_COMMENT_SEPERATOR)) = 0 Then
        sDatUseDSN = Trim(sDatUseDSN)
    Else
        sDatUseDSN = Trim(Mid(sDatUseDSN, 1, InStr(sDatUseDSN, GC_COMMENT_SEPERATOR) - 1))
    End If
    
    If UCase(sDatUseTrans) = "TRUE" Then
       iUseTransactions = -1
    Else
       iUseTransactions = 0
    End If
    
    If UCase(sDatLogging) = "TRUE" Then
       iLogging = -1
    Else
       iLogging = 0
    End If
    
    If UCase(sDatUseDSN) = "TRUE" Then
        iUseDSN = -1
    Else
        iUseDSN = 0
    End If
   
    ' set the record values
   
    If (InStr(sDatUser, GC_COMMENT_SEPERATOR)) = 0 Then
        sArchUser = Trim(sDatUser)
    Else
        sArchUser = Trim(Mid(sDatUser, 1, InStr(sDatUser, GC_COMMENT_SEPERATOR) - 1))
    End If
    
    If (InStr(sDatUserPass, GC_COMMENT_SEPERATOR)) = 0 Then
        sArchUserPass = Trim(sDatUserPass)
    Else
        sArchUserPass = Trim(Mid(sDatUserPass, 1, InStr(sDatUserPass, GC_COMMENT_SEPERATOR) - 1))
    End If
    
    If (InStr(sDatDefaultDSN, GC_COMMENT_SEPERATOR)) = 0 Then
        sDefaultDSN = Trim(sDatDefaultDSN)
    Else
        sDefaultDSN = Trim(Mid(sDatDefaultDSN, 1, InStr(sDatDefaultDSN, GC_COMMENT_SEPERATOR) - 1))
    End If
   
    If (InStr(sDatDataSource, GC_COMMENT_SEPERATOR)) = 0 Then
        sDataSource = Trim(sDatDataSource)
    Else
        sDataSource = Trim(Mid(sDatDataSource, 1, InStr(sDatDataSource, GC_COMMENT_SEPERATOR) - 1))
    End If
   
    If (InStr(sDatFilePath, GC_COMMENT_SEPERATOR)) = 0 Then
        sFilePath = Trim(sDatFilePath)
    Else
        sFilePath = Trim(Mid(sDatFilePath, 1, InStr(sDatFilePath, GC_COMMENT_SEPERATOR) - 1))
    End If
   
    gsArchUser = sArchUser
    gsUser = sArchUser
    
    'close the file
   
   Close #1
    
   Exit Sub
   
errorhandle:

    WriteLog Err.Description, "USTRSoftwareObjects_Globals.GetGlobalArchSettings", gsUser, Err.Number
    
    'close the file if it's still open
    
    On Error Resume Next
    Close #1

End Sub

