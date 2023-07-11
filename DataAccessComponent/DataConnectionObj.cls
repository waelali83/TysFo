VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "DataConnectionObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ObjectControl

Private mobjCommand             As ADODB.Command
Private mrsRecordSet            As ADODB.Recordset

'MTS related declarations

Private oObjectContext          As ObjectContext
Private mbUseTransaction        As Boolean

Private mbKeepTransActive       As Boolean
Private mbKeepConnectionOpen    As Boolean

Private mbLogging               As Boolean
Private msArchUser              As String
Private msArchUserPass          As String
Private msDSN                   As String
Private msFilePath              As String

' Added For OLEDB

Private msDataSource            As String
Private mbUseDSN                As Boolean

Public Function SetAbort() As String
     
Dim vCode As Variant

    On Error GoTo errorhandle:
    
    If Not (mobjCommand Is Nothing) Then
         If mobjCommand.ActiveConnection.Errors.Count > 0 Then
             Dim Error As ADODB.Error
             
             For Each Error In mobjCommand.ActiveConnection.Errors
                 vCode = "(" & Error.NativeError & ") " & Error.Description
                 If mbLogging Then WriteLog CStr(vCode)
             Next
         Else
             vCode = Err.Description
         End If
    End If
    
    'determine if this object was transactional
    
    If mbUseTransaction Then
    
        'MTS aborts transaction
        
        If Not (oObjectContext Is Nothing) Then
             oObjectContext.SetAbort
             If mbLogging Then WriteLog "Transaction aborted", "DataConnectionObj.SetAbort", gsUser
        End If
    End If
    
    'destroy objects
    
    Cleanup
    
    SetAbort = vCode
    
    Exit Function
    
errorhandle:

    If mbLogging Then WriteLog Err.Description, "DataConnectionObj.SetAbort", gsUser, Err
    On Error Resume Next
    
    If Not (oObjectContext Is Nothing) Then
        oObjectContext.SetAbort
    End If
    Set mobjCommand = Nothing
    Set oObjectContext = Nothing

End Function

Public Function Execute(ExpectingRS As Boolean, Optional RecordsAffected As Variant, Optional ByRef ErrorStr As String) As Long
         
Dim rstemp As ADODB.Recordset
    
    'Executes the prepared command object
    
    On Error GoTo errorhandle:
      
    'check if this SQL statement or SP returns a stored procedure
    
    If ExpectingRS = True Then
        
        '**create a client based cursor so that we can disconnect it from the db
        
        mobjCommand.ActiveConnection.CursorLocation = adUseClient
        
        'if this is a SQL statement, then we want to open static so we can
        'clone recordset if we need to
        
        If mobjCommand.CommandType = adCmdText Then
        
            Set rstemp = New Recordset
            rstemp.LockType = adLockBatchOptimistic
            Set rstemp.ActiveConnection = mobjCommand.ActiveConnection
            
            If mbLogging Then WriteLog "Executing SQL and capturing recordset", "DataConnectionObj.Execute", gsUser
            
            'execute the command and capture the new recordset
            
            rstemp.Open mobjCommand.CommandText, , adOpenStatic
            
        Else
        
            If mbLogging Then WriteLog "Executing Stored Procedure and capturing recordset", "DataConnectionObj.Execute", gsUser
            
            'execute the command and capture the new recordset
            
            Set rstemp = mobjCommand.Execute(, , adOpenForwardOnly)
            
        End If
        
        'set the temporary recordset to the output parameter
        Set mrsRecordSet = rstemp
        
    Else
    
        'since this execute does not produce a recordset, we can capture the rows returned
        
        mobjCommand.Execute RecordsAffected
        If mbLogging Then WriteLog "SQL command executed: RowAffected =" & CStr(RecordsAffected), "DataConnectionObj.Execute", gsUser
        
    End If
     
    Exit Function
    
errorhandle:

    Execute = Err
    If Not IsMissing(ErrorStr) Then ErrorStr = Err.Description
    If mbLogging Then WriteLog Err.Description, "DataConnectionObj.Execute", gsUser, Err
    Set mobjCommand = Nothing
    
End Function

Public Sub SetComplete()

   If mbUseTransaction And mbKeepTransActive = False Then
        If Not (oObjectContext Is Nothing) Then
            oObjectContext.SetComplete
            If mbLogging Then WriteLog gsUser & "committing transaction."
        End If
    End If
    
    'destroy objects
    
    Cleanup
    
End Sub

Public Sub Cleanup()
       
Dim iIndex As Integer
 
    ' Removes any parameters from the command objects parameter collection and destroys the MTS context object
    
    On Error GoTo errorhandle:
    
    If Not (mrsRecordSet Is Nothing) Then
        Set mrsRecordSet = Nothing
    End If
    
    If Not (mobjCommand Is Nothing) Then
    
        On Error Resume Next
        
      'delete parameters
      
        For iIndex = (mobjCommand.Parameters.Count - 1) To 0 Step -1
        
            'check if there was an error because the parameter collection wasn't setup
            
            If Err = -2147217839 Then
                On Error GoTo 0
                Exit For
            End If
            On Error GoTo 0
            mobjCommand.Parameters.Delete (iIndex)
                
        Next
        
        If mbKeepConnectionOpen = False Then
            Set mobjCommand = Nothing
        End If
        
    End If

    'check if an MTS oObjectContext was used
    
    If mbUseTransaction And mbKeepTransActive = False Then
        If Not (oObjectContext Is Nothing) Then
            Set oObjectContext = Nothing
        End If
    End If
    
    Exit Sub
    
errorhandle:

    If mbLogging Then WriteLog Err.Description, "DataConnectionObj.Cleanup", gsArchUser, Err
    
    Set mobjCommand = Nothing
    Set oObjectContext = Nothing
    
End Sub

Public Function Init(Optional lDSN As String, Optional ByVal DBUser As String, Optional ByVal DBUserPass As String) As Long
    
    'Description: Establishes a connection to the database
       
    On Error GoTo errorhandle:
  
    'check if the programmer is specifying a custom DSN
    'otherwise we will use the Arch DSN
    
    If Not IsMissing(lDSN) And Len(lDSN) > 0 Then
        msDSN = lDSN
    End If
  
    If mobjCommand Is Nothing Then
    
        'check if the architecture is running under the MTS context
        
        If mbUseTransaction Then
           If mbLogging Then WriteLog "Got MTS Context"
           Set mobjCommand = oObjectContext.CreateInstance("ADODB.Command")
           If mbLogging Then WriteLog "Created ADODB instance"
         Else
            Set oObjectContext = Nothing
            Set mobjCommand = New ADODB.Command
        End If
        
    End If
    
    'check if programmer has specified a username and pass
    
    If Not IsMissing(DBUser) Then
        If DBUser <> "" Then
            msArchUser = DBUser
            msArchUserPass = DBUserPass
        End If
    End If
    
    'check if the connection is already open
    
    If mobjCommand.ActiveConnection Is Nothing Then
    
        If mbLogging Then WriteLog " Connecting to " & msDSN & " datasource"
        
        'connect to database server
        
        If mbUseDSN Then
            mobjCommand.ActiveConnection = "DSN=" & msDSN & ";UID=" & msArchUser & ";PWD=" & msArchUserPass & ";"
        Else
            mobjCommand.ActiveConnection = "Provider=" & GC_PROVIDER & ";Password=" & msArchUserPass & ";User ID=" & msArchUser & ";Data Source=" & msDataSource & ";Persist Security Info=True"
        End If
        
    End If
        
    Exit Function
    
errorhandle:

    Init = Err
    If mbLogging Then WriteLog Err.Description, "DataConnectionObj.Init", gsUser, Err
    
    Set mobjCommand = Nothing
    Set oObjectContext = Nothing
    
End Function

Private Sub Class_Initialize()
        
Dim iLogging        As Integer
Dim sArchUser       As String
Dim sArchUserPass   As String
Dim iUseTransaction As Integer
Dim sDefaultDSN     As String
Dim sDataSource     As String
Dim iUseDSN         As Integer
Dim sFilePath       As String

    'Description: This sub retrieves all settings from the ArchSettings file so that logging, transactions, and connections are handle appropriately
    
    'get the app settings from the registry
    
    GetGlobalArchSettings iUseTransaction, iLogging, sArchUser, sArchUserPass, sDefaultDSN, sFilePath, iUseDSN, sDataSource

    'set class level variables
    
    mbLogging = iLogging
    mbUseTransaction = iUseTransaction
    msArchUser = sArchUser
    msArchUserPass = sArchUserPass
    msFilePath = sFilePath
        
    'set default DSN
    
    msDSN = sDefaultDSN
    
    ' set the OleDB Params
    
    msDataSource = sDataSource
    mbUseDSN = iUseDSN

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = True
    
End Function

Private Sub ObjectControl_Deactivate()

    Set oObjectContext = Nothing
    
End Sub

Private Sub ObjectControl_Activate()

    If mbUseTransaction Then
        Set oObjectContext = GetObjectContext
    End If
    
End Sub

Public Property Let ObjCommand(ByRef objCmd As ADODB.Command)
    
    Set mobjCommand = objCmd
    
End Property

Public Property Get ObjCommand() As ADODB.Command
    
    Set ObjCommand = mobjCommand
    
End Property

Public Property Let RsRecordSet(ByRef rsRecord As ADODB.Recordset)
    
    Set mrsRecordSet = rsRecord
    
End Property

Public Property Get RsRecordSet() As ADODB.Recordset
    
    Set RsRecordSet = mrsRecordSet
    
End Property

Public Property Let KeepTransActive(ByVal bKeepTransActive As Boolean)
    
    mbKeepTransActive = bKeepTransActive
    
End Property

Public Property Get KeepTransActive() As Boolean
    
    KeepTransActive = mbKeepTransActive
    
End Property

Public Property Let KeepConnectionOpen(ByVal bKeepConnectionOpen As Boolean)
    
    mbKeepConnectionOpen = bKeepConnectionOpen
    
End Property

Public Property Get KeepConnectionOpen() As Boolean
    
    KeepConnectionOpen = mbKeepConnectionOpen
    
End Property

