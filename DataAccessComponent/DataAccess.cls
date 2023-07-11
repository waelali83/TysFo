VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements ObjectControl

'local variable(s) to hold property value(s)

Private mobjParentArch      As ArchInterface 'local copy

Private mbLogging           As Boolean

Private msFilePath          As String

'MTS related declarations

Private oObjectContext      As ObjectContext
Private mbUseTransaction    As Boolean

Public Property Set ParentArch(ByVal vData As ArchInterface)

'used when assigning an Object to the property, on the left side of a Set statement.

    Set mobjParentArch = vData
    
End Property

Public Property Get ParentArch() As ArchInterface

'used when retrieving value of a property, on the right side of an assignment.

    Set ParentArch = mobjParentArch
    
End Property

Public Function GoSQL(ByVal sSQLText As String, _
                      Optional ByRef vResultset As Variant, _
                      Optional ByRef sDSN As String, _
                      Optional ByVal bUseArrayResult As Boolean, _
                      Optional ByRef vRowsAffected As Variant, _
                      Optional ByVal iCursorOptions As Integer, _
                      Optional ByRef vDataConnObj As Variant, _
                      Optional ByRef sErrorString As String) As Long

Dim lReturnCode         As Long
Dim objDataCon          As DataConnectionObj
Dim bCreateArrayResult  As Boolean
Dim sTempErrorString    As String
    
    On Error GoTo errorhandle
    
    Set objDataCon = CreateObject("USTRObjects.DataConnectionObj")
    
    'check if the programmer is expecting the results in array format
    
    If Not IsMissing(bUseArrayResult) Then
        If bUseArrayResult = True Then
            bCreateArrayResult = True
        End If
    End If
    
    If IsMissing(vDataConnObj) Then
        Set objDataCon = CreateObject("USTRObjects.DataConnectionObj")
    Else
        Set objDataCon = vDataConnObj
    End If
    
    'connect to database server
    
    lReturnCode = objDataCon.Init(sDSN)
    If lReturnCode <> 0 Then
        Err.Raise lReturnCode, "DataAcces.GoSQL", "Error connecting to database server."
    End If
    
    'prepare the command object to send SQL to database server
    
    objDataCon.ObjCommand.CommandText = sSQLText
    objDataCon.ObjCommand.CommandType = adCmdText
    
    If mbLogging Then WriteLog "Executing SQL: " & sSQLText, "DataAccess.GoSQL"

    'run the SQL statement
    
    lReturnCode = objDataCon.Execute(Not IsMissing(vResultset), vRowsAffected, sTempErrorString)
    If lReturnCode <> 0 Then
        Err.Raise lReturnCode, "DataAccess.GoSQL"
    End If
    
    'if the programmer is expecting a recordset, we need to pass it back
    
    If Not IsMissing(vResultset) Then
    
        'check if the user requested an array resultset
        
        If bCreateArrayResult Then
        
            'copy the recordset to a variant array
            
            vResultset = objDataCon.RsRecordSet.GetRows
            
            'close the recordset to free memory
            
            objDataCon.RsRecordSet.Close
            
        Else
            Set vResultset = objDataCon.RsRecordSet
        End If
        
        If mbLogging Then WriteLog "Setting recordset", "DataAccess.GoSQL"
        
    End If
    
    ' complete the transaction, if there is one and destroy the command object
    
    objDataCon.SetComplete
    
    ' shut down object if using MTS
    
    If mbUseTransaction Then oObjectContext.SetComplete
    
    'clear the objDataCon from memory
    
    Set objDataCon = Nothing
    
    Exit Function
    
errorhandle:
    
    GoSQL = Err
    
    'returning the sErrorString
    
    If Not IsMissing(sErrorString) Then
        sErrorString = sTempErrorString
    End If
    
    If mbLogging Then WriteLog Err.Description, "DataAccess.GoSQL", gsUser, Err
    
    On Error Resume Next
    
    'abort the transaction
    
    If Not (objDataCon Is Nothing) Then
        objDataCon.SetAbort
    End If
    If mbUseTransaction Then oObjectContext.SetAbort
    
    Set objDataCon = Nothing
    
    If ExpectingResultset Then
        Set vResultset = Nothing
    End If
    
End Function

Public Function GoSP(ByVal sSPName As String, _
                     Optional ByRef vParamColList As Variant, _
                     Optional ByRef vResultset As Variant, _
                     Optional ByRef sDSN As String, _
                     Optional ByRef vRowsAffected As Variant, _
                     Optional ByRef vDataConnObj As Variant, _
                     Optional ByRef sErrorString As String) As Long
   
Dim lReturnCode As Long
Dim lLongValue As Long
Dim dDoubleValue As Double
Dim iIndex As Integer
Dim iInnerIndex As Integer
Dim iDataType As Integer
Dim sTempErrorString As String
Dim iInOutFlag As Integer
Dim bParamsCreated As Boolean
Dim iLength As Integer
Dim objCurrParamCol As ParamCollection
Dim objDataCon As DataConnectionObj
Dim bExpectingResultset As Boolean
 
    On Error GoTo errorhandle:
    
    bExpectingResultset = Not IsMissing(vResultset)
   
    If IsMissing(vDataConnObj) Then
        Set objDataCon = CreateObject("USTRObjects.DataConnectionObj")
    Else
        Set objDataCon = vDataConnObj
    End If
    
   'connect to database server
   
    lReturnCode = objDataCon.Init(sDSN)
    If lReturnCode <> 0 Then
        Err.Raise lReturnCode, "DataAcces.GoSQL", "Error connecting to database server."
    End If
    
    'set flag to indicate that the parameters have not been created yet
    
    bParamsCreated = False
     
    'prepare the command object to send SQL to database server
    
    objDataCon.ObjCommand.CommandText = sSPName
    objDataCon.ObjCommand.CommandType = adCmdStoredProc
    
    If mbLogging Then WriteLog "Creating custom parameters for sp..." & sSPName, "DataAccess.GoSP"

    'check if user passed in an array of parameter collections
    
    If Not IsMissing(vParamColList) Then
        For iIndex = 1 To vParamColList.Count
        
              'get a reference to the current paramter set
              
            Set objCurrParamCol = vParamColList(iIndex)
              
            If mbLogging Then WriteLog "Looping through vParamColList " & CStr(iIndex), "DataAccess.GoSP"
              
              'set all custom parameters
            For iInnerIndex = 1 To objCurrParamCol.Count
              
                With objCurrParamCol(iInnerIndex)
                  
                    'get a copy of the Parameters data type and Inouttype
                    
                    iDataType = .DataType
                    
                    ' TO CHECK WHETHER THE PARAMETER IS OF INPUT OR OUTPUT TYPE
              
                    iInOutFlag = .InOutFlag
                            
                    If mbLogging Then WriteLog "Param: " & .ParameterName & " " & CStr(iInnerIndex) & " iDataType is " & iDataType & " value = " & .Value, "DataAccess.GoSP"
                  
                    If Not bParamsCreated Then
                         
                        'CHECKING WHETHER INPUT OR OUTPUT
                         
                        If iInOutFlag = 1 Then
                        
                            If iDataType = adDBTimeStamp Or iDataType = adDBDate Then
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, , SetADOParamDate(.Value))
                            ElseIf iDataType = adInteger Then
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, LenB(lLongValue), CLng(.Value))
                            ElseIf iDataType = adSmallInt Then
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, LenB(.Value), CInt(.Value))
                            ElseIf iDataType = adDouble Then
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, LenB(dDoubleValue), CDbl(.Value))
                            ElseIf iDataType = adVarChar Or iDataType = adLongVarChar Then
                            
                                'we need to pad strings because ADO 2.0 does not allow empty strings as varchar parameters
                                'the stored procedure being called should trim any strings being passed
                                
                                SetADOParamString .Value, iLength
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, iLength, .Value)
                            Else
                                objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, .InOutFlag, LenB(.Value), .Value)
                            End If
                        Else
                            
                            'TO CREATE THE OUTPUTPARAMETER
                                
                            If iDataType = adDBTimeStamp Or iDataType = adDBDate Then
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, iInOutFlag, .Size)
                                ElseIf iDataType = adInteger Then
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, iInOutFlag, .Size)
                                ElseIf iDataType = adSmallInt Then
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, iInOutFlag, .Size)
                                ElseIf iDataType = adDouble Then
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, iInOutFlag, .Size)
                                ElseIf iDataType = adVarChar Or iDataType = adLongVarChar Then
                                
                                    'we need to pad strings because ADO 2.0 does not allow empty strings as varchar parameters
                                    'the stored procedure being called should trim any strings being passed
                                    
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, .DataType, .InOutFlag, .Size)
                                
                                Else
                                    objDataCon.ObjCommand.Parameters.Append objDataCon.ObjCommand.CreateParameter(.ParameterName, iDataType, iInOutFlag, .Size)
                            End If
                        End If

                    Else
                        If iDataType = adVarChar Or iDataType = adLongVarChar Then
                        
                         'we need to pad strings because ADO 2.0 does not allow empty strings as varchar parameters
                         'the stored procedure being called should trim any strings being passed
                           
                           objDataCon.ObjCommand.Parameters(.ParameterName).Size = .Size
                           objDataCon.ObjCommand.Parameters(.ParameterName) = .Value
                           
                        ElseIf iDataType = adDBTimeStamp Or iDataType = adDBDate Then
                           objDataCon.ObjCommand.Parameters(.ParameterName) = SetADOParamDate(.Value)
                        Else
                            objDataCon.ObjCommand.Parameters(.ParameterName) = .Value
                        End If
                    End If
                  
                End With
                  
            Next
          
              'set flag so that parameters are not recreated
              
            bParamsCreated = True
                  
            If mbLogging Then WriteLog "Executing stored procedure.", "DataAccess.GoSP"

            'run the stored procedure
            
            lReturnCode = objDataCon.Execute(bExpectingResultset, , sTempErrorString)
            
            If lReturnCode <> 0 Then
                Err.Raise lReturnCode, "DataAccess.GoSP"
            End If
          
            'if there was an error stop sending data to the sp
          
            If lReturnCode <> 0 Then
                Exit For
            End If
              
            'get output variables and assign them to the parameter object
            
            For k = 1 To objCurrParamCol.Count
                If objCurrParamCol(k).InOutFlag = adParamOutput Or objCurrParamCol(k).InOutFlag = adParamInputOutput Then
                    objCurrParamCol(k).Value = objDataCon.ObjCommand(objCurrParamCol(k).ParameterName) & ""
                    If mbLogging Then WriteLog "Getting output param " & objCurrParamCol(k).ParameterName & " value - " & objCurrParamCol(k).Value, "DataAccess.GoSP"
                End If
            Next
            
        Next
        
    Else 'no custom parameters
    
        If mbLogging Then WriteLog "Executing stored procedure.", "DataAccess.GoSP"
            
           'run the stored procedure
           
        lReturnCode = objDataCon.Execute(bExpectingResultset, , sTempErrorString)
        If lReturnCode <> 0 Then
          Err.Raise lReturnCode, "DataAccess.GoSP"
        End If
            
    End If
    
    'set return value
    
    GoSP = lReturnCode
    
    'check if the call was successful
    
    If lReturnCode = 0 Then
    
        'if the programmer is expecting a recordset, we need to pass it back
        
        If bExpectingResultset Then
            Set vResultset = objDataCon.RsRecordSet
            If mbLogging Then WriteLog "Setting recordset object.", "DataAccess.GoSP"
        End If
    Else
        Err.Raise lReturnCode, "DataAccess.GoSP", "Error calling stored procedure " & sSPName
    End If
    
    'complete the transaction, if there is one and destroy the command object
    
    objDataCon.SetComplete
    
    'shut down object if using MTS
    
    If mbUseTransaction Then oObjectContext.SetComplete
    
    'clear the objDataCon from memory
    
    Set objDataCon = Nothing
    
    Exit Function
    
errorhandle:
        
    GoSP = Err
    If mbLogging Then WriteLog Err.Description, "DataAccess.GoSP", gsUser, Err
    If Not IsMissing(sErrorString) Then
        sErrorString = sTempErrorString
    End If
    
    On Error Resume Next
    
    'abort the transaction
    
    If Not (objDataCon Is Nothing) Then
        objDataCon.SetAbort
    End If
    
    'shut down object if using MTS
    
    If mbUseTransaction Then oObjectContext.SetAbort
    
    Set objDataCon = Nothing
    If bExpectingResultset Then
        Set vResultset = Nothing
    End If
    
End Function

Public Function GoSPGroup(ByVal sSPNameList As Variant, _
                          ByRef colParamListGroup As Variant, _
                          Optional ByRef sDSN As String) As Long

Dim ReturnCode As Long
Dim DataObj As DataConnectionObj
Dim ExpectingResultset As Boolean
Dim RowsAffected As Long
    
    On Error GoTo errorhandle:
    
    'create data connection obj
    
    If mbUseTransaction Then
        Set DataObj = oObjectContext.CreateInstance("USTRObjects.DataConnectionObj")
    Else
       Set DataObj = CreateObject("USTRObjects.DataConnectionObj")
    End If
          
    DataObj.KeepConnectionOpen = True
    DataObj.KeepTransActive = True
    
   'connect to database server
   
    ReturnCode = DataObj.Init(sDSN)
    If ReturnCode <> 0 Then
        Err.Raise ReturnCode, "DataAcces.GoSPGROUP", "Error connecting to database server."
    End If
    
    For i = 1 To colParamListGroup.Count
        If i = colParamListGroup.Count Then
            DataObj.KeepConnectionOpen = False
            DataObj.KeepTransActive = False
        End If
        
        retval = GoSP(CStr(sSPNameList(i)), colParamListGroup.Item(i), , sDSN, RowsAffected, DataObj)
        If retval <> 0 Then
            Exit For
        End If
    Next
    
    If retval <> 0 Then
        Err.Raise retval, , "Error running " & CStr(sSPNameList(i))
    End If
                   
    'complete the transaction, if there is one and destroy the command object
    
    DataObj.SetComplete
    
    'shut down object if using MTS
    
    If mbUseTransaction Then oObjectContext.SetComplete
    
    'clear the DataObj from memory
    
    Set DataObj = Nothing
    
    Exit Function
    
errorhandle:
        
    GoSPGroup = Err
    If mbLogging Then WriteLog Err.Description, "DataAccess.GoSPGroup", gsUser, Err
    
    
    On Error Resume Next
    
    'shut down object if using MTS
    
    If mbUseTransaction Then oObjectContext.SetAbort
    
    Set DataObj = Nothing
    
End Function

Public Function GetParamCollectionList(ByRef ParamColList As Variant, _
                                    ByVal ParameterColCount As Integer)
       
Dim colTempParamList As New Collection
Dim iIndex As Integer

    For iIndex = 1 To ParameterColCount
        'add a new paramCollection to the collection list.
        colTempParamList.Add New ParamCollection
    Next
   
   'pass ParamColList back to calling procedure
   
   Set ParamColList = colTempParamList
   
End Function

Public Function GetFilePath(ByRef FilePath As String)
    
    FilePath = msFilePath

End Function

Private Sub Class_Initialize()
     
Dim iLogging As Integer
Dim sArchUser As String
Dim sArchUserPass As String
Dim iUseTransaction As Integer
Dim sDefaultDSN As String
Dim sFilePath As String
    
    'get the app settings from the File
    
    GetGlobalArchSettings iUseTransaction, iLogging, sArchUser, sArchUserPass, sDefaultDSN, sFilePath
    
    mbLogging = iLogging
    mbUseTransaction = iUseTransaction
    msFilePath = sFilePath
    
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


