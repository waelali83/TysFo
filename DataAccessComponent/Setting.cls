VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'MTS related declarations

Implements ObjectControl

Private oObjectContext As ObjectContext
Private mbUseTransaction As Boolean

Private Sub Class_Initialize()

    'Description: This sub retrieves all settings from the ArchSettings file so that logging, transactions, and connections are handle appropriately
           
    Dim iLogging        As Integer
    Dim sArchUser       As String
    Dim sArchUserPass   As String
    Dim iUseTransaction As Integer
    Dim sDefaultDSN     As String
    Dim sFilePath       As String
        
    'get the app settings from the config file
    
    GetGlobalArchSettings iUseTransaction, iLogging, sArchUser, sArchUserPass, sDefaultDSN, sFilePath
   
    mbUseTransaction = iUseTransaction

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

