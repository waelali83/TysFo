VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ParamObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)

Private msParameterName     As String
Private miDataType          As Integer
Private msValue             As String
Private miSize              As Integer
Private miInOutFlag         As Integer

Public Property Let InOutFlag(ByVal iData As Integer)
    
    miInOutFlag = iData
    
End Property


Public Property Get InOutFlag() As Integer
    
    InOutFlag = miInOutFlag
    
End Property



Public Property Let Value(ByVal sData As String)
    
    msValue = sData
    
End Property


Public Property Get Value() As String
    
    Value = msValue
    
End Property

Public Property Let DataType(ByVal iData As Integer)
    
    miDataType = iData
    
End Property


Public Property Get DataType() As Integer
   
   DataType = miDataType
   
End Property



Public Property Let ParameterName(ByVal sData As String)
    
    msParameterName = sData
    
End Property


Public Property Get ParameterName() As String
    
    ParameterName = msParameterName
    
End Property

Public Property Get Size() As Integer
    
    Size = miSize
    
End Property

Public Property Let Size(ByVal iNewValue As Integer)
    
    miSize = iNewValue
    
End Property
