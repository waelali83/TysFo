VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "ParamCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ParamObj"
Attribute VB_Ext_KEY = "Member0" ,"ParamObj"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection

Private mCol As Collection

Public Function Add(ByVal sParameterName As String, _
                    ByVal iDataType As Integer, _
                    ByVal sValue As String, _
                    ByVal iInOutFlag As Integer, _
                    Optional ByRef vSize As Variant) As ParamObj
    
Dim objNewMember    As ParamObj
Dim bExpectingSize  As Boolean
   
    Set objNewMember = New ParamObj
    
    'set the properties passed into the method
    
    objNewMember.ParameterName = sParameterName
    objNewMember.DataType = iDataType
    objNewMember.Value = sValue
    objNewMember.InOutFlag = iInOutFlag
    
    '  to include the size parameter
                    
    bExpectingSize = IsMissing(vSize)
    If bExpectingSize Then
        objNewMember.Size = 50
    Else
     objNewMember.Size = vSize
    End If
                
     
    mCol.Add objNewMember, sParameterName

    'return the object created
    
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vIndexKey As Variant) As ParamObj
Attribute Item.VB_UserMemId = 0

    'used when referencing an element in the collection
    'vIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    
    Set Item = mCol(vIndexKey)
  
End Property



Public Property Get Count() As Long

    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    
    Count = mCol.Count
    
End Property


Public Sub Remove(vIndexKey As Variant)

    'used when removing an element from the collection
    'vIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant

    mCol.Remove vIndexKey
    
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    'this property allows you to enumerate
    'this collection with the For...Each syntax
    
    Set NewEnum = mCol.[_NewEnum]
    
End Property


Private Sub Class_Initialize()

    'creates the collection when this class is created
    
    Set mCol = New Collection
    
End Sub


Private Sub Class_Terminate()

    'destroys collection when this class is terminated
    
    Set mCol = Nothing
    
End Sub

