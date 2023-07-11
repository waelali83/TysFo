VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ArchInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements ObjectControl

'local variable(s) to hold property value(s)

Private mobjDataAccessObj       As DataAccess 'local copy

'MTS related declarations

Private oObjectContext          As ObjectContext
Private mbUseTransaction        As Boolean

Public Property Set DataAccessObj(ByVal objData As DataAccess)

    'used when assigning an Object to the property, on the left side of a Set statement.
    'Syntax: Set x.DataAccessObj = Form1
    
    Set mobjDataAccessObj = objData
    
End Property

Public Property Get DataAccessObj() As DataAccess

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.DataAccessObj
'MsgBox "begin get dataaccessobj()"
    If mobjDataAccessObj Is Nothing Then
    
        Set mobjDataAccessObj = CreateObject("USTRObjects.DataAccess")
        Set mobjDataAccessObj.ParentArch = Me
        
    End If
    Set DataAccessObj = mobjDataAccessObj
'MsgBox "end get dataaccessobj()"
    
End Property

Public Function LinkCustomComp(Comp As Variant)

    Comp.SetParentRef Me
    
End Function

Private Sub Class_Terminate()

    Set DataAccessObj = Nothing
    
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
