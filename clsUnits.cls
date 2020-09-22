VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A collection class of units."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsUnit"
Attribute VB_Ext_KEY = "Member0" ,"clsUnit"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection


Public Function getUnitsString() As String
Attribute getUnitsString.VB_Description = "Returns a string containing the units data."
    Dim strReturn As String
    Dim Unit As clsUnit
    
    For Each Unit In mCol
        strReturn = strReturn & Unit.getString & vbCrLf
    Next Unit
    
    If strReturn <> "" Then
        'Chop off the last vbCrLf.
        strReturn = Left$(strReturn, Len(strReturn) - 2)
    End If
    
    getUnitsString = strReturn
End Function


Public Function Add(Unit As clsUnit) As clsUnit
    'create a new object
    Dim objNewMember As clsUnit
    Set objNewMember = New clsUnit

    'set the properties passed into the method
    objNewMember.strName = Unit.strName
    objNewMember.dblRelation = Unit.dblRelation
    objNewMember.dblOffset = Unit.dblOffset
    
    mCol.Add objNewMember, Unit.strName

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsUnit
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
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

