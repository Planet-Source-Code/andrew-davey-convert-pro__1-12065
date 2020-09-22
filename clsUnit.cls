VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to hold infomation about a unit."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarstrName As String 'local copy
Private mvardblRelation As Double 'local copy
Private mvardblOffset As Double 'local copy

Public Function getString() As String
Attribute getString.VB_Description = "Returns the name, relationship and offset in a comma delimited string."
    Dim strReturn As String
    
    strReturn = mvarstrName & "," & CStr(mvardblRelation) & "," & CStr(mvardblOffset)
    
    getString = strReturn
End Function

Public Function loadFromString(strData As String) As Boolean
Attribute loadFromString.VB_Description = "Loads the name, relationship and offset from a comma delimited string."
    Dim strValues() As String
    
    strValues = Split(strData, ",")
    
    mvarstrName = strValues(0)
    mvardblRelation = strValues(1)
    mvardblOffset = strValues(2)
End Function

Public Property Let dblOffset(ByVal vData As Double)
Attribute dblOffset.VB_Description = "The offset of this unit."
    mvardblOffset = vData
End Property


Public Property Get dblOffset() As Double
    dblOffset = mvardblOffset
End Property



Public Property Let dblRelation(ByVal vData As Double)
Attribute dblRelation.VB_Description = "The relation to the base unit."
    mvardblRelation = vData
End Property


Public Property Get dblRelation() As Double
    dblRelation = mvardblRelation
End Property



Public Property Let strName(ByVal vData As String)
Attribute strName.VB_Description = "The name of the unit."
    mvarstrName = vData
End Property


Public Property Get strName() As String
    strName = mvarstrName
End Property



