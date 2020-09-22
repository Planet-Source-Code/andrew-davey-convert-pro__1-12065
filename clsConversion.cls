VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A class to handle converting units."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardblValue    As Double
Private mvarstrFromUnit As String
Private mvarstrToUnit   As String
Private mvarstrName     As String
Private mvarUnits       As clsUnits

Public Function getConversionData() As String
Attribute getConversionData.VB_Description = "Returns a string containing all the data for this conversion type."
    Dim strReturn As String
    
    strReturn = mvarstrName & vbCrLf
    strReturn = strReturn & mvarUnits.getUnitsString
    
    getConversionData = strReturn
End Function

Public Property Let strName(ByVal vData As String)
    mvarstrName = vData
End Property


Public Property Get strName() As String
    strName = mvarstrName
End Property



Public Property Get Units() As clsUnits
    Set Units = mvarUnits
End Property



Public Property Let strToUnit(ByVal vData As String)
Attribute strToUnit.VB_Description = "A string pointing to a unit in the units collection."
    mvarstrToUnit = vData
End Property


Public Property Get strToUnit() As String
    strToUnit = mvarstrToUnit
End Property



Public Property Let strFromUnit(ByVal vData As String)
Attribute strFromUnit.VB_Description = "A string pointing to a unit in the units collection."
    mvarstrFromUnit = vData
End Property


Public Property Get strFromUnit() As String
    strFromUnit = mvarstrFromUnit
End Property



Public Property Let dblValue(ByVal vData As Double)
Attribute dblValue.VB_Description = "The value being converted."
    mvardblValue = vData
End Property


Public Property Get dblValue() As Double
    dblValue = mvardblValue
End Property


Public Function Convert() As Double
Attribute Convert.VB_Description = "Returns a value after converting 'value' from 'FromUnits' to 'ToUnits'."
    Dim toUnit As clsUnit, fromUnit As clsUnit
    
    Set toUnit = mvarUnits(mvarstrToUnit)
    Set fromUnit = mvarUnits(mvarstrFromUnit)
            
    Convert = ((mvardblValue - fromUnit.dblOffset) * fromUnit.dblRelation) _
                / toUnit.dblRelation + toUnit.dblOffset
    
    Set toUnit = Nothing
    Set fromUnit = Nothing
End Function

Private Sub Class_Initialize()
    Set mvarUnits = New clsUnits
End Sub

Private Sub Class_Terminate()
    Set mvarUnits = Nothing
End Sub
