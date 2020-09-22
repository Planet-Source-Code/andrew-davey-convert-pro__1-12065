VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConversions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A collection class of the class clsConversion."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsConversion"
Attribute VB_Ext_KEY = "Member0" ,"clsConversion"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private mCol As Collection


Public Function loadConversionData() As Boolean
    Dim Conversion As clsConversion, Unit As clsUnit
    Dim strPath As String
    Dim strConversion As String, strTemp As String
    
    On Error GoTo ErrHandler
    
    strPath = App.Path & "\"
    loadConversionData = True
    
    Open strPath & "index.dat" For Input As #1
        Do Until EOF(1)
            Input #1, strConversion
            If strConversion = "" Then
                Close #1
                loadConversionData = False
                Exit Function
            End If
                
            Open strPath & strConversion For Input As #2
                Set Conversion = New clsConversion
                
                With Conversion
                    Input #2, strTemp
                    .strName = strTemp
                    Do Until EOF(2)
                        Line Input #2, strTemp
                        Set Unit = New clsUnit
                        
                        Unit.loadFromString strTemp
                        .Units.Add Unit
                        
                        Set Unit = Nothing
                    Loop
                End With
                
                Add Conversion
                Set Conversion = Nothing
            Close #2
        Loop
    Close #1
    
    Exit Function
    
ErrHandler:
    Select Case Err.Number
    Case 457
        MsgBox "You have already loaded the data!", vbExclamation
    Case Else
        MsgBox "Error: " & CStr(Err.Number) & vbCrLf & _
                "Description: " & Err.Description
    End Select
    
    Err.Clear
End Function


Public Sub saveConversionsData()
Attribute saveConversionsData.VB_Description = "Saves all the conversion in this colection files and generates an index file."
    Dim Conversion As clsConversion
    Dim strConversions As String
    Dim strPath As String
    
    strPath = App.Path & "\"
    
    'Write conversion files.
    For Each Conversion In mCol
        With Conversion
            'Add file to conversions list.
            strConversions = strConversions & .strName & ".con" & vbCrLf
            
            'Write current data to new file.
            Open strPath & .strName & ".con" For Output As #1
                'Print #1, .strName
                Print #1, .getConversionData
            Close #1
        End With
    Next Conversion
    
    'Write the index.dat file.
    Open strPath & "index.dat" For Output As #1
        Print #1, strConversions
    Close #1
End Sub


Public Function Add(Conversion As clsConversion) As clsConversion
    mCol.Add Conversion, Conversion.strName

    'return the object created
    Set Add = Conversion
End Function

Public Property Get Item(vntIndexKey As Variant) As clsConversion
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

