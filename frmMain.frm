VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Pro"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   3240
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Default         =   -1  'True
      Height          =   615
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboConversion 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "&To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Convert &From:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "C&onversion:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conversions As clsConversions


Private Sub cboConversion_Click()
    'Update the other cbo's
    
    With Conversions(cboConversion.Text)
        loadUnits cboFrom
        loadUnits cboTo
    End With
End Sub

Private Sub cmdConvert_Click()
    With Conversions(cboConversion.Text)
        .strFromUnit = cboFrom.Text
        .strToUnit = cboTo.Text
        .dblValue = Val(txtFrom)
        
        txtTo = CStr(.Convert)
    End With
End Sub

Private Sub cmdEdit_Click()
    frmEdit.Show 1
    'Clear old values.
    Set Conversions = Nothing
    'Load new ones.
    Set Conversions = New clsConversions
    Conversions.loadConversionData
    
    loadConversions
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Load button images.
    cmdConvert.Picture = LoadResPicture("Convert", vbResBitmap)
    cmdEdit.Picture = LoadResPicture("Edit", vbResBitmap)
    cmdExit.Picture = LoadResPicture("Exit", vbResBitmap)
    
    Set Conversions = New clsConversions
    Conversions.loadConversionData
    loadConversions
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Conversions = Nothing
End Sub

Sub loadConversions()
    Dim Conversion As clsConversion
    
    cboConversion.Clear
    For Each Conversion In Conversions
        cboConversion.AddItem Conversion.strName
    Next Conversion
    
    cboConversion.ListIndex = cboConversion.ListCount - 1
End Sub

Sub loadUnits(cbo As ComboBox)
    Dim Unit As clsUnit
    
    cbo.Clear
    For Each Unit In Conversions(cboConversion.Text).Units
        cbo.AddItem Unit.strName
    Next Unit
    cbo.ListIndex = cbo.ListCount - 1
End Sub
