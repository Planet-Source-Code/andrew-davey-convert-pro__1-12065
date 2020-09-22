VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Pro"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3600
   Icon            =   "frmTemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   615
      Left            =   2760
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "&Edit..."
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.ComboBox cboConversion 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Default         =   -1  'True
      Height          =   615
      Left            =   1080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   240
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   240
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   240
      Y1              =   97
      Y2              =   97
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   240
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label4 
      Caption         =   "C&onversion:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "&From:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "&To:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Menu mnuConvert 
      Caption         =   "&Convert"
      Begin VB.Menu mnuConvert_Convert 
         Caption         =   "&Convert Value"
      End
      Begin VB.Menu mnuConvert_Edit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuConvert_Break0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Contents 
         Caption         =   "&Help File..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp_WebSite 
         Caption         =   "&Web Site..."
      End
      Begin VB.Menu mnuHelp_Break0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnDragging As Boolean
Dim sngOffsetX As Single
Dim sngOffsetY As Single
Dim Conversions As clsConversions

Private Sub cboConversion_Click()
    'Update the other cbo's
    
    With Conversions(cboConversion.Text)
        loadUnits cboFrom
        loadUnits cboTo
    End With
End Sub

Private Sub cmdConvert_Click()
    Dim dblAns As Double
    With Conversions(cboConversion.Text)
        .strFromUnit = cboFrom.Text
        .strToUnit = cboTo.Text
        .dblValue = Val(txtFrom)
        dblAns = .Convert
        
        txtTo.Text = Round(dblAns, 6)
        
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
    cmdConvert.Picture = LoadResPicture("CONVERT", vbResBitmap)
    cmdEdit.Picture = LoadResPicture("EDIT", vbResBitmap)
    cmdExit.Picture = LoadResPicture("EXIT", vbResBitmap)
    Set Conversions = New clsConversions
    Conversions.loadConversionData
    loadConversions
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    Conversions.saveConversionsData
    Set Conversions = Nothing
End Sub

Sub loadConversions()
    Dim Conversion As clsConversion
    
    cboConversion.Clear
    For Each Conversion In Conversions
        cboConversion.AddItem Conversion.strName
    Next Conversion
    
    If cboConversion.ListCount >= 0 Then
        cboConversion.ListIndex = 0
    End If
End Sub

Sub loadUnits(cbo As ComboBox)
    Dim Unit As clsUnit
    
    cbo.Clear
    For Each Unit In Conversions(cboConversion.Text).Units
        cbo.AddItem Unit.strName
    Next Unit
    If cbo.Name = "cboFrom" Then
        cbo.ListIndex = IIf(cbo.ListCount < 0, -1, 0)
    Else
        cbo.ListIndex = IIf(cbo.ListCount < 1, -1, 1)
    End If
End Sub



Private Sub mnuConvert_Convert_Click()
    cmdConvert_Click
End Sub

Private Sub mnuConvert_Edit_Click()
    cmdEdit_Click
End Sub

Private Sub mnuConvert_Exit_Click()
    cmdExit_Click
End Sub


Private Sub mnuHelp_About_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuHelp_Contents_Click()
    ShellExecute Me.hWnd, "Open", "Help.htm", "", App.Path, 0
End Sub

Private Sub mnuHelp_WebSite_Click()
    ShellExecute Me.hWnd, "Open", "http://www.starsoftsoftware.co.uk", "", "", 0
End Sub
