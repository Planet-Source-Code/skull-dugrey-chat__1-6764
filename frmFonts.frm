VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFonts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fonts"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUnderline 
      Caption         =   "&Underline"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "&Color..."
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraSample 
      Caption         =   "Sample"
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   2880
      Width           =   2775
      Begin VB.TextBox txtSample 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "AaBbYyZz"
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstSize 
      Height          =   2010
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtSize 
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox lstFontStyle 
      Height          =   2010
      ItemData        =   "frmFonts.frx":0000
      Left            =   2040
      List            =   "frmFonts.frx":0010
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtFontStyle 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtFont 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox lstFont 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkStrikeThru 
      Caption         =   "Strike &Through"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label LblLabels 
      Caption         =   "Size"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label LblLabels 
      Caption         =   "Font Style"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblLabels 
      Caption         =   "Font"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Private Const LB_FINDSTRINGEXACT = &H1A2

Private f_RTB As Object

Public Function ListLocation(strItem As String, lstList As ListBox) As Long
    ListLocation = SendMessage(lstList.hwnd, LB_FINDSTRINGEXACT, -1, _
        ByVal CStr(strItem))
End Function

Public Property Get RTB() As Object
    Set RTB = f_RTB
End Property

Public Property Set RTB(ByRef rtbRTB As Object)
    Set f_RTB = rtbRTB
End Property

Private Sub chkStrikeThru_Click()
    txtSample.FontStrikethru = chkStrikeThru.Value
End Sub

Private Sub chkUnderline_Click()
    txtSample.FontUnderline = chkUnderline.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()
On Error Resume Next
    dlg.ShowColor
    txtSample.ForeColor = dlg.Color
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
    RTB.SelFontName = txtSample.FontName
    RTB.SelBold = txtSample.FontBold
    RTB.SelItalic = txtSample.FontItalic
    RTB.SelFontSize = txtSample.FontSize
    RTB.SelStrikeThru = chkStrikeThru.Value
    RTB.SelUnderline = chkUnderline.Value
    RTB.SelColor = txtSample.ForeColor
    Unload Me
End Sub

Private Sub Form_Load()
    Dim I As Integer
    If RTB Is Nothing Then
        MsgBox "Rich Text Box is not set.", vbCritical, Me.Caption
    Else
        If Screen.FontCount <> 0 Then
            For I = 0 To Screen.FontCount - 1
                lstFont.AddItem Screen.Fonts(I)
            Next I
        End If
        For I = 2 To 72
            lstSize.AddItem I
        Next I
        txtSample.FontName = RTB.SelFontName
        txtSample.FontBold = RTB.SelBold
        txtSample.FontItalic = RTB.SelItalic
        txtSample.FontSize = RTB.SelFontSize
        txtSample.FontStrikethru = RTB.SelStrikeThru
        txtSample.FontUnderline = RTB.SelUnderline
        txtSample.ForeColor = RTB.SelColor
        lstFont.ListIndex = ListLocation(RTB.SelFontName, lstFont)
        Select Case True
            Case RTB.SelBold = True And RTB.SelItalic = True
                lstFontStyle.ListIndex = 3
            Case RTB.SelBold = True
                lstFontStyle.ListIndex = 2
            Case RTB.SelItalic = True
                lstFontStyle.ListIndex = 1
            Case Else
                lstFontStyle.ListIndex = 0
        End Select
        lstSize.ListIndex = RTB.SelFontSize - 2
        chkUnderline.Value = RTB.SelUnderline
        chkStrikeThru.Value = RTB.SelStrikeThru
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RTB = Nothing
End Sub

Private Sub lstFont_Click()
    txtFont.Text = lstFont.List(lstFont.ListIndex)
    txtSample.FontName = txtFont.Text
End Sub

Private Sub lstFontStyle_Click()
    txtFontStyle.Text = lstFontStyle.List(lstFontStyle.ListIndex)
    txtSample.FontBold = False
    txtSample.FontItalic = False
    Select Case lstFontStyle.ListIndex
        Case 0
            txtSample.FontBold = False
        Case 1
            txtSample.FontItalic = True
        Case 2
            txtSample.FontBold = True
        Case 3
            txtSample.FontItalic = True
            txtSample.FontBold = True
    End Select
End Sub

Private Sub lstSize_Click()
    txtSize.Text = lstSize.List(lstSize.ListIndex)
    txtSample.FontSize = txtSize.Text
End Sub

Private Sub txtSample_GotFocus()
    cmdOK.SetFocus
End Sub
