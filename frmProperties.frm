VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtComputerName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtNetworkID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtUDPPort 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   4320
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":0354
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgNorm 
      Left            =   3720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":06A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperties.frx":09FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin MSComctlLib.Toolbar tbSetup 
      Height          =   330
      Left            =   1065
      TabIndex        =   6
      Top             =   2520
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   582
      ButtonWidth     =   1826
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgNorm"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&OK    "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel "
            ImageIndex      =   2
            Object.Width           =   500
         EndProperty
      EndProperty
   End
   Begin VB.Label lblLabels 
      Caption         =   "Network ID"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Computer Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Local IP"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Screen Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "&UDP Port"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrHndlr
    Screen.MousePointer = vbHourglass
    Call SaveNewSettings
    Call Unload(Me)
ErrHndlr:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    txtNetworkID.Text = frmMain.NetworkID
    txtIP.Text = frmMain.LongIP
    txtComputerName.Text = frmMain.ComputerName
    txtName.Text = frmMain.ScreenName
    txtUDPPort.Text = frmMain.wsUDP.LocalPort
    Call FieldsCanUpdate
End Sub

Private Sub tbSetup_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Call cmdOK_Click
        Case 2
            Call cmdCancel_Click
    End Select
End Sub

Private Sub txtName_Change()
    Call FieldsCanUpdate
End Sub

Private Sub txtUDPPort_Change()
    Call FieldsCanUpdate
End Sub

Private Sub txtUDPPort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3
            'ctrl c
        Case 8
            'backspace
        Case 24
            'ctrl x
        Case Is < 48
            KeyAscii = 0
        Case Is > 57
            KeyAscii = 0
    End Select
End Sub

Private Function SaveNewSettings()
    If txtName.Text <> frmMain.ScreenName Or _
        txtUDPPort.Text <> frmMain.wsUDP.LocalPort Then
        
        Call frmMain.Logoff
        frmMain.ScreenName = Trim(txtName.Text)
        Call frmMain.SaveScreenName(frmMain.ScreenName)
        Call frmMain.SaveUDPPort(txtUDPPort.Text)
        frmMain.wsUDP.Close
        Call frmMain.SetUDPPort(txtUDPPort.Text)
        Call frmMain.Logon
    
    End If
End Function

Private Function FieldsCanUpdate() As Boolean
    FieldsCanUpdate = False
    Select Case True
        Case Len(Trim(txtName.Text)) = 0
        Case Len(txtUDPPort.Text) = 0
        Case Else
            FieldsCanUpdate = True
    End Select
    cmdOK.Enabled = FieldsCanUpdate
    tbSetup.Buttons(1).Enabled = FieldsCanUpdate
End Function
