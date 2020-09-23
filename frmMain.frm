VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Chat"
   ClientHeight    =   2910
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   3735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   3735
   Begin VB.ListBox lstOnline 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":0CCA
      Left            =   2040
      List            =   "frmMain.frx":0CCC
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picProgress 
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   170
      Left            =   2070
      ScaleHeight     =   165
      ScaleWidth      =   1605
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2710
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.Frame fraBuffers 
      Height          =   2055
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
      Begin MSComctlLib.ImageList imgHot 
         Left            =   1440
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0CCE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1222
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1776
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1ACA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2172
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":281A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E92
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":392A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgNormal 
         Left            =   720
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F92
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":44E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":483A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4B8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4EE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5236
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":558A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5C02
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6146
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":669A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock wsUDP 
         Left            =   120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
   End
   Begin MSComctlLib.Toolbar tbChatOptions 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   2295
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgNormal"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "MS Sans Serif"
            ImageIndex      =   3
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stirke Thru"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "List"
            ImageIndex      =   8
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.Toolbar tbSend 
         Height          =   330
         Left            =   2520
         TabIndex        =   12
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonWidth     =   2090
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgNormal"
         HotImageList    =   "imgHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Send       "
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox txtChatIn 
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2778
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":6BEE
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgNormal"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Properties..."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete History"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2655
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   3069
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChatOut 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393217
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":6CB7
   End
   Begin VB.PictureBox picHide 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   3960
      ScaleHeight     =   615
      ScaleWidth      =   1575
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
      Begin VB.PictureBox picHide 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   1
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mFilePorts 
         Caption         =   "View Ports"
      End
      Begin VB.Menu mFileSp 
         Caption         =   "-"
      End
      Begin VB.Menu mFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mEditSp 
         Caption         =   "-"
      End
      Begin VB.Menu mEditProperties 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Private Enum Data_Arrival_Processes
    dapLogon = 0
    dapLogonRet = 1
    dapLogoff = 3
    dapChatMsg = 4
End Enum

Private Const f_SectionGeneral As String = "General"
Private Const f_KeyScreenName As String = "ScreenName"
Private Const f_KeyUDPPort As String = "UDPPort"

Private f_ScreenName As String
Private f_NetworkID As String
Private f_LongIP As String
Private f_ComputerName As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Call SendChat
End Sub

Private Sub SendChat()
On Error GoTo ErrHndlr
    Dim I As Integer
    If txtChatOut.Text <> "" Then
        Screen.MousePointer = vbHourglass
        Call SetPanelText("Sending...", 1)
        For I = 0 To lstOnline.ListCount - 1
            Call UDPSendData(Me.BaseIP & lstOnline.ItemData(I), _
                    dapChatMsg, Me.LongIP & "|" & Me.ScreenName & "|" & _
                    txtChatOut.TextRTF)
            Call PerCnt(I + 1, lstOnline.ListCount, picProgress)
        Next I
    End If
ErrHndlr:
    txtChatOut.Text = ""
    Screen.MousePointer = vbDefault
    Call SetPanelText("", 1)
End Sub

Private Sub Form_Load()
    Dim I As Integer
    For I = 0 To Screen.FontCount - 1
        tbChatOptions.Buttons(1).ButtonMenus.Add I + 1, , Screen.Fonts(I)
    Next I
    Me.Show
    DoEvents
    Me.NetworkID = GetNetworkID
    Me.ComputerName = wsUDP.LocalHostName
    Me.ScreenName = GetScreenName
    Me.NetworkID = GetNetworkID
    Me.ComputerName = wsUDP.LocalHostName
    Call SetUDPPort(GetUDPPort)
    Call Logon
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If frmMain.WindowState <> vbMinimized Then
        If frmMain.Width < 3855 Then
            frmMain.Width = 3855
        End If
        If frmMain.Height < 3600 Then
            frmMain.Height = 3600
        End If
        
        If lstOnline.Visible Then
            txtChatIn.Width = frmMain.ScaleWidth - lstOnline.Width + 10
        Else
            txtChatIn.Width = frmMain.ScaleWidth
        End If
        txtChatIn.Height = frmMain.ScaleHeight - tbMain.Height - txtChatOut.Height _
                            - tbChatOptions.Height - sbMain.Height
        lstOnline.Height = frmMain.ScaleHeight - tbMain.Height - txtChatOut.Height _
                            - tbChatOptions.Height - sbMain.Height
        lstOnline.Left = txtChatIn.Width - 10
        txtChatOut.Width = frmMain.ScaleWidth
        txtChatOut.Top = tbMain.Height + txtChatIn.Height - 10
        picProgress.Top = frmMain.ScaleHeight - 200
        picProgress.Left = frmMain.ScaleWidth - picProgress.Width - 320
    End If
End Sub

Private Sub PerCnt(ByVal percent, ByVal total, picture As Control)
On Error Resume Next
    Dim num$
    If total = 0 Then Exit Sub
    percent = percent / total * 100
    If percent > 0 Then picture.Visible = True
    If percent > 100 Or percent < 0 Then
        Exit Sub
    End If
    If Not picture.AutoRedraw Then
        picture.AutoRedraw = -1
    End If
    picture.Cls
    picture.ScaleWidth = 100
    picture.DrawMode = 10
    num$ = Format$(percent, "###") + "%"
    picture.CurrentX = 50 - picture.TextWidth(num$) / 2
    picture.CurrentY = (picture.ScaleHeight - picture.TextHeight(num$)) / 2
    picture.Print num$
    picture.Line (0, 0)-(percent, picture.ScaleHeight), , BF
    If percent >= 100 Then
        picture.Cls
        picture.Visible = False
    End If
    picture.Refresh
    DoEvents
End Sub

Private Function Parse(ByVal strString As String, lItemNum As Long) As String
    Dim arrItems() As String
    lItemNum = lItemNum - 1
    arrItems = Split(strString, "|", , vbTextCompare)
    If lItemNum <= UBound(arrItems) Then
        Parse = arrItems(lItemNum)
    End If
    Erase arrItems
End Function

Public Property Get BaseIP() As String
    BaseIP = Trim(Left(wsUDP.LocalIP, _
                    InStrRev(wsUDP.LocalIP, ".", , vbTextCompare)))
End Property

Public Property Get MyIP() As String
    Dim arrIP() As String
    arrIP = Split(wsUDP.LocalIP, ".", , vbTextCompare)
    MyIP = Trim(arrIP(3))
    Erase arrIP
End Property

Public Property Let ScreenName(strScreenName As String)
    f_ScreenName = strScreenName
End Property

Public Property Get ScreenName() As String
   ScreenName = f_ScreenName
End Property

Public Property Get LongIP() As String
    LongIP = wsUDP.LocalIP
End Property

Public Property Get NetworkID() As String
    NetworkID = f_NetworkID
End Property

Public Property Let NetworkID(strNetworkID As String)
    f_NetworkID = strNetworkID
End Property

Public Property Get ComputerName() As String
    ComputerName = f_ComputerName
End Property

Public Property Let ComputerName(strComputerName As String)
    f_ComputerName = strComputerName
End Property

Private Sub Form_Unload(Cancel As Integer)
    Call Logoff
    Call UnloadForms
End Sub

Private Function GetScreenName() As String
    GetScreenName = GetSetting(App.EXEName, f_SectionGeneral, f_KeyScreenName, "")
    If Trim(GetScreenName) = "" Then GetScreenName = Me.NetworkID
End Function

Private Function GetUDPPort() As String
    GetUDPPort = GetSetting(App.EXEName, f_SectionGeneral, f_KeyUDPPort, "")
    If Trim(GetUDPPort) = "" Then GetUDPPort = Str(5400)
End Function

Public Sub SaveScreenName(ByVal strScreenName As String)
    Call SaveSetting(App.EXEName, f_SectionGeneral, f_KeyScreenName, strScreenName)
End Sub

Public Sub SaveUDPPort(ByVal strUDPPort)
    Call SaveSetting(App.EXEName, f_SectionGeneral, f_KeyUDPPort, strUDPPort)
End Sub

Public Sub Logon()
On Error Resume Next
    Dim I As Integer
    Dim x As Integer: x = 255
    Call SetPanelText("Logging on...", 1)
    For I = 1 To x
        Call UDPSendData(Trim(Me.BaseIP & Trim(Str(I))), dapLogon, _
                            Me.MyIP & "|" & Me.ScreenName & "|")
        Call PerCnt(I, x, picProgress)
        DoEvents
    Next I
    Call SetPanelText("", 1)
End Sub

Public Sub Logoff()
On Error Resume Next
    Dim I As Integer
    Dim x As Integer: x = 255
    Call SetPanelText("Logging off...", 1)
    For I = 1 To x
        Call UDPSendData(Trim(Me.BaseIP & Trim(Str(I))), dapLogoff, _
                            Me.MyIP & "|" & Me.ScreenName & "|")
        Call PerCnt(I, x, picProgress)
        DoEvents
    Next I
    Call SetPanelText("", 1)
End Sub


Private Sub UnloadForms()
On Error Resume Next
    Dim oFrm As Form
    For Each oFrm In Forms
        Unload oFrm
    Next
End Sub

Private Sub SetPanelText(strText As String, intPanelNum As Integer)
    sbMain.Panels(intPanelNum).Text = strText
    DoEvents
End Sub

Public Sub SetUDPPort(intPortNum As Integer)
    wsUDP.RemotePort = intPortNum
    wsUDP.LocalPort = intPortNum
End Sub

Public Function GetNetworkID() As String
On Error Resume Next
    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        GetNetworkID = Left$(sBuffer, lSize)
    Else
        GetNetworkID = vbNullString
    End If
End Function

Private Sub mEditProperties_Click()
    Call frmProperties.Show(vbModal, frmMain)
End Sub

Private Sub mFileExit_Click()
    Call Unload(Me)
End Sub

Private Sub mFilePorts_Click()
    Call MsgBox("Local: " & wsUDP.LocalPort & vbCrLf & "Remote: " & wsUDP.RemotePort, _
        vbInformation, Me.Caption)
End Sub

Private Sub tbChatOptions_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Call ShowFonts
        Case 2
            txtChatOut.SelBold = Not txtChatOut.SelBold
        Case 3
            txtChatOut.SelItalic = Not txtChatOut.SelItalic
        Case 4
            txtChatOut.SelUnderline = Not txtChatOut.SelUnderline
        Case 5
            txtChatOut.SelStrikeThru = Not txtChatOut.SelStrikeThru
        Case 7
            lstOnline.Visible = Not lstOnline.Visible
            Call Form_Resize
    End Select
End Sub

Private Sub ShowFonts()
    Set frmFonts.RTB = txtChatOut
    Call frmFonts.Show(vbModal, frmMain)
End Sub

Private Sub tbChatOptions_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    txtChatOut.SelFontName = ButtonMenu
    tbChatOptions.Buttons(1).ToolTipText = ButtonMenu
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Call Logon
        Case 2
            Call Logoff
        Case 4
            Call frmProperties.Show(vbModal, frmMain)
        Case 5
            txtChatIn.Text = ""
        Case 7
            Call Unload(Me)
    End Select
End Sub

Private Sub tbSend_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Call SendChat
    End Select
End Sub

Private Sub txtChatOut_SelChange()
On Error Resume Next
    tbChatOptions.Buttons(1).ToolTipText = txtChatOut.SelFontName
    tbChatOptions.Buttons(2).Value = Abs(CInt(txtChatOut.SelBold))
    tbChatOptions.Buttons(3).Value = Abs(CInt(txtChatOut.SelItalic))
    tbChatOptions.Buttons(4).Value = Abs(CInt(txtChatOut.SelUnderline))
    tbChatOptions.Buttons(5).Value = Abs(CInt(txtChatOut.SelStrikeThru))
End Sub

Private Sub wsUDP_DataArrival(ByVal bytesTotal As Long)
    Dim InData As String
    wsUDP.GetData InData
    Select Case Parse(InData, 1)
        Case dapLogon
            If Not IsInItemData(Parse(InData, 2), lstOnline) Then
                Call lstOnline.AddItem(Parse(InData, 3))
                lstOnline.ItemData(lstOnline.ListCount - 1) = Parse(InData, 2)
            End If
            Call UDPSendData(Me.BaseIP & Parse(InData, 2), dapLogonRet, _
                            Me.MyIP & "|" & Me.ScreenName & "|")
        Case dapLogonRet
            If Not IsInItemData(Parse(InData, 2), lstOnline) Then
                Call lstOnline.AddItem(Parse(InData, 3))
                lstOnline.ItemData(lstOnline.ListCount - 1) = Parse(InData, 2)
            End If
        Case dapLogoff
            If IsInItemData(Parse(InData, 2), lstOnline) Then
                Call lstOnline.RemoveItem(ItemDataLocation(Parse(InData, 2), lstOnline))
            End If
        Case dapChatMsg
            Call AddNameToChat(Parse(InData, 3))
            Call AddMsgToChat(Right(InData, Len(InData) - _
                        Len(Parse(InData, 1)) - Len(Parse(InData, 2)) - _
                        Len(Parse(InData, 3)) - 3))
    End Select
End Sub

Private Sub AddNameToChat(strName As String)
    txtChatIn.SelStart = Len(txtChatIn.Text)
    txtChatIn.SelText = vbCrLf
    txtChatIn.SelStart = Len(txtChatIn.Text)
    txtChatIn.SelBold = True
    txtChatIn.SelColor = vbBlack
    txtChatIn.SelFontName = "MS Sans Serif"
    txtChatIn.SelFontSize = 8
    txtChatIn.SelItalic = False
    txtChatIn.SelStrikeThru = False
    txtChatIn.SelUnderline = False
    txtChatIn.SelText = "<< " & strName & " >>  "
    txtChatIn.SelStart = Len(txtChatIn.Text)
End Sub

Private Sub AddMsgToChat(strMsg As String, Optional blCarrRtn As Boolean = False)
    If blCarrRtn Then
        txtChatIn.SelStart = Len(txtChatIn.Text)
        txtChatIn.SelText = vbCrLf
    End If
    txtChatIn.SelStart = Len(txtChatIn.Text)
    txtChatIn.SelRTF = strMsg
    txtChatIn.SelStart = Len(txtChatIn.Text)
End Sub

Private Function IsInItemData(strItem As String, ctlControl As Control) As Boolean
    Dim I As Integer
    IsInItemData = False
    For I = 0 To ctlControl.ListCount - 1
        If strItem = ctlControl.ItemData(I) Then
            IsInItemData = True
            Exit For
        End If
    Next I
End Function

Private Function ItemDataLocation(strItem As String, ctlControl As Control) As Long
    Dim I As Integer
    ItemDataLocation = -1
    For I = 0 To ctlControl.ListCount - 1
        If strItem = ctlControl.ItemData(I) Then
            ItemDataLocation = I
            Exit For
        End If
    Next I
End Function

Private Sub UDPSendData(strRemoteHost As String, _
                dapMode As Data_Arrival_Processes, strMsg As String)
    
On Error Resume Next

    Dim strFullMsg As String
    strFullMsg = dapMode & "|" & strMsg
    wsUDP.RemoteHost = strRemoteHost
    Call wsUDP.SendData(strFullMsg)

End Sub
