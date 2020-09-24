VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   2670
   ClientTop       =   2445
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "udpchat1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6795
   Begin VB.CommandButton Command2 
      Caption         =   "A"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2040
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   3360
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock udp 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   1001
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   375
         Left            =   5880
         MousePointer    =   10  'Up Arrow
         TabIndex        =   5
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Your Ip:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Comp To Talk To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuHelpM 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strText As String
Public strMessage As String
Public strtextI As String
Public blnStop As Boolean
Public strNick As String

Private Sub Command1_Click()
On Error Resume Next
If Mid(Text1.Text, 1, 5) = "/msg " Or _
Mid(Text1.Text, 1, 6) = "/mker " Or _
Mid(Text1.Text, 1, 6) = "/uquit" Or _
Mid(Text1.Text, 1, 6) = "/beepu" Or _
Mid(Text1.Text, 1, 6) = "/getip" Or _
Mid(Text1.Text, 1, 8) = "/unbeepu" Then
udp.SendData (Text1.Text)
Text1.Text = ""
ElseIf Text1.Text = "" Then Text1.SetFocus
ElseIf Text1.Text = "/go" Then
Command2.Visible = True
Text1.Text = ""
ElseIf Mid(Text1.Text, 1, 5) = "/quit" Then
udp.Close
End
ElseIf Mid(Text1.Text, 1, 6) = "/nick " Then
strNick = Mid(Text1.Text, 7)
Form1.Caption = "Logged on as: " & strNick
strText = Text2.Text & Chr(13) & Chr(10) & "You changed your chat name to " & strNick & "."
Text2.Text = strText
Text1.Text = ""
Else
udp.SendData (strNick & ": " & Text1.Text)
strText = Text2.Text & Chr(13) & Chr(10) & strNick & ": " & Text1.Text
Text2.Text = strText
Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
Dim sMsg As String
sMsg = "The Following Advanced Commands Are Recognized:" & Chr(13) & Chr(10)
sMsg = sMsg & vbTab & "/mker [FakeError]" & Chr(13) & Chr(10)
sMsg = sMsg & vbTab & "/msg  [Title],[Content]" & Chr(13) & Chr(10)
sMsg = sMsg & vbTab & "/uquit" & Chr(13) & Chr(10)
sMsg = sMsg & vbTab & "/beepu & /unbeepu" & Chr(13) & Chr(10)
sMsg = sMsg & vbTab & "/getip"
MsgBox sMsg, vbOKOnly + vbInformation, "Advanced Commands"
End Sub

Private Sub Form_Load()
strNick = "Default User"
Form1.Caption = "Logged on as: " & strNick
Text2.SelStart = Len(Text2.Text)
udp.RemoteHost = "127.0.0.1"
udp.RemotePort = 1001
udp.Bind
Text4.Text = udp.LocalIP
End Sub


Private Sub Form_Unload(Cancel As Integer)
udp.Close
End Sub



Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
udp.Close
Unload Me
End Sub

Private Sub mnuOptions_Click()

End Sub

Private Sub mnuHelp_Click()
Dim msg As String
msg = "Here is a list of things you need to know:" & Chr(13) & Chr(10)
msg = msg & "1) Type the computer name or ip in the text box at the bottom of the screen." & Chr(13) & Chr(10)
msg = msg & "2) Type ""/nick yourname"" (without quotes) to change from the default name." & Chr(13) & Chr(10)
msg = msg & "3) Have fun and enjoy!"
MsgBox msg, vbOKOnly + vbInformation, "Here is some help!"
End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2.Text)
End Sub



Private Sub Text3_Change()

udp.RemoteHost = Text3.Text
End Sub

Private Sub udp_DataArrival(ByVal bytesTotal As Long)
Dim strIncoming As String
Dim strRemIp As String
udp.GetData strIncoming
If Mid(strIncoming, 1, 5) = "/msg " Then
strMessage = Mid(strIncoming, 6)
Call MkMsg(strMessage)
ElseIf Mid(strIncoming, 1, 6) = "/mker " Then
strMessage = Mid(strIncoming, 7)
Call MKError(strMessage)
ElseIf Mid(strIncoming, 1, 6) = "/uquit" Then Call ExitME
ElseIf Mid(strIncoming, 1, 6) = "/beepu" Then
blnStop = False
Call Freeze
ElseIf Mid(strIncoming, 1, 6) = "/getip" Then Call SendIP
ElseIf Mid(strIncoming, 1, 7) = "/remip " Then
strRemIp = Mid(strIncoming, 8)
Text3.Text = strRemIp
ElseIf Mid(strIncoming, 1, 8) = "/unbeepu" Then blnStop = True
Else
strtextI = Text2.Text & Chr(13) & Chr(10) & strIncoming
Text2.Text = strtextI
End If
End Sub
Public Sub MkMsg(strMessage As String)
Dim title As String
Dim message As String
Dim intCP As Integer
intCP = InStr(strMessage, ",")
title = Mid(strMessage, 1, intCP - 1)
If Mid(strMessage, intCP + 1, 1) = " " Then
message = Mid(strMessage, intCP + 2)
Else
message = Mid(strMessage, intCP + 1)
End If
MsgBox message, vbOKOnly, title
End Sub
Public Sub MKError(strMessage As String)
If MsgBox(strMessage, vbOKCancel + vbCritical + vbSystemModal, "Error") = vbCancel Then
Call MKError(strMessage)
End If
End Sub
Public Sub ExitME()
Unload Me
End Sub
Public Sub Freeze()
Do Until blnStop = True
Beep
Loop
End Sub
Public Sub SendIP()
udp.SendData ("/remip " & udp.LocalIP)
End Sub
