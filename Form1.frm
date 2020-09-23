VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple VB Chat Client"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Chat 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2835
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0442
      ToolTipText     =   "Incoming Chat Data"
      Top             =   0
      Width           =   4335
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2790
      Left            =   4440
      TabIndex        =   5
      Top             =   0
      Width           =   1650
   End
   Begin VB.CheckBox chkScroll 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto Scroll"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "&Config"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3240
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Port 
      Left            =   2400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   9000
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Message 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   6075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim WasBooted As Boolean
 Dim LastData As String

Private Sub Chat_Change()
  
    
    'Should we auto-scrol
    If chkScroll.Value = 1 Then
        'Auto Scroll
        Chat.SelStart = Len(Chat.Text)
    End If
End Sub

Private Sub cmdConfig_Click()
    'Show config form
    Me.Hide
    frmConfig.Show
End Sub

Private Sub cmdQuit_Click()
    
    If WasBooted = False Then
        Port.SendData "Bye_" & frmConfig.txtUser.Text
    End If
    
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSend_Click()
    'Check for codes
    If Mid(Message.Text, 1, 5) = "Drop_" Then
        'Send coded chat data
        Port.SendData Message.Text
    ElseIf Mid(Message.Text, 1, 6) = "Color_" Then
        'Send random color command
        Port.SendData Message.Text
    Else
        'Send normal chat data
        Port.SendData frmConfig.txtUser.Text & ": " & Message.Text
    End If
    'Clear message
    Message.Text = ""
End Sub

Private Sub Command2_Click()


End Sub

Private Sub Command3_Click()
    On Error Resume Next
    'Send close command to chat server
    Port.SendData "Bye_" & frmConfig.txtUser.Text
    Unload Me
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Chat.Text = "Welcome to VB Chat! -" & Now & vbNewLine & Space(5) & "Your Nick: " & frmConfig.txtUser.Text & vbNewLine & Space(5) & "Your IP: " & Port.LocalIP & vbNewLine & vbNewLine
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If WasBooted = False Then
        'We need to send a goodbye string
        Port.SendData "Bye_" & frmConfig.txtUser.Text & vbNewLine
    End If
    'Unload frmConfig
    Unload frmConfig
    
    'Don't send leaving flag to server if they
    'have been booted by admin/ other user
End Sub

Private Sub lstUsers_DblClick()
If lstUsers.Text <> "" Then Message.Text = Message.Text & lstUsers.Text
End Sub

Private Sub Message_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      'Send chat data
      Call cmdSend_Click
   End If
End Sub

Private Sub Port_Close()
    Chat.Text = Chat.Text & vbNewLine & "Connection to host lost!"
End Sub

Private Sub Port_DataArrival(ByVal bytesTotal As Long)
  On Error Resume Next
  Dim strTemp As String
  Port.GetData strTemp, vbString
    
    'Make sure we don't build up to much text
    If Len(Chat.Text) > 63000 Then
            Chat.Text = ">>>  Cleared Chat buffer (64k):" & vbNewLine & Space(5) & "Your Nick: " & frmConfig.txtUser.Text & vbNewLine & Space(5) & "Your IP: " & Port.LocalIP  ' & vbNewLine
    End If
    'Check for codes
    If strTemp = "Drop_" & frmConfig.txtUser.Text & vbNewLine Then
        Port.SendData frmConfig.txtUser.Text & " has left the chat"
        MsgBox "Disconnected from server!"
        Unload frmConfig
        Unload Me
        End
        Exit Sub
    ElseIf strTemp = "Close_" Then
        'Tell everyone you are gone
        Port.SendData frmConfig.txtUser.Text & " has left the chat"
        'Say good bye to server (so slot opens)
        Port.SendData "Bye_" & frmConfig.txtUser.Text
        'Disable buttons because of booted
        cmdSend.Enabled = False
        cmdConfig.Enabled = False
        Message.Enabled = False
        Message.Text = "Bye_"
        'A variable for "Quit" button
        WasBooted = True
        
        'You were forcefully disconnect from the
        'VB Chat Server.  Either the server shut
        'down, or you were booted.
        Chat.Text = Chat.Text & vbNewLine & "You were disconnected from server!"
        Exit Sub
    ElseIf strTemp = "Color_" & frmConfig.txtUser.Text & vbNewLine Then
        'Someone doesn't like the color of your chat
        Randomize
        DoEvents
        Chat.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Chat.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        Message.ForeColor = Chat.ForeColor
        Message.BackColor = Chat.BackColor
        lstUsers.ForeColor = Chat.ForeColor
        lstUsers.BackColor = Chat.BackColor
        Exit Sub
    End If
    'Incoming User List
   If Mid(strTemp, 1, 5) = "User_" Then
      'Check to see if we already have
      'updated list of online users
      If LastData = strTemp Then Exit Sub
      'Catch list of users
      LastData = strTemp
      strTemp = Right(Mid(strTemp, InStr(1, strTemp, "_") + 1), Len(strTemp))
      'Add list of users
      lstUsers.Clear
      Do
        DoEvents
         'Exit loop if no more users
         'Not sure what will return.
         'Will try to catch what I expect
         If Trim(strTemp) = "User_" Or Trim(strTemp) = "" Or Trim(strTemp) = ";" Then Exit Sub
         'Add users to list
         lstUsers.AddItem Mid(strTemp, 1, InStr(1, strTemp, ";") - 1)
         'Strip user in front of first semi-colon
         strTemp = Right(Mid(strTemp, InStr(1, strTemp, ";") + 1), Len(strTemp) - 1)
      Loop
    End If
    
    'If any extra commands were sent to us, lets not post them
    If InStr(1, strTemp, "Color_") Or InStr(1, strTemp, "Drop_") Then Exit Sub
    
    Chat.Text = Chat.Text & strTemp
End Sub


