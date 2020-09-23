VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple VB Chat Server"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5685
   Begin VB.TextBox txtIdle 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "5"
      Top             =   2400
      Width           =   375
   End
   Begin VB.CheckBox chkIdle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Idle Limit"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox chkGuard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Guard Users"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Block user commands"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   5000
      Left            =   2280
      Top             =   1080
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   2010
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Connected Users"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Start"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Incoming Chat data"
      Top             =   0
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock ChatServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "minutes"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim User(1 To 16) As New MyClients

Private Sub ChatServer_DataArrival(ByVal bytesTotal As Long)
    Dim X As Byte
    Dim msg

    'Retrieve sent data
    ChatServer.GetData msg, vbString
    
    'Update Idle Limit for user submitting data
    For X = 1 To 15
        If User(X).MyIP = ChatServer.RemoteHostIP And User(X).MyPort = ChatServer.RemotePort Then
            'Update user's idle limit
            User(X).Idle = Time
            Exit For
        End If
    Next X
    
    'Check for any user close command
    If Mid(msg, 1, 4) = "Bye_" Then
        For X = 1 To 15
            'If they are the user on this port and have matching name
            If User(X).MyPort = ChatServer.RemotePort And User(X).MyName = Mid(msg, 5) Then
                'They want to disconnect
                User(X).MyIP = ""
                User(X).MyPort = 0
                User(X).IsUsed = False
                'Let other users no he /she is gone
                msg = User(X).MyName & " has left the chat" & vbNewLine
                User(X).MyName = ""
                GoTo Broadcast
            End If
        Next X
    End If
    
    'Check  for new connections
    If Mid(msg, 1, 8) = "Connect_" Then
        'Try to find avaaiable slot
        For X = 1 To 16
            If X = 16 Then Exit For
            If User(X).IsUsed = False Then
                'Found a slot
                User(X).IsUsed = True
                User(X).MyIP = ChatServer.RemoteHostIP
                User(X).MyPort = ChatServer.RemotePort
                User(X).Idle = Time
                User(X).MyName = Mid(msg, 9)
                msg = "Connected " & User(X).MyName & " as User#" & X & vbNewLine
                GoTo Broadcast
            End If
        Next X
        
        'User 6 is phantom.
        'Is only used to let user know they
        'could not connect.
        User(16).IsUsed = True
        User(16).MyIP = ChatServer.RemoteHost
        User(16).MyPort = ChatServer.RemotePort
        ChatServer.SendData "Could not allow " & _
              "you to connect.  Too many users." & vbNewLine
        'Close connection
        ChatServer.SendData "Close_"
        User(16).IsUsed = False
        User(16).MyIP = ""
        User(16).MyPort = 0
        Exit Sub
    End If
    
    'Check for booting code
    If Mid(msg, 1, 5) = "Drop_" Then
            If chkGuard.Value = 1 Then
                'This should kill user boot code
                msg = "VB Chat Server Intercepted '" & msg & "' command!" & vbNewLine
                GoTo Broadcast
            End If
                
            'Find user to drop
            For X = 1 To 15
            'If they are the user on this port and have matching name
            If User(X).MyName = Mid(msg, 5) Then
                ChatServer.RemotePort = User(X).MyPort
                ChatServer.SendData "Close_"
                'Disconnect them
                User(X).MyIP = ""
                User(X).MyPort = 0
                User(X).IsUsed = False
                'Let other users know he /she is gone
                msg = User(X).MyName & " has left the chat" & vbNewLine
                User(X).MyName = ""
                GoTo Broadcast
            End If
        Next X
    End If
    
    If Mid(msg, 1, 6) = "Color_" Then
            If chkGuard.Value = 1 Then
                'This should kill the color code
                msg = "VB Chat Server Intercepted '" & msg & "' & command!" & vbNewLine
                GoTo Broadcast
            End If
    End If
        
    'Broad Cast messages to all users
Broadcast:
        'Don't add blank stuff to chatsever log
    If Trim(msg) <> "" Then
        txtData.Text = txtData.Text & msg
    End If
    X = 1
    'Send message to known connected users only
    Do
        'Set to a valid chat client
        If User(X).IsUsed = True Then
            'Some1 connected as this user
            ChatServer.RemoteHost = User(X).MyIP
            ChatServer.RemotePort = User(X).MyPort
            'Send the data
            ChatServer.SendData msg
        End If
        X = X + 1
    Loop Until X > 15
End Sub

Private Sub chkIdle_Click()
'Toggle Enabling of Idle text box
txtIdle.Enabled = Not txtIdle.Enabled
End Sub

Private Sub cmdListen_Click()
    'On port 1000
    On Error Resume Next
    ChatServer.Bind
    cmdListen.Enabled = False
    
    If Err <> 0 Then
        'Probably already attemped bind method
        'Address in use or address family not
        'supported error is my guess
        MsgBox "Error# " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbExclamation, "Error"
        MsgBox "Error. Aborting Program!", vbCritical, "Error"
        Unload Me
        End
        Exit Sub
    End If
    
    txtData.Text = txtData.Text & vbNewLine & vbNewLine & "VB ChatServer@" & ChatServer.LocalIP & ":" & ChatServer.LocalPort
'Server works by broadcasting message
'May not be very efficient, but gets the job
'done for now
End Sub

Private Sub Form_Load()
'Resize chat window at runtime
'txtData.Move 0, 0, ScaleWidth, ScaleHeight - cmdListen.Height
txtData.Text = "Copyright 1998-99. Kenneth Gilbert Jr." & vbNewLine & "All rights resereved." & vbNewLine & "VB Chat Server: Up to 15 users may chat at one time on this server!." & vbNewLine
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Disconnect all users conencted to VB Chat Server
Dim X As Byte
For X = 1 To 5
    If User(X).IsUsed = True Then
        'Boot the user
        ChatServer.RemotePort = User(X).MyPort
        ChatServer.SendData "Close_"
    End If
    'Kill MyClients object
    Set User(X) = Nothing
Next X
'Done.  Good bye!
End Sub

Private Sub lstUsers_DblClick()
    Dim X As Byte
    'Boot selected user
      'Find user to drop
        For X = 1 To 15
            'If they are the user on this port and have matching name
            If User(X).MyName = lstUsers.Text And User(X).IsUsed = True Then
                ChatServer.RemotePort = User(X).MyPort
                ChatServer.SendData "Close_"
                'They want to disconnect
                User(X).MyIP = ""
                User(X).MyPort = 0
                User(X).MyName = ""
                User(X).Idle = ""
                User(X).IsUsed = False
                Exit Sub
            End If
        Next X
        'If we get here, we did not find the user.
        txtData.Text = txtData.Text & vbNewLine & vbNewLine & "User was not found in MyClients!" & vbNewLine & vbNewLine
End Sub

Private Sub Text1_Change()
    
End Sub

Private Sub tmrUpdate_Timer()
    Dim X As Byte
    Dim strUsers As String
    strUsers = "User_"
    'Update user listing ever 5 seconds
    'Update user listing
    lstUsers.Clear
    For X = 1 To 15
        If User(X).IsUsed = True Then
            'User connected here
            lstUsers.AddItem User(X).MyName
            strUsers = strUsers & User(X).MyName & ";"
        End If
    Next
        
   'Send users to all chat clients
   For X = 1 To 15
      If User(X).IsUsed = True Then
         'User connected.
         'Send a user list
         ChatServer.RemoteHost = User(X).MyIP
         ChatServer.RemotePort = User(X).MyPort
         ChatServer.SendData strUsers
      End If
   Next
   
   'Check for Idlers
   If chkIdle.Value = 1 Then
        For X = 1 To 15
           If User(X).IsUsed = True Then
             If DateDiff("s", User(X).Idle, Time) >= Int(txtIdle.Text * 60) Then
                 'User connected.
                 'Send a user list
                 ChatServer.RemoteHost = User(X).MyIP
                 ChatServer.RemotePort = User(X).MyPort
                 ChatServer.SendData "Close_"
             End If
          End If
        Next
 End If
End Sub

Private Sub txtData_Change()
    On Error Resume Next
    'Move to bottom of text box
    If Len(txtData.Text) > 63000 Then
        'Getting to much information
        '64k limit
        txtData.Text = ""
    End If
    txtData.SelStart = Len(txtData.Text)
End Sub


Private Sub txtIdle_Change()
'Make sure they have entered numbers
If IsNumeric(txtIdle) = False Then txtIdle = 5

'If they enter < 1 then set off
If txtIdle.Text < 1 Then
    chkIdle.Value = 0
    txtIdle.Text = 0
    txtIdle.Enabled = False
End If
End Sub

Private Sub txtIdle_GotFocus()
'Select all text from start to end
txtIdle.SelStart = 0
txtIdle.SelLength = Len(txtIdle.Text)
'Done
End Sub
