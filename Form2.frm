VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2940
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "My VB Chat Client Optiosn"
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtLocal 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "9000"
         ToolTipText     =   "Your Nick name. Default:  User-X"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "User-X"
         ToolTipText     =   "Your Nick name. Default:  User-X"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nick"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Simple VB Chat Server"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
      Begin VB.TextBox txtHost 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         MaxLength       =   255
         TabIndex        =   3
         Text            =   "127.0.0.1"
         ToolTipText     =   "Host IP.  Default is LocalPC"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Host"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    On Error Resume Next
    frmMain.Caption = "Simple VB Chat Client - " & txtUser.Text
    If frmMain.cmdSend.Enabled = True Then
        'Assume already connected
        Me.Hide
        frmMain.Show
        Exit Sub
    End If
    'Hide so we can use values in text boxes
    
    frmMain.Port.RemoteHost = IIf(txtHost.Text <> "", txtHost.Text, "127.0.0.1")
    'VB Chat server always uses port
    frmMain.Port.LocalPort = IIf(IsNumeric(txtLocal.Text), txtLocal.Text, 9000)
    frmMain.Port.RemotePort = 1000
    'UDP is easier than TCP / IP
    'Make where no other App can use our chat port
    frmMain.Port.Bind
    'Send connection to Simple VB Chat Server
    frmMain.Port.SendData "Connect_" & txtUser.Text
    frmMain.cmdSend.Enabled = True
    frmMain.Show
    'If any errors, take care of them
    If Err <> 0 Then
        MsgBox "Error# " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbExclamation, "Error"
        MsgBox "Aborting Program", vbCritical, "Error"
        Unload frmMain
        End
    End If
    Me.Hide
End Sub

Private Sub Command2_Click()
    If frmMain.cmdSend.Enabled = False Then
        'Assume we didn't want to connect
        Unload Me
        End
        Exit Sub
    End If
    Me.Hide
    frmMain.Show

End Sub

Private Sub Form_Load()
    'Choose random port
    Randomize
    txtLocal.Text = 9000 + Int(Rnd * 99)
End Sub

