VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VBChatJr (UDP chat) Class
'Client handling Class by Kenneth Gilibert Jr. (Sage_Z)

Private strMyName As String
Private intMyPort As Integer
Private strMyIP As String
Private bolInUse As Boolean
Private strIdle As String

Public Property Get MyName() As String
    MyName = strMyName
End Property

Public Property Let MyName(strValue As String)
    strMyName = strValue
End Property

Public Property Get MyPort() As Integer
    MyPort = intMyPort
End Property

Public Property Let MyPort(intValue As Integer)
    intMyPort = intValue
End Property

Public Property Get MyIP() As String
    MyIP = strMyIP
End Property

Public Property Let MyIP(strValue As String)
    strMyIP = strValue
End Property

Public Property Get IsUsed() As Boolean
    IsUsed = bolInUse
End Property

Public Property Let IsUsed(bolValue As Boolean)
    bolInUse = bolValue
End Property

Public Property Get Idle() As String
    Idle = strIdle
End Property

Public Property Let Idle(strValue As String)
    strIdle = strValue
End Property
