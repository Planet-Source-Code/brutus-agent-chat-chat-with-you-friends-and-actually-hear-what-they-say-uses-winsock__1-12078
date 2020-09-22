VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Agent chat"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws 
      Left            =   1800
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Speak"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   " "
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Listen"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   960
      Top             =   480
      _cx             =   847
      _cy             =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Genie As IAgentCtlCharacter
 Dim wsdata As String
Const DATAPATH = "\windows\msagent\chars\genie.acs"
 
 
 Private Sub Command1_Click() ' connect
ws.RemoteHost = Text1.Text
ws.RemotePort = "80"
ws.Connect
Command3.Enabled = False
Command2.Enabled = False
End Sub

Private Sub Command2_Click() ' disconnect
ws.Close
End Sub

Private Sub Command3_Click() ' listen
ws.LocalPort = "21"
ws.Listen
Command1.Enabled = False
Command2.Enabled = False
 End Sub


Private Sub Command4_Click()
wsdata = Text2.Text
ws.SendData wsdata
End Sub

Private Sub ws_Connect()
Genie.Show
Genie.Speak "Welcome to Richard's chat Program. To send text,simply press send"
Label1.Caption = "Connected"
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
ws.Close
ws.Accept requestID
End Sub

Private Sub Form_Load()
Agent1.Characters.Load "Genie", DATAPATH
    Set Genie = Agent1.Characters("Genie")
  Text1.Text = ws.LocalIP
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
ws.GetData wsdata
 

Genie.MoveTo 300, 300, 150
Genie.Speak wsdata
End Sub
