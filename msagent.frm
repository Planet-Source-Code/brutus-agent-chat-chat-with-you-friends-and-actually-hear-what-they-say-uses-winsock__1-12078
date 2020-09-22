VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   2925
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "msagent.frx":0000
      Top             =   120
      Width           =   4095
   End
   Begin AgentObjectsCtl.Agent Agent2 
      Left            =   2520
      Top             =   3240
      _cx             =   847
      _cy             =   847
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1320
      Top             =   840
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
 
Const DATAPATH = "\windows\msagent\chars\genie.acs"
 
 
Private Sub Command1_Click()
Genie.Show
Genie.speak Text1.Text
Genie.MoveTo 230, 600, 23
 Genie.Activate 1
 

End Sub

Private Sub Form_Load()
    Agent1.Characters.Load "Genie", DATAPATH
    Set Genie = Agent1.Characters("Genie")
     
    End Sub

 

