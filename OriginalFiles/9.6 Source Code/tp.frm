VERSION 5.00
Begin VB.Form tp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Team Power"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2580
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   2580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "tp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
TeamRank = "0"

End Sub

Private Sub Form_Load()
Text1.Text = TeamRank
End Sub
