VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SID"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Normal"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      MaxLength       =   13
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Len(Text1.Text) <> 13 Then
MsgBox ("Length must be 13")
Exit Sub
End If
StationID = GetSerialNumber(Text1.Text)
MsgBox ("Your SID is now " & DecompressSID(StationID))
End Sub

Private Sub Command2_Click()
StationID = GetSerialNumber
MsgBox ("Normal SID Restored.")
End Sub

Private Sub Form_Load()
Text1.Text = StationID
End Sub

