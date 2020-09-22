VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   2040
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Danish Mujeeb       Feb 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "trees©"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================
'
'  Danish Mujeeb
'  d_mujeeb@hotmail.com
'
'  any comments? please send
'
'=========================================================


Dim siz1
Dim siz2
Dim oldFont

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Form2.Left = (Screen.Width / 2) - (Form2.Width / 2)
Form2.Top = (Screen.Height / 2) - (Form2.Height / 2)

'siz1 = Label1.Font.Size - 3
siz2 = Label2.Font.Size - 3


End Sub


Private Sub Timer1_Timer()
Randomize Timer
stylee = Int(Rnd * 3) + 1
siz = Int(Rnd * 5) + 1

fon = Int(Rnd * 2) + 1

If fon <> oldFont Then
    If fon = 1 Then Label2.Font.Name = "Arial"
    If fon = 2 Then Label2.Font.Name = "Impact"
    If fon = 3 Then Label2.Font.Name = "Courier New"
    oldFont = fon
    Label2.Font.Size = siz2 + siz
End If

End Sub
