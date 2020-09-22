VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox pix 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3135
      Width           =   690
   End
   Begin VB.HScrollBar v_scale 
      Height          =   420
      Left            =   120
      Max             =   200
      Min             =   2
      TabIndex        =   19
      Top             =   3135
      Value           =   2
      Width           =   1050
   End
   Begin VB.TextBox cmplx 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2280
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
      Begin VB.OptionButton op_red 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton op_yellow 
         BackColor       =   &H0080FFFF&
         Caption         =   "Yellow"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton op_blue 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton op_green 
         BackColor       =   &H0000FF00&
         Caption         =   "Green"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox branch 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   315
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1455
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Render"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   5415
      Left            =   1320
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"tree.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1440
         TabIndex        =   22
         Top             =   120
         Width           =   4095
      End
      Begin VB.Line Line3 
         BorderWidth     =   10
         X1              =   32
         X2              =   16
         Y1              =   40
         Y2              =   24
      End
      Begin VB.Line Line2 
         BorderWidth     =   10
         X1              =   32
         X2              =   16
         Y1              =   8
         Y2              =   24
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   16
         X2              =   88
         Y1              =   24
         Y2              =   24
      End
   End
   Begin VB.HScrollBar v_branch 
      Height          =   420
      Left            =   135
      Max             =   200
      Min             =   2
      TabIndex        =   16
      Top             =   1455
      Value           =   2
      Width           =   1050
   End
   Begin VB.HScrollBar v_cmplx 
      Height          =   420
      Left            =   120
      Max             =   200
      Min             =   2
      TabIndex        =   17
      Top             =   2280
      Value           =   2
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
      Begin VB.OptionButton less 
         BackColor       =   &H00000000&
         Caption         =   "Less"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton more 
         BackColor       =   &H00000000&
         Caption         =   "More"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox lenth_active 
         BackColor       =   &H00000000&
         Caption         =   "Active"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Length Randomization"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Scale (pixels)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Complexity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Branches"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
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

Dim cg
Dim rr
Dim aa
Dim cc
Dim prog As Long
Dim progg As Long

Private Sub angle_active_Click()
If angle_active.Value = 1 Then aa = 1 Else aa = 0
End Sub

Public Sub Command1_Click()
'clearing the arrow and the box
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False

Label5.Visible = False

Picture1.Cls
Dim aa As Integer
Dim bb As Integer
Dim cc As Double

aa = v_branch.Value
bb = v_cmplx.Value
cc = v_scale.Value

prog = aa ^ bb

progg = 0
cg = bb
Picture1.DrawWidth = 10
Picture1.Line (Picture1.ScaleWidth / 2, Picture1.ScaleHeight)-(Picture1.ScaleWidth / 2, (Picture1.ScaleHeight - 100))
Call branches(aa, bb, cc, (Picture1.ScaleWidth) / 2, (Picture1.ScaleHeight - 100))
End Sub

'*************************************************************************************
'
' The recursive method
'
'*************************************************************************************

Sub branches(trunks As Integer, level As Integer, sc As Double, X As Double, Y As Double)

' trunks = number of braches to draw
' level = at what level of recursion in the function on
' sc = scale of pixcels per branch
' X, Y = Starting co-ordinates of any particular recursive call for the brances

'temporary variables to draw each branch from X, Y to xx, yy for each branc
Dim xx As Double
Dim yy As Double

Randomize Timer

'The call needs to finish somewhere. The value of level is set initally by the user
'through the complexity input. Initial valuse are provided by the render button
If level = 0 Then GoTo finishit

'figuring out the angle interval for each branch
theta = 3.14 / (trunks)

'if the user wants to randomize the lengths

'less
If rr = 0 Then lenth = level * sc

'more
If rr = 2 Then lenth = Int(Rnd * (level * sc * 2))
 
'Draw the branches
For i = 0 To (trunks - 1)
    
    addd = theta * i
    Randomize Timer
    If rr = 1 Then lenth = Int(Rnd * (level * sc * 2))
    
    xx = lenth * Cos((Rnd * theta) + addd)
    yy = lenth * Sin((Rnd * theta) + addd)
    Picture1.DrawWidth = level
    
    'draw the branch with the selected color
    If cc = 1 Then Picture1.Line (X, Y)-((X + xx), (Y - yy)), RGB(level * 10, (256 - (level * 256 / cg)), 0)
    If cc = 2 Then Picture1.Line (X, Y)-((X + xx), (Y - yy)), RGB(0, level * 10, (256 - (level * 256 / cg)))
    If cc = 3 Then Picture1.Line (X, Y)-((X + xx), (Y - yy)), RGB((256 - (level * 256 / cg)), (256 - (level * 256 / cg)), 0)
    If cc = 4 Then Picture1.Line (X, Y)-((X + xx), (Y - yy)), RGB((256 - (level * 256 / cg)), level * 10, 0)
    
    'IMPORTANT - The Recursive call, with level decremented and new Valuse of X and Y
    Call branches(trunks, (level - 1), sc, (X + xx), (Y - yy))
    
    'Remove DoEvents to make the process faster if you want. Other wise with large values
    'for complexity and the number of branches, it can take a very long time.
    DoEvents
    
    'ignore, used only for testing purposes
    progg = progg + 1
    Form1.Caption = Str$(progg)
Next i

finishit:
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize Timer
siz = Int(Rnd * 5) + 10
fon = Int(Rnd * 2) + 1
If fon = 1 Then Command1.Font.Name = "Arial"
If fon = 2 Then Command1.Font.Name = "Impact"
If fon = 3 Then Command1.Font.Name = "Courier New"
Command1.Font.Size = siz
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize Timer
siz = Int(Rnd * 5) + 18
fon = Int(Rnd * 2) + 1
If fon = 1 Then Command2.Font.Name = "Arial"
If fon = 2 Then Command2.Font.Name = "Impact"
If fon = 3 Then Command2.Font.Name = "Courier New"



Command2.Font.Size = siz
End Sub

Private Sub Form_Load()
Form1.Left = 0
Form1.Top = 0
Form1.Width = Screen.Width
Form1.Height = Screen.Height

v_branch.Value = 5
v_cmplx.Value = 7
v_scale.Value = 12


lenth_active.Value = 0
less.Value = True
rr = 0
op_green.Value = True
cc = 1

Form2.Show 1
End Sub

Private Sub Form_Resize()
Picture1.Width = Form1.ScaleWidth - Picture1.Left - 130
Picture1.Height = Form1.ScaleHeight - Picture1.Top - 130


End Sub



Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize Timer
siz = Int(Rnd * 2) + 7
fon = Int(Rnd * 2) + 1
If fon = 1 Then Label1.Font.Name = "Arial"
If fon = 2 Then Label1.Font.Name = "Impact"
If fon = 3 Then Label1.Font.Name = "Courier New"
Label1.Font.Size = siz
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize Timer
siz = Int(Rnd * 2) + 7
fon = Int(Rnd * 2) + 1
If fon = 1 Then Label2.Font.Name = "Arial"
If fon = 2 Then Label2.Font.Name = "Impact"
If fon = 3 Then Label2.Font.Name = "Courier New"
Label2.Font.Size = siz
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize Timer
siz = Int(Rnd * 2) + 7
fon = Int(Rnd * 2) + 1
If fon = 1 Then Label3.Font.Name = "Arial"
If fon = 2 Then Label3.Font.Name = "Impact"
If fon = 3 Then Label3.Font.Name = "Courier New"
Label3.Font.Size = siz
End Sub


Private Sub lenth_active_Click()
If lenth_active.Value = 1 Then
    If more.Value = True Then rr = 1
    If less.Value = True Then rr = 2
    'MsgBox Str$(rr)
Else
    rr = 0
End If
End Sub

Private Sub op_blue_Click()
cc = 2
End Sub

Private Sub op_green_Click()
cc = 1
End Sub

Private Sub op_red_Click()
cc = 4
End Sub

Private Sub op_yellow_Click()
cc = 3
End Sub

Private Sub v_branch_Change()
branch.Text = Str$(v_branch.Value)
End Sub

Private Sub v_cmplx_Change()
cmplx.Text = Str$(v_cmplx.Value)
End Sub

Private Sub v_scale_Change()
pix.Text = Str$(v_scale.Value)
End Sub
