VERSION 5.00
Begin VB.Form TestBox 
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin Project1.BoxList LstFont 
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1931
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton BtnAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton BtnSel 
      Caption         =   "&Select"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "* Double-click or press Delete  over ListBox"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced BoxList example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "TestBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Advanced ListBox implementation
'Author: Eleusmario Mariano Rabelo
'This is a listbox where the user can define diverse colors and fonts
'at the same time, alignment of columns, and several other resources.
'This component generates three events: Box_Click(), Box_DblClick() and Box_KeyDown()
'You'll be able to add any type of data to listbox,
'in this example, it add the Windows font names.

Dim FSize As Integer, QLines As Integer, ColorBack As Long
Dim ColorCell As Long, ColorText As Long
Dim Orange As Long, Green As Long, Blue As Long
Const LeftPos = 0, RightPos = 1, CenterPos = 2

Private Sub Form_Load()
FSize = 10   'font Size
QLines = 15  'amount of lines
ColorBack = RGB(200, 200, 200) 'box backcolor
ColorCell = RGB(255, 255, 200) 'default cells color
ColorText = 0                  'default text color
Orange = RGB(255, 230, 180)    'orange color
Green = RGB(200, 255, 200)     'green color
Blue = RGB(150, 150, 255)      'blue color

'config box
LstFont.Config Me, "Tahoma", FSize, QLines, ColorBack, ColorCell, ColorText

'config columns
LstFont.Title "Sequence", 10: LstFont.Title "Font Name", 30
LstFont.Title "Name Size", 10, CenterPos
LstFont.Title "Float demo", 10, RightPos

'activate box with heading
LstFont.Activate True

'add data to listbox
Cont = 0
Do While Cont < Screen.FontCount - 1
  LstFont.Add Str(Cont), Orange, 0, "Verdana", 10, True
  LstFont.Add " " & Screen.Fonts(Cont)
  LstFont.Add Str(Len(Screen.Fonts(Cont))), Green, 0, "Arial", 10, True
  LstFont.Add Format(Len(Screen.Fonts(Cont)), "#0.00"), Blue, vbWhite, "MS Sans Serif", 10, True
  LstFont.BoxNew  'new line
  Cont = Cont + 1
Loop
End Sub

Private Sub BtnAdd_Click()
LstFont.Add " 999", Orange, 0, "Verdana", 10, True
LstFont.Add " Test"
LstFont.Add "123", Green, 0, "Arial", 10, True
LstFont.Add Format(12345.67, "###,##0.00"), Blue, vbWhite, "MS Sans Serif", 10, True
LstFont.BoxNew  'new line
LstFont.Selected LstFont.ListCount - 1
End Sub

Private Sub BtnSel_Click()
LstFont.Selected 20
MsgBox LstFont.ListIndex, vbInformation
End Sub

Private Sub BtnClear_Click()
LstFont.Clear
End Sub

Public Sub Box_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
  MsgBox LstFont.Arg(0), vbInformation
End If
End Sub

Public Sub Box_Click()
End Sub

Public Sub Box_DblClick()
MsgBox LstFont.Arg(0) & " - " & LstFont.Arg(1) & " - " & LstFont.Arg(2) & " - " & LstFont.Arg(3), vbInformation
End Sub

