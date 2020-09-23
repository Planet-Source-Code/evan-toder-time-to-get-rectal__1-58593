VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "rect drawing demo"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   285
      Left            =   3150
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.CheckBox Check1 
      Caption         =   "with color"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   1530
      Width           =   1905
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2835
      Top             =   585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type Rect
   Left As Long
   Top  As Long
   Right As Long
   Bottom As Long
End Type

Dim R(15)   As Rect
Dim cRect            As cdraw_rect
 
 

Private Sub cmdClose_Click()
 
 Timer1 = False
 Unload Me
 End
 
End Sub

Private Sub Form_Load()
  
  '
  'so we only have to draw the form once (or twice)
  AutoRedraw = True
  Set cRect = New cdraw_rect
  '
  'set the current pallette to the form
  Set cRect.your_obj_pallete = Form1
  '
  
  'start timer
  Timer1.Interval = 100
  Timer1 = True
  
End Sub

Private Sub Form_Terminate()

  Set cRect = Nothing
  
End Sub
 
 
Private Sub Timer1_Timer()

 On Error Resume Next
 
 Dim wid&, hei&, lcnt&  'longs
 Dim MyValue As Single
 
 wid = Width * 0.8
 hei = Height * 0.7

 
  ' Initialize random-number generator.
 
 For lcnt = 200 To 600 Step 20
   Cls
   MyValue = Int((Rnd * lcnt) + 1)
   cRect.draw_rect MyValue, (MyValue * 0.5), (MyValue + wid), _
                 (MyValue * 0.5) + hei, svtwips, , _
                 Int((7 * Rnd) + 1), Int((1000 * Rnd) + 1)
                 
   If Check1.Value = 1 Then
       cRect.fill_curr_rectRect RGB(Int((255 * Rnd) + 1), Int((255 * Rnd) + 1), Int((255 * Rnd) + 1))
   End If
   DoEvents
 Next lcnt
 
 
End Sub
