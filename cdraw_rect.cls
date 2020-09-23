VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdraw_rect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 
 
Private Type Rect
   Left As Long
   Top  As Long
   Right As Long
   Bottom As Long
End Type
 
Enum enScale
      svtwips = 0
      svpixels = 1
End Enum
 
Enum enDrawType
    drCOLOR_LINE = 1
    drSUNKEN = 2
    drSUNKEN_SOFT = 3
    drRAISED = 4
    drRAISED_SOFT = 5
    drETCHED = 6
    drSEAM = 7
End Enum
 
Enum enDrawLocal
    BF_BOTTOM = &H8
    BF_LEFT = &H1
    BF_RIGHT = &H4
    BF_DIAGONAL = &H10
    BF_TOP = &H2
    BF_FLAT = &H4000
    BF_MIDDLE = &H800
    BF_MONO = &H8000
    BF_SOFT = &H1000
    
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    
    BF_TOPRIGHTBOTTOM = (BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_RIGHTBOTTOMLEFT = (BF_LEFT Or BF_RIGHT Or BF_BOTTOM)
    BF_BOTTOMLEFTTOP = (BF_LEFT Or BF_TOP Or BF_BOTTOM)
    BF_LEFTTOPRIGHT = (BF_LEFT Or BF_TOP Or BF_RIGHT)
    
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
End Enum


Public your_obj_pallete      As Object
Dim m_curr_dc                As Long
Dim m_curr_R                 As Rect
 
Sub fill_curr_rectRect(lcolor As Long)
   
  Dim lbrush& 'long
  lbrush& = CreateSolidBrush(lcolor)
  FillRect m_curr_dc, m_curr_R, lbrush
  DeleteObject lbrush
  
End Sub

'Sub type_in_curr_rect(str_text$)
'
'
'End Sub


Sub draw_rect(Left!, Top!, Right!, Bottom!, _
              inputs_scale_val As enScale, _
              Optional brush_color&, _
              Optional BDR As enDrawType, _
              Optional BF As enDrawLocal)
                         
  On Local Error GoTo local_error:
 
  Dim lbrush&
  '
  'if user provide twips values->convert to pixels
  If inputs_scale_val = svtwips Then
1   Left = (Left / Screen.TwipsPerPixelX)
2   Right = (Right / Screen.TwipsPerPixelX)
3   Top = (Top / Screen.TwipsPerPixelY)
4   Bottom = (Bottom / Screen.TwipsPerPixelY)
  End If
  '
  'this statement is actually a test to see if
  'the control passed has a .hdc property
  'If it doesnt then an error will be returned
  'and code will branch to "local_error:"
5 m_curr_dc = your_obj_pallete.hdc
  '
  'create the rect area
6 SetRect m_curr_R, Left, Top, Right, Bottom
  
  If BDR = drCOLOR_LINE Then
7     lbrush& = CreateSolidBrush(brush_color)
8     FrameRect m_curr_dc, m_curr_R, lbrush
  Else
      DrawEdge m_curr_dc, m_curr_R, func_SelectStyle(BDR), BF
  End If
  '
  'clean up
  DeleteObject lbrush
  
local_error:
   With Err
      If .Number <> 0 Then
         If Erl() = 5 Then
           If Err.Number = 438 Then
              m_curr_dc = GetWindowDC(your_obj_pallete.hwnd)
              Resume Next
           Else
              Debug.Print "cdraw_rect.draw_rect." & _
                          .Number & "." & .Description
              Resume Next
           End If
         End If
      End If
   End With
End Sub
 

Private Function func_SelectStyle(item_selected As enDrawType) As Long
  
  Const BDR_INNER As Long = &HC
  Const BDR_OUTER As Long = &H3
  Const BDR_RAISED As Long = &H5
  Const BDR_RAISEDINNER As Long = &H4
  Const BDR_RAISEDOUTER As Long = &H1
  Const BDR_SUNKEN As Long = &HA
  Const BDR_SUNKENINNER As Long = &H8
  Const BDR_SUNKENOUTER As Long = &H2
   
  If item_selected = drETCHED Then
     func_SelectStyle = (BDR_RAISEDINNER Or BDR_SUNKENOUTER)
  ElseIf item_selected = drRAISED Then
     func_SelectStyle = BDR_RAISED
  ElseIf item_selected = drRAISED_SOFT Then
     func_SelectStyle = BDR_RAISEDINNER
  ElseIf item_selected = drSEAM Then
     func_SelectStyle = (BDR_SUNKENINNER Or BDR_RAISEDOUTER)
  ElseIf item_selected = drSUNKEN Then
     func_SelectStyle = BDR_SUNKEN
  ElseIf item_selected = drSUNKEN_SOFT Then
     func_SelectStyle = BDR_SUNKENINNER
  End If
  
End Function
 
 
Private Sub Class_Terminate()
 
 Set your_obj_pallete = Nothing
 
End Sub
