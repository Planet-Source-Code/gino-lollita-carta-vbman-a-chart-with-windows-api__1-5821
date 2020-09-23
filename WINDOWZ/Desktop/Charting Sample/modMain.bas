Attribute VB_Name = "modMain"
DefInt A-Z
Global I, J, TNums, SPos, FPos, NLen
Global NArray(50, 50) As Single, ScaledNArray(50, 50) As Single
Global PtXMax As Single, PtYMax As Single
Global X1 As Single, Y1 As Single
Global X2 As Single, Y2 As Single


Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700

Public Const OEM_CHARSET = 255
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const DEFAULT_QUALITY = 0

Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

