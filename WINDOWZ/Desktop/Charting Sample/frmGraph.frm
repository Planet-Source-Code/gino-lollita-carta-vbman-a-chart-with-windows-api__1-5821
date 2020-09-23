VERSION 5.00
Begin VB.Form frmGraph 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample by never@penteres.it w3.penteres.it/~never"
   ClientHeight    =   6270
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   7860
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7860
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Etichetta per valori"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HnFont As Long, HoFont As Long
Dim R As Long
Dim Etichetta As Integer


Public Sub Plot()
'Imposto la scala
Cls
ScaleLeft = -200
ScaleTop = 1400
ScaleWidth = 1400
ScaleHeight = -1700

'Disegno l'asse X
Line (0, 0)-(1000, 0)

'Disegno l'asse Y
Line (0, 0)-(0, 1000)


'Disegno le tacchette e i valori per le scale
For I = 100 To 1000 Step 100

    Line (I, -5)-(I, 15)
    CurrentX = CurrentX - 30
    CurrentY = CurrentY - 20
    Print I
    
    Line (-5, I)-(6, I)
    CurrentX = CurrentX - 100
    Print I
    
Next

'Disegno le linee del grafico
DrawStyle = 3
For I = 1 To TNums - 1
    X1 = NArray(I, 0): Y1 = NArray(I, 1)
    X2 = NArray(I + 1, 0): Y2 = NArray(I + 1, 1)
    Line (X1, Y1)-(X2, Y2), QBColor(1)
Next

'Disegno le crocette, i cerchi e le label con i valori
DrawStyle = 0
For I = 1 To TNums

    X1 = NArray(I, 0)
    Y1 = NArray(I, 1)

    'crocette
    Line (X1 - 18, Y1)-(X1 + 22, Y1), QBColor(5)
    Line (X1, Y1 - 30)-(X1, Y1 + 30), QBColor(5)
    Circle (X1, Y1), 5
    Me.ForeColor = vbRed
    Load Label1(I)
    Label1(I).Caption = Trim(Str$(X1)) & "," & Trim(Str$(Y1))
    Label1(I).Left = CurrentX + 20
    Label1(I).Top = CurrentY + 20
    Label1(I).Visible = True

Next

Me.ForeColor = vbBlack

'Stampo il nome del grafico
HnFont = CreateFont(36, 0, 0, 0, FW_BOLD, False, False, False, OEM_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, defauly_quality, 34, "Arial")
HoFont = SelectObject(hdc, HnFont)
sstr$ = frmSet.Text3
R = TextOut(frmGraph.hdc, 320 - (Len(sstr$) * 15), 15, sstr$, Len(sstr$))

'Stampo il testo dell'asse Y
HnFont = CreateFont(16, 0, 0, 0, FW_NORMAL, False, False, False, OEM_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, defauly_quality, 34, "Arial")
HoFont = SelectObject(hdc, HnFont)
sstr$ = frmSet.Text4
R = TextOut(frmGraph.hdc, 200 - (Len(sstr$) * 4), 370, sstr$, Len(sstr$))

'Stampo il testo dell'asse X
HnFont = CreateFont(16, 0, 900, 900, FW_NORMAL, False, False, False, OEM_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, defauly_quality, 34, "Arial")
HoFont = SelectObject(hdc, HnFont)
sstr$ = frmSet.Text5
R = TextOut(frmGraph.hdc, 20, 250 + (Len(sstr$) * 4), sstr$, Len(sstr$))

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1(Etichetta).FontBold = False
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1(Index).FontBold = True
    Etichetta = Index
End Sub
