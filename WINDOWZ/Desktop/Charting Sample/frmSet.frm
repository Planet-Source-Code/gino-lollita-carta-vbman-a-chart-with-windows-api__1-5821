VERSION 5.00
Begin VB.Form frmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esempio di grafico con le API"
   ClientHeight    =   2760
   ClientLeft      =   2745
   ClientTop       =   2625
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6435
   Begin VB.CommandButton Command1 
      Caption         =   "Esci"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disegna"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Text            =   "Scala valori Asse X"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "Scala valori Asse Y"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Text            =   "Grafico con le API"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "70,2,36,81,52,18,61"
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "4,11,26,42,60,74,97"
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valori asse Y"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valori asse X"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Titolo asse X"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Titolo asse Y"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Testo del grafico"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case Is = 0
        ProcessData
        frmGraph.Show
        frmGraph.Plot
    Case Is = 1
        Unload Me
End Select
End Sub

Private Sub ProcessData()
TNums = 1
For I = 1 To Len(Text1)
    XCh$ = Mid$(Text1, I, 1)
    yCh$ = Mid$(Text2, I, 1)
    If XCh$ = "," Then TNums = TNums + 1
Next

xnewstr$ = Text1 + ","
SPos = 0
J = 0

For I = 1 To TNums
    FPos = InStr(SPos + 1, xnewstr$, Chr$(44))
    NLen = (FPos - SPos) - 1
    NArray(I, J) = (Val(Mid$(xnewstr$, SPos + 1, NLen))) * 10
    SPos = FPos
Next
    
ynewstr$ = Text2 + ","
SPos = 0
J = 1

For I = 1 To TNums
    FPos = InStr(SPos + 1, ynewstr$, Chr$(44))
    NLen = (FPos - SPos) - 1
    NArray(I, J) = (Val(Mid$(ynewstr$, SPos + 1, NLen))) * 10
    SPos = FPos
Next


End Sub

