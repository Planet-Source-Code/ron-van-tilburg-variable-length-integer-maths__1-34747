VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TestRVTvbIM - Variable Length Integers"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   13425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vf Tan, ATan"
      Height          =   465
      Index           =   10
      Left            =   12210
      TabIndex        =   17
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfCos,ACos"
      Height          =   465
      Index           =   9
      Left            =   11010
      TabIndex        =   16
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfSin,ASin"
      Height          =   465
      Index           =   8
      Left            =   9810
      TabIndex        =   15
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfExp,Log"
      Height          =   465
      Index           =   7
      Left            =   8610
      TabIndex        =   14
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "Solve Hx=B"
      Height          =   465
      Index           =   6
      Left            =   7410
      TabIndex        =   13
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "Convergents"
      Height          =   465
      Index           =   5
      Left            =   6210
      TabIndex        =   12
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "viKRoot"
      Height          =   465
      Index           =   5
      Left            =   6210
      TabIndex        =   11
      Top             =   9060
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Height          =   465
      Index           =   4
      Left            =   5010
      TabIndex        =   10
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfPow"
      Height          =   465
      Index           =   3
      Left            =   3780
      TabIndex        =   9
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfSqrt"
      Height          =   465
      Index           =   2
      Left            =   2550
      TabIndex        =   8
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vfInt,Frc, Floor,Ceil"
      Height          =   465
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLF 
      Caption         =   "vf + - * /"
      Height          =   465
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   9570
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "vi Rand"
      Height          =   465
      Index           =   4
      Left            =   5010
      TabIndex        =   5
      Top             =   9060
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "Fact,nPr,nCr"
      Height          =   465
      Index           =   3
      Left            =   3780
      TabIndex        =   4
      Top             =   9060
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "vi LCD,LCM"
      Height          =   465
      Index           =   2
      Left            =   2550
      TabIndex        =   3
      Top             =   9060
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "vi Sqrt, Pow"
      Height          =   465
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   9060
      Width           =   1155
   End
   Begin VB.CommandButton cmdVLI 
      Caption         =   "vi + - * /"
      Height          =   465
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   9060
      Width           =   1155
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8940
      ItemData        =   "TestRVTvbIM.frx":0000
      Left            =   60
      List            =   "TestRVTvbIM.frx":0002
      TabIndex        =   0
      Top             =   60
      Width           =   13245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copyright ©2002 R. van Tilburg
'Public rights for educational use are given. Commercial rights retained by author
'===============================================================================================================

Private Sub DispClear()
  List1.Clear
End Sub

Private Sub Disp(fnName As String, Optional p1 As Long, Optional p2 As Long, Optional result As String)
  If fnName = "" Then
    List1.AddItem ""
  Else
    List1.AddItem fnName & "(" & Str$(p1) & "," & Str$(p2) & ") = " & result
  End If
End Sub

Private Function FmtInt(ByVal x As Long, ByVal Width As Long) As String
  FmtInt = Str$(x)
  If Len(FmtInt) < Abs(Width) Then
    If Width < 0 Then
      FmtInt = FmtInt & String$(Abs(Width) - Len(FmtInt), 32)
    Else
      FmtInt = String$(Width - Len(FmtInt), 32) & FmtInt
    End If
  End If
End Function


Private Sub cmdVLI_Click(Index As Integer)
  Select Case Index
    Case 0: VLI_AddSubMulDivMod
    Case 1: VLI_SqrtPow
    Case 2: VLI_LCDLCM
    Case 3: VLI_Fact
    Case 4: VLI_Rand
    Case 5: VLI_KRoot
  End Select
End Sub

Private Sub cmdVLF_Click(Index As Integer)
  Select Case Index
    Case 0:  VLF_AddSubMulDivMod
    Case 1:  VLF_IntFrcCeilFloor
    Case 2:  VLF_Sqrt
    Case 3:  VLF_Pow
    Case 4:
    Case 5:  VLF_Convergent
    Case 6:  Call VLF_Hilbert(7)
    Case 7:  VLF_ExpLog
    Case 8:  VLF_SinASin
    Case 9:  VLF_CosACos
    Case 10: VLF_TanATan
  End Select
End Sub
'========================== VLI Demos ==================================================================

Private Sub VLI_AddSubMulDivMod()
  DispClear
'  viASMDM 0, 0
'  viASMDM 1, 0
  viASMDM 0, 1
'  viASMDM -1, 0
  viASMDM 0, -1
  viASMDM 1, 1
  viASMDM -1, -1
  viASMDM 1, -1
  viASMDM -1, 1
  viASMDM 1, 2
  viASMDM -1, -2
  viASMDM 1, -2
  viASMDM -1, 2
  viASMDM 3, -2
  viASMDM -3, 2
  viASMDM 2, -3
  viASMDM -2, 3
  viASMDM 2, 3
  viASMDM -2, -3
  viASMDM 3, 2
  viASMDM -3, -2
End Sub

Private Sub viASMDM(x As Long, y As Long)
  Dim p As vInt, q As vInt
  
  p = viFromL(x)
  q = viFromL(y)
  List1.AddItem "viAdd(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & viToStr(viAdd(p, q)) _
              & ",   viSub(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & viToStr(viSub(p, q)) _
              & ",   viMul(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & viToStr(viMul(p, q)) _
              & ",   viDiv(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & viToStr(viDiv(p, q)) _
              & ",   viMod(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & viToStr(viMod(p, q))
End Sub

Private Sub VLI_SqrtPow()
  Dim i As Long, j As Long, k As Long
  Dim q As vInt, w As vInt, s As String
  
  DispClear
  For i = 1 To 29
    j = (2 ^ 31 - 1) * Rnd
    q = viFromL(j)
    w = viSqrt(q) 'viMul(q, q))
    List1.AddItem "Given q =" _
                 & FmtInt(j, 11) _
                 & ", viSqrt(q)=" _
                 & viFmt(w, 7) _
                 & ", Excess=" _
                 & viFmt(viSub(q, viMul(w, w)), 7)
  Next
  Disp ""
  List1.AddItem "-33^13= " & viToStr(viPow(viFromL(-33), viFromL(13)))
  List1.AddItem " 31^15= " & viToStr(viPow(viFromL(31), viFromL(15)))
  List1.AddItem " 11^33= " & viToStr(viPow(viFromL(11), viFromL(33)))
  Disp ""
  For i = 0 To 150
    List1.AddItem " 2^" & FmtInt(i, -4) & "=" & viFmt(viTwoToThe(viFromL(i)), 50)
  Next
  For i = 0 To 30
    List1.AddItem "10^" & FmtInt(i, -4) & "=" & viFmt(viTenToThe(viFromL(i)), 50)
  Next
  For i = 0 To 30
    List1.AddItem "13^" & FmtInt(i, -4) & "=" & viFmt(viPow(viFromL(13), viFromL(i)), 50)
  Next
End Sub

Private Sub VLI_KRoot()
  Dim i As Long, j As Long, k As Long
  Dim q As vInt, w As vInt, s As String
  
  DispClear
  For i = 1 To 20
    j = (2 ^ 31 - 1) * Rnd
    q = viFromL(j)
    w = viKRoot(q, viFromL(7))
    List1.AddItem "Given q =" _
                 & FmtInt(j, 11) _
                 & ", viKRoot(q,7)=" _
                 & viFmt(w, 7) _
                 & ", Excess=" _
                 & viFmt(viSub(q, viPow(w, viFromL(7))), 7)
  Next
End Sub

Private Sub VLI_LCDLCM()
  Dim i As Long
  Dim viR1 As vInt, viR2 As vInt, Vi232 As vInt
  
  DispClear
  Vi232 = viTwoToThe(viFromL(32))
  For i = 1 To 30
    viR1 = viRand(Vi232)
    viR2 = viRand(Vi232)
    List1.AddItem "viGCD(" & viFmt(viR1, 11) & "," & viFmt(viR2, 11) & ") =" & viFmt(viGCD(viR1, viR2), 5) _
                & ",  viLCM(" & viFmt(viR1, 11) & "," & viFmt(viR2, 11) & ") =" & viFmt(viLCM(viR1, viR2), 11)
  Next
End Sub

Private Sub VLI_Fact()
  Dim i As Long, j As Long
  
  DispClear
  For i = 0 To 32
    List1.AddItem "viFact(" & FmtInt(i, 3) & ") = " & viFmt(viFact(viFromL(i)), 40)
  Next
  Disp ""
  For i = 0 To 32
    List1.AddItem "viNPr( 32," & FmtInt(i, 3) & ") = " & viFmt(viNPr(viFromL(32), viFromL(i)), 40)
  Next
  Disp ""
  For i = 0 To 32
    List1.AddItem "viNCr( 32," & FmtInt(i, 3) & ") = " & viFmt(viNCr(viFromL(32), viFromL(i)), 40)
  Next
End Sub

Private Sub VLI_Rand()
  Dim z(0 To 99) As Integer
  Dim i As Long, j As Long
  Dim vi100 As vInt
  
  DispClear
  List1.AddItem "working...."
  DoEvents
  Call vimRandomize     ' uses timer for random values
  
  ' The uniform number generator
  vi100 = viTenToThe(viTwo)
  For i = 1 To 100000
    j = viToL(viRand(vi100))
    z(j) = z(j) + 1
  Next
  
  DispClear
  List1.AddItem ""
  List1.AddItem "Distribution of 100000 uniform random variates"
  For i = 0 To 99
    List1.AddItem "Bin " & i & " has " & z(i)
  Next

End Sub

'========================== VLF Demos ==================================================================

Private Sub VLF_AddSubMulDivMod()
  DispClear
'  vfASMDM 0, 0
'  vfASMDM 1, 0
  vfASMDM 0, 1
'  vfASMDM -1, 0
  vfASMDM 0, -1
  vfASMDM 1, 1
  vfASMDM -1, -1
  vfASMDM 1, -1
  vfASMDM -1, 1
  vfASMDM 1, 2
  vfASMDM -1, -2
  vfASMDM 1, -2
  vfASMDM -1, 2
  vfASMDM 3, -2
  vfASMDM -3, 2
  vfASMDM 2, -3
  vfASMDM -2, 3
  vfASMDM 2, 3
  vfASMDM -2, -3
  vfASMDM 3, 2
  vfASMDM -3, -2
End Sub

Private Sub vfASMDM(x As Long, y As Long)
  Dim p As vFrac, q As vFrac
  
  p = vfFromL(x)
  q = vfFromL(y)
  List1.AddItem "vfAdd(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & vfToStr(vfAdd(p, q)) _
              & ",   vfSub(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & vfToStr(vfSub(p, q)) _
              & ",   vfMul(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & vfToStr(vfMul(p, q)) _
              & ",   vfDiv(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & vfToStr(vfDiv(p, q)) _
              & ",   vfMod(" & FmtInt(x, 2) & "," & FmtInt(y, 2) & ")= " & vfToStr(vfMod(p, q))
End Sub

Private Sub VLF_IntFrcCeilFloor()
  DispClear
'  IFFC 0, 0
'  IFFC 1, 0
  IFFC 0, 1
'  IFFC -1, 0
  IFFC 0, -1
  IFFC 1, 1
  IFFC -1, -1
  IFFC 1, -1
  IFFC -1, 1
  IFFC 1, 2
  IFFC -1, -2
  IFFC 1, -2
  IFFC -1, 2
  IFFC 3, -2
  IFFC -3, 2
  IFFC 2, -3
  IFFC -2, 3
  IFFC 2, 3
  IFFC -2, -3
  IFFC 3, 2
  IFFC -3, -2
End Sub

Private Sub IFFC(x As Long, y As Long)
  Dim z As vFrac
  
  z = vfFromL(x, y)
  List1.AddItem "vfInt(" & vfToStr(z) & ")= " & vfToStr(vfInt(z)) _
              & ",   vfFrc(" & vfToStr(z) & ")= " & vfToStr(vfFrc(z)) _
              & ",   vfFloor(" & vfToStr(z) & ")= " & vfToStr(vfFloor(z)) _
              & ",   vfCeil(" & vfToStr(z) & ")= " & vfToStr(vfCeil(z))
End Sub

Private Sub VLF_Sqrt()
  Dim i As Long, j As Long, k As Long
  Dim q As vFrac, w As vFrac, s As String
  
  DispClear

  For i = 1 To 25
    q = vfFromL(i, 1)
    w = vfSqrt(q)
    List1.AddItem "vfSqrt(" _
                 & vfFmtDec(q, 11, 0) _
                 & ")=" _
                 & vfFmtDec(w, 30) _
                 & ",   XSqr = " _
                 & vfFmtDec(vfMul(w, w), 33)
    DoEvents
  Next
  Disp ""
  For i = 1 To 25
    q = vfFromL(1, i)
    w = vfSqrt(q)
    List1.AddItem "vfSqrt(" _
                 & vfFmtDec(q, 11, 5) _
                 & ")=" _
                 & vfFmtDec(w, 30) _
                 & ",   XSqr = " _
                 & vfFmtDec(vfMul(w, w), 33)
    DoEvents
  Next
  Disp ""
  For i = 1 To 29
    j = (2 ^ 31 - 1) * Rnd
    q = vfFromL(j)
    w = vfSqrt(q)
    List1.AddItem "vfSqrt(" _
                 & vfFmtDec(q, 11, 0) _
                 & ")=" _
                 & vfFmtDec(w, 30) _
                 & ",   XSqr = " _
                 & vfFmtDec(vfMul(w, w), 33)
    DoEvents
  Next

End Sub
  
Private Sub VLF_Pow()
  Dim i As Long, j As Long, k As Long
  Dim q As vFrac, w As vFrac, s As String
  
  DispClear
  List1.AddItem "-33^13= " & vfToStr(vfPow(vfFromL(-33), vfFromL(13)))
  List1.AddItem " 31^15= " & vfToStr(vfPow(vfFromL(31), vfFromL(15)))
  List1.AddItem " 11^33= " & vfToStr(vfPow(vfFromL(11), vfFromL(33)))
  DoEvents
  Disp ""
  List1.AddItem " 33/5^13/3= " & vfToDecStr(vfPow(vfFromL(33, 5), vfFromL(13, 3)), 10)
  List1.AddItem " 31/4^15/4= " & vfToDecStr(vfPow(vfFromL(31, 4), vfFromL(15, 4)), 10)
  List1.AddItem " 11/3^33/5= " & vfToDecStr(vfPow(vfFromL(11, 3), vfFromL(33, 5)), 10)
  DoEvents
  Disp ""
  For i = -10 To 10
    q = vfFromL(i, 10)
    List1.AddItem " 2^" & vfFmtDec(q, 4, 1) & "=" & vfFmtDec(vfTwoToThe(q), 15, 10) _
                & ",  10^" & vfFmtDec(q, 4, 1) & "=" & vfFmtDec(vfTenToThe(q), 15, 10) _
                & ",  13^" & vfFmtDec(q, 4, 1) & "=" & vfFmtDec(vfPow(vfFromL(13), q), 15, 10)
    DoEvents
  Next
End Sub

Private Sub VLF_Convergent()
  
  Dim i As Long, j As Long
  Dim z As vFrac
  
  z = vfcPi()
  DispClear
  List1.AddItem "Convergents of Pi" & vfToDecStr(z, 40)
  For i = 1 To 32
    List1.AddItem "C(" & FmtInt(i, 3) & ")= " _
                & vfFmtDec(vfConvergent(z, i), 22, 20) _
                & ",  " _
                & vffmt(vfConvergent(z, i), 0)
  Next
  DoEvents
  Disp ""
  z = vfcLog10
  List1.AddItem "Convergents of Log10" & vfToDecStr(z, 40)
  For i = 1 To 32
    List1.AddItem "C(" & FmtInt(i, 3) & ")= " _
                & vfFmtDec(vfConvergent(z, i), 22, 20) _
                & ",  " _
                & vffmt(vfConvergent(z, i), 0)
  Next
  DoEvents
  Disp ""
  z = vfcLog2
  List1.AddItem "Convergents of Log2" & vfToDecStr(z, 40)
  For i = 1 To 32
    List1.AddItem "C(" & FmtInt(i, 3) & ")= " _
                & vfFmtDec(vfConvergent(z, i), 22, 20) _
                & ",  " _
                & vffmt(vfConvergent(z, i), 0)
  Next
  DoEvents
  Disp ""
  z = vfcE
  List1.AddItem "Convergents of e" & vfToDecStr(z, 40)
  For i = 1 To 32
    List1.AddItem "C(" & FmtInt(i, 3) & ")= " _
                & vfFmtDec(vfConvergent(z, i), 22, 20) _
                & ",  " _
                & vffmt(vfConvergent(z, i), 0)
  Next
  
End Sub

Private Sub VLF_ExpLog()
  Dim i As Long, j As Long
  Dim vs As vFrac, vc As vFrac, q As vFrac, x As vFrac, y As vFrac, z As vFrac, z2 As vFrac

  DispClear
  For i = -10 To 10
    q = vfFromL(i)
    x = vfExp(q)
    z = vfLog(x)
    List1.AddItem "Exp(" & vfFmtDec(q, 8, 4) & ")= " & vfFmtDec(x, 30) _
             & "   Log(" & vfFmtDec(x, 8, 4) & "...)= " & vfFmtDec(z, 30)
    DoEvents
  Next
  Disp ""
  For i = -10 To 10
    If i <> 0 Then
      q = vfFromL(1, i)
      x = vfExp(q)
      z = vfLog(x)
    List1.AddItem "Exp(" & vfFmtDec(q, 8, 4) & ")= " & vfFmtDec(x, 30) _
             & "   Log(" & vfFmtDec(x, 8, 4) & "...)= " & vfFmtDec(z, 30)
      DoEvents
    End If
  Next
End Sub

Private Sub VLF_SinASin()
  Dim i As Long, j As Long
  Dim vs As vFrac, vc As vFrac, q As vFrac, x As vFrac, y As vFrac, z As vFrac, z2 As vFrac

  DispClear
  For i = 0 To 360 Step 10
    q = vfDtoR(vfFromL(i))
    Call vfSinCos(q, y, x)
    z = vfRtoD(vfASin(y))
    List1.AddItem "Sin(" & FmtInt(i, 4) & ")= " & vfToDecStr(y) _
             & "  ASin(" & vfFmtDec(y, 8, 4) & "...)= " & vfToDecStr(z)
    DoEvents
  Next
End Sub

Private Sub VLF_CosACos()
  Dim i As Long, j As Long
  Dim vs As vFrac, vc As vFrac, q As vFrac, x As vFrac, y As vFrac, z As vFrac, z2 As vFrac

  DispClear
  For i = 0 To 360 Step 10
    q = vfDtoR(vfFromL(i))
    Call vfSinCos(q, y, x)
    z = vfRtoD(vfAcos(x))
    List1.AddItem "Cos(" & FmtInt(i, 4) & ")= " & vfToDecStr(x) _
             & "  ACos(" & vfFmtDec(y, 8, 4) & "...)= " & vfToDecStr(z)
    DoEvents
  Next
End Sub

Private Sub VLF_TanATan()
  Dim i As Long, j As Long
  Dim vs As vFrac, vc As vFrac, q As vFrac, x As vFrac, y As vFrac, z As vFrac, z2 As vFrac

  DispClear
  For i = 0 To 360 Step 10
    q = vfDtoR(vfFromL(i))
    y = vfTan(q)
    z = vfRtoD(vfATan(y))
    List1.AddItem "Tan(" & FmtInt(i, 4) & ")= " & vfToDecStr(y) _
             & "  ATan(" & vfFmtDec(y, 8, 4) & "...)= " & vfToDecStr(z)
    DoEvents
  Next
End Sub


Private Function Gauss(a() As vFrac, b() As vFrac, n As Long) As Boolean

  'solve Ax=b using Gaussian elimination solution x returned in b
  Dim i As Long, j As Long, k As Long, m As Long
  Dim ok As Boolean
  Dim w As vFrac, s As vFrac

  w = vfZero
  s = vfZero
  ok = True
  For i = 0 To n - 1
    a(i, n) = b(i)
  Next
  
  For i = 0 To n - 1       ' Gaussian elimination
    m = i
    For j = i + 1 To n - 1
      w = vfAbs(a(j, i))
      s = vfAbs(a(m, i))
      If vfCmp(w, s) = HIGHER Then m = j
    Next
    If m <> i Then
      For k = i To n - 1
        w = a(i, k)
        a(i, k) = a(m, k)
        a(m, k) = w
      Next
    End If
    If vfisZero(a(i, i)) Then
      ok = False
      Exit For
    End If
    
    For j = i + 1 To n - 1
      s = vfDiv(a(j, i), a(i, i))
      For k = n To i Step -1
        w = vfMul(s, a(i, k))
        a(j, k) = vfSub(a(j, k), w)
      Next
    Next
  Next
  Call ListMatrices(a, b, n, "After Elimination")
  
  If ok Then
    For j = n - 1 To 0 Step -1  ' Backward substitution
      s = vfZero
      For k = j + 1 To n - 1
        w = vfMul(a(j, k), b(k))
        s = vfAdd(s, w)
      Next
      w = vfSub(a(j, n), s)
      If vfisZero(a(j, j)) Then
        ok = False
        Exit For
      End If
      b(j) = vfDiv(w, a(j, j))
    Next
  End If
  Call ListMatrices(a, b, n, "After BackSubstitution")
  
  'check
  w = vfZero
  For i = 0 To n - 1
    w = vfAdd(w, vfMul(a(0, i), b(i)))
  Next
  List1.AddItem vfToStr(vfSub(w, vfOne))
  Gauss = ok
End Function

Private Sub ListMatrices(a() As vFrac, b() As vFrac, n As Long, t As String)
  Dim i As Long, j As Long
  Dim z As String
  
  Disp ""
  List1.AddItem t
  For i = 0 To n - 1
    z = ""
    For j = 0 To n - 1
      z = z & vffmt(a(i, j), 10) & "   "
    Next
    List1.AddItem z & " = " & vffmt(b(i), 10)
  Next
  DoEvents
End Sub
 
 ' Solve set of linear equations involving a Hilbert matrix
 ' i.e. solves   Hx=b, where b is the vector [1,1,1....1]

Private Sub VLF_Hilbert(ByVal Order As Long)    ' solve set of linear equations
  Dim i As Long, j As Long, n As Long
  Dim a() As vFrac, b() As vFrac
  
  DispClear
  n = Order
  ReDim a(0 To n - 1, 0 To n) As vFrac
  ReDim b(0 To n - 1) As vFrac
  For i = 0 To n - 1
    a(i, n) = vfZero
    b(i) = vfOne
    For j = 0 To n - 1
      a(i, j) = vfFromL(1, i + j + 1)
    Next
  Next
  Call ListMatrices(a, b, n, "Problem Matrix")
  
  If Gauss(a, b, n) Then
    List1.AddItem "Solution is"
    For i = 0 To n - 1
      List1.AddItem "x(" & FmtInt(i + 1, 2) & ") = " & vfFmtDec(b(i), 10, 5)
    Next
  Else
    List1.AddItem "No Solution"
  End If
End Sub


