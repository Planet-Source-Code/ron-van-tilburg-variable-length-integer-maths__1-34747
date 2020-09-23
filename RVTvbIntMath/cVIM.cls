VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVIM"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variable Integer Maths - cVIM.cls
'
'Variable length Integer Maths  - vInt  =    8 + 16k  bits where k=1,2,3...
'Variable length Rational Maths - vFrac = 2*(8 + 16k) bits where k=1,2,3...
'===============================================================================================================
'Values are manipulated in variable length structures called vInts where these are always arrays as
'a multiple of ints (16bits) but where the items are addressed in byte sized quantities
'The length of a vInt is given by the dimension of its internal array of bytes
'as this can be quite large a vInt can in theory represent values of +-255 ^ 2,147,483,646 (a FLN)
'in practice only a string representing +-10^2,147,483,646 -1 (also a FLN) can be displayed
'A VFrac is made of two Vints, one as numerator, one as denominator. The sign is carried by the numerator

'Copyright ©2002 R. van Tilburg
'Public rights for educational use are given. Commercial rights retained by author
'===============================================================================================================

Public Enum vimCMP_VALUES   'when Comparing Signed values
  HIGHER = 1
  EQUAL = 0
  LOWER = -1
End Enum

Public Enum vimTST_VALUES   'when Testing signed values
  POSITIVE = 1
  ZERO = 0
  NEGATIVE = -1
End Enum

Public Type vInt
  s() As Byte             'SiB is sign &HFF is Negative, &H01 is Positive>0, &H00 is POSITIVE=ZERO
End Type

Public Type vFrac
  Numer As vInt           'SiB is sign &HFF is Negative, &H01 is Positive>0, &H00 is POSITIVE=ZERO
  Denom As vInt           'SiB is always &H01=POSITIVE
End Type

'===============================================================================================================
Private Const BASE      As Long = 256
Private Const BASEMAX   As Long = BASE - 1

Private Const LoB       As Long = 1
Private Const SiB       As Long = 0
Private Const SiBNeg    As Byte = &HFF
Private Const SiBZero   As Byte = &H0
Private Const SiBPos    As Byte = &H1

Private Reduce          As Boolean
Private Precision       As Long
Private Const MAXBITS   As Long = 64     'default internal precision for vfFrac (should be multiple of 8)

'constants calculated first when needed to MAXBITS+8 accuracy
Private cE      As vFrac 'Fe      2.7181 etc
Private cPi     As vFrac 'FPi     3.1416 etc
Private c2Pi    As vFrac 'FPi     2Pi
Private cPi2    As vFrac 'FPi     Pi/2
Private cLog2   As vFrac 'FLog2
Private cLog10  As vFrac 'FLog10
Private cMaxVal As vFrac 'FMaxVal 2^MAXBITS-1

'Flags to show which constants have been loaded
Private ConstF        As Long
Private Const Fe      As Long = &H1  'e
Private Const FPi     As Long = &H2  'Pi,2Pi,Pi/2
Private Const FLog2   As Long = &H4  'Log2
Private Const FLog10  As Long = &H8  'Log10
Private Const FMaxVal As Long = &H10

'first 100 decimal places of (about 332 bits worth)
Private Const sE  As String = "2.7182818284 5904523536 0287471352 6624977572 4709369995 9574966967 6277240766 3035354759 4571382178 5251664274"
Private Const sPI As String = "3.1415926535 8979323846 2643383279 5028841971 6939937510 5820974944 5923078164 0628620899 8628034825 3421170680"

'Error messages
Private Const ErrStrDivZero As String = "Division by Zero"
Private Const ErrValDivZero As Long = vbObjectError + 1

Private Const ErrStrArgNegZero As String = "Argument Negative or Zero"
Private Const ErrValArgNegZero As Long = vbObjectError + 2

Private Const ErrStrNegPow As String = "(Negative) power of Negative argument"
Private Const ErrValNegPow As Long = vbObjectError + 3

Private Const ErrStrTooBig As String = "Value too Large"
Private Const ErrValTooBig As Long = vbObjectError + 4

Private Const ErrStrNegRoot As String = "Attempt to take Root of Negative Argument"
Private Const ErrValNegRoot As Long = vbObjectError + 5

Private Const ErrStrOverflow As String = "Value overflow, string too long"
Private Const ErrValOverflow As Long = vbObjectError + 6

Private Const ErrStrERoot As String = "Attempting to take an even root of a negative number"
Private Const ErrValERoot As Long = vbObjectError + 7

Private Const ErrStrGreaterOne As String = "Absolute value of argument > 1"
Private Const ErrValGreaterOne As Long = vbObjectError + 8

'===============================================================================================================

Private Sub Class_Initialize()

  Reduce = True
  Precision = MAXBITS
  Call Randomize

End Sub

'================================================================================================================
'================================== VFRACS ======================================================================
'================================================================================================================

'Absolute value
Public Function vfAbs(ByRef x As vFrac) As vFrac
Attribute vfAbs.VB_Description = "Return the Absolute value of a vFrac"

  vfAbs.Numer = viAbs(x.Numer)
  vfAbs.Denom = x.Denom
  'already normalised

End Function

'acos(x)
Public Function vfACos(ByRef x As vFrac) As vFrac
Attribute vfACos.VB_Description = "The angle for which x is the Cos -Pi/2<Acos<Pi/2"

  If vfCmpOne(vfAbs(x)) = HIGHER Then
    vfACos = vfZero
    Err.Raise ErrValGreaterOne, "RVTvbIM.ACos()", ErrStrGreaterOne
   ElseIf vfIsZero(x) Then
    vfACos = vfcPi2
   ElseIf vfCmpOne(vfAbs(x)) = EQUAL Then
    vfACos = vfZero
   Else
    vfACos = vfATan(vfDiv(vfPythSub(vfOne, x), x))
  End If

End Function

'Signed addition of two vFrac Numbers
Public Function vfAdd(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfAdd.VB_Description = "Add two vFracs together"

  vfAdd.Numer = viAdd(viMul(x.Numer, y.Denom), viMul(y.Numer, x.Denom))
  vfAdd.Denom = viMul(x.Denom, y.Denom)
  Call vfNormalise(vfAdd)

End Function

'asin(x)    -1<=x<=1
Public Function vfASin(ByRef x As vFrac) As vFrac
Attribute vfASin.VB_Description = "The angle for which x is the Sin -Pi/2 < Asin < Pi/2"

  If vfCmpOne(vfAbs(x)) = HIGHER Then
    vfASin = vfZero
    Err.Raise ErrValGreaterOne, "RVTvbIM.ASin()", ErrStrGreaterOne
   ElseIf vfIsZero(x) Then
    vfASin = vfZero
   ElseIf vfCmpOne(vfAbs(x)) = EQUAL Then
    vfASin = vfcPi2
    If vfIsNeg(x) Then vfASin = vfNeg(vfASin)
   Else
    vfASin = vfATan(vfDiv(x, vfPythSub(vfOne, x)))
  End If

End Function

'atan(x)  result is -PI2<= x <=Pi2
Public Function vfATan(ByRef x As vFrac) As vFrac
Attribute vfATan.VB_Description = "The angle for which x is the Tan -Pi/2 < Atan < Pi/2"

 Dim i As Long, z As vFrac, z2 As vFrac, p1 As vInt, p2 As vInt

  Precision = MAXBITS + 4
  z = vfAbs(x)
  If vfIsZero(z) Then
    vfATan = vfZero
   ElseIf vfCmpOne(z) = EQUAL Then
    vfATan = vfDiv(vfcPi2, vfTwo)
   Else
    If vfCmpOne(z) = HIGHER Then z = vfInv(z)

    z = vfDiv(vfSub(vfPythAdd(vfOne, z), vfOne), z)   'z is now tan(x/2) z<sqrt2-1 ie. 0.41421..
    z2.Numer = viMul(z.Numer, z.Numer)
    z2.Denom = viMul(z.Denom, z.Denom)

    vfATan = vfOne
    p1 = viFromL(Precision - 1)
    p2 = viFromL(Precision + 1)
    For i = Precision \ 2 To 1 Step -1        'should do as a max
      vfATan.Denom = viMul3(vfATan.Denom, z2.Denom, p2)
      vfATan.Numer = viSub(vfATan.Denom, viMul3(vfATan.Numer, z2.Numer, p1))
      p1 = viSub(p1, viTwo)
      p2 = viSub(p2, viTwo)
    Next i
    vfATan.Numer = viMul3(viTwo, z.Numer, vfATan.Numer)
    vfATan.Denom = viMul(z.Denom, vfATan.Denom)
    Call vfNormalise(vfATan)
    If vfCmpOne(vfAbs(x)) = HIGHER Then       'atan(x)=Pi/2-atan(1/x)
      vfATan = vfSub(vfcPi2, vfATan)
    End If
  End If
  If vfIsNeg(x) Then vfATan = vfNeg(vfATan)
  Precision = MAXBITS
  Call vfNormalise(vfATan)

End Function

'atan(y/x) result is -2PI <= x <= 2PI  .ie In the right Quadrant - in c ATan2 will only give -Pi<=x<=Pi
Public Function vfATan2(ByRef y As vFrac, ByRef x As vFrac) As vFrac
Attribute vfATan2.VB_Description = "The angle which is y/x is the Tan. Returns ATan in  the correct quadrant"

  If vfIsZero(x) And vfIsZero(y) Then
    vfATan2 = vfZero                                            'y=0/x=0 => 0
   ElseIf vfIsZero(x) Then
    vfATan2 = vfcPi2                                            'y/x=0   => PI/2 or -PI/2
    If vfIsNeg(y) Then vfATan2 = vfSub(vfc2Pi, vfATan2)
   ElseIf vfIsZero(y) Then
    If vfIsNeg(x) Then vfATan2 = vfcPi Else vfATan2 = vfZero    'y=0/x => 0 or PI
   Else
    vfATan2 = vfATan(vfDiv(y, x))                               '-PI/2 <= Atan(y/x) <= PI/2
    If vfIsNeg(x) And vfIsNeg(y) Then
      vfATan2 = vfAdd(vfcPi, vfATan2)                           'Atan2 = -PI + Atan2 (which is positive)
     ElseIf vfIsNeg(x) Then
      vfATan2 = vfAdd(vfcPi, vfATan2)                           'Atan2 =  PI + Atan2 (which is negative)
     ElseIf vfIsNeg(y) Then
      vfATan2 = vfAdd(vfc2Pi, vfATan2)                          'Atan2 =  PI + Atan2 (which is negative)
    End If
  End If

End Function

'return (b*u + a*v) ' used in evaluation of continued fractions
Private Function vfBuPlusAv(ByRef b As vFrac, ByRef u As vFrac, ByRef a As vFrac, ByRef v As vFrac) As vFrac

 Dim xa As vInt, xb As vInt, xc As vInt, xd As vInt

  xa = viMul(b.Numer, u.Numer)
  xb = viMul(a.Denom, v.Denom)
  xc = viMul(a.Numer, v.Numer)
  xd = viMul(b.Denom, u.Denom)
  vfBuPlusAv.Numer = viAdd(viMul(xa, xb), viMul(xc, xd))
  vfBuPlusAv.Denom = viMul(xb, xd)
  Call vfNormalise(vfBuPlusAv)

End Function

'2Pi
Public Property Get vfc2Pi() As vFrac
Attribute vfc2Pi.VB_Description = "Return the constant  2Pi=6.283..."

  If (ConstF And FPi) = 0 Then
    vfc2Pi = vfMul(vfcPi, vfTwo)
   Else
    vfc2Pi = c2Pi
  End If

End Property

Public Property Get vfcE() As vFrac
Attribute vfcE.VB_Description = "Return the constant e = 2.718182..."

  If (ConstF And Fe) = 0 Then
    Precision = MAXBITS + 4
    cE = vfFromDecStr(sE)
    ConstF = ConstF Or Fe
    Precision = MAXBITS
  End If
  vfcE = cE

End Property

'Return integer next higher from value
Public Function vfCeil(ByRef x As vFrac) As vFrac
Attribute vfCeil.VB_Description = "Return the nearest integer greater than x"

 Dim z As vFrac

  Call vfIntFrc(x, vfCeil, z)
  If Not vfIsInteger(x) Then                              'integer has denom=1
    If Not vfIsNeg(x) Then vfCeil = vfAdd(vfCeil, vfOne)
  End If
  'already normalised

End Function

'Log10 = 2.302585093...
Public Property Get vfcLog10() As vFrac
Attribute vfcLog10.VB_Description = "Return the constant Log10=2.302..."

  If (ConstF And FLog10) = 0 Then
    Precision = MAXBITS + 4
    cLog10 = vfLog(vfTen)
    ConstF = ConstF Or FLog10
    Precision = MAXBITS
  End If
  vfcLog10 = cLog10

End Property

'Log2=0.693147181...
Public Property Get vfcLog2() As vFrac
Attribute vfcLog2.VB_Description = "Return the constant Log2 = 0.6931..."

  If (ConstF And FLog2) = 0 Then
    Precision = MAXBITS + 4
    cLog2 = vfLog(vfTwo)
    ConstF = ConstF Or FLog2
    Precision = MAXBITS
  End If
  vfcLog2 = cLog2

End Property

'Maximum value at Precision
Public Property Get vfcMaxVal() As vFrac
Attribute vfcMaxVal.VB_Description = "The largets representable value at background precision "

  If (ConstF And FMaxVal) = 0 Then
    cMaxVal.Numer = viSub(viTwoToThe(viFromL(Precision)), viOne)
    cMaxVal.Denom = viOne
    ConstF = ConstF Or FMaxVal
  End If
  vfcMaxVal = cMaxVal

End Property

'Taking account of sign --- returns -1 if x<y,0 if x=y, +1 if x>y
Public Function vfCmp(ByRef x As vFrac, ByRef y As vFrac) As Long
Attribute vfCmp.VB_Description = "Compare two vFRacs, returns HIGHER, EQUAL, LOWER for first relative to second"

 Dim z As vFrac

  z = vfSub(x, y)
  If vfIsZero(z) Then         'within Precision
    vfCmp = EQUAL
   ElseIf vfIsNeg(z) Then
    vfCmp = LOWER
   Else
    vfCmp = HIGHER
  End If

End Function

'signed compare with One
Public Function vfCmpOne(ByRef x As vFrac) As Long
Attribute vfCmpOne.VB_Description = "Compare x with One, if x>one return HIGHER, x=1 then EQUAL, x<1 LOWER"

  vfCmpOne = viCmp(x.Numer, x.Denom)

End Function

'Return the Nth Convergent of a vfrac inside the Precision range
Public Function vfConvergent(ByRef x As vFrac, ByVal NthC As Long) As vFrac
Attribute vfConvergent.VB_Description = "Return the nth convergent (fractional approximation) if a vFrac"

 Dim i As Long
 Dim a As vInt, b As vInt, v As vInt, w As vInt
 Dim p0 As vInt, p1 As vInt, q0 As vInt, q1 As vInt

  If NthC <= 0 Then NthC = 1
  i = 0
  a = viAbs(x.Numer)
  b = x.Denom
  Do While Not viIsZero(b)  'unroll continued fraction terms with Euclids Algorithm
    v = viDiv(a, b)
    w = b
    b = viSub(a, viMul(v, w))
    a = w

    If i = 0 Then           'now reroll it
      p1 = viOne
      p0 = v
      q1 = viZero
      q0 = viOne
      vfConvergent.Numer = p0
      vfConvergent.Denom = q0
     Else
      vfConvergent.Numer = viAdd(viMul(v, p0), p1)
      vfConvergent.Denom = viAdd(viMul(v, q0), q1)
      p1 = p0
      q1 = q0
      p0 = vfConvergent.Numer
      q0 = vfConvergent.Denom

      If viULogB2(q0) > Precision Then
        vfConvergent.Numer = p1
        vfConvergent.Denom = q1
        Exit Do                     'its big enough so quit
      End If
    End If
    If i = NthC Then Exit Do
    i = i + 1
  Loop
  If vfIsNeg(x) Then vfConvergent = vfNeg(vfConvergent)
  'this is normalised (by defn, but could be very out of BitSize range in the numerator)

End Function

'Duplicates src vFrac
Public Function vfCopy(ByRef src As vFrac) As vFrac
Attribute vfCopy.VB_Description = "Duplicate src into dest with Normalisation"

  vfCopy.Numer = viCopy(src.Numer)
  vfCopy.Denom = viCopy(src.Denom)
  Call vfNormalise(vfCopy)

End Function

'Pi
Public Property Get vfcPi() As vFrac
Attribute vfcPi.VB_Description = "Return the consatnt Pi=3.141592..."

  If (ConstF And FPi) = 0 Then
    Precision = MAXBITS + 4
    cPi = vfFromDecStr(sPI)
    c2Pi = vfMul(cPi, vfTwo)
    cPi2 = vfDiv(cPi, vfTwo)
    ConstF = ConstF Or FPi
    Precision = MAXBITS
  End If
  vfcPi = cPi

End Property

'Pi/2
Public Property Get vfcPi2() As vFrac
Attribute vfcPi2.VB_Description = "Return the constant PI/2 = 1.57079..."

  If (ConstF And FPi) = 0 Then
    vfcPi2 = vfDiv(vfcPi, vfTwo)
   Else
    vfcPi2 = cPi2
  End If

End Property

'Signed Division - result = arithmetic quotient
Public Function vfDiv(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfDiv.VB_Description = "Divide y into x"

  vfDiv.Numer = viMul(x.Numer, y.Denom)
  vfDiv.Denom = viMul(x.Denom, y.Numer)
  Call vfNormalise(vfDiv)

End Function

'Signed Division and Modulo - q = arithmetic quotient, r= remainder has sign of divisor
Public Sub vfDivMod(ByRef x As vFrac, ByRef y As vFrac, ByRef q As vFrac, ByRef r As vFrac)
Attribute vfDivMod.VB_Description = "Return quotient and Modulo of x/y"

  q.Numer = viMul(x.Numer, y.Denom)
  q.Denom = viMul(x.Denom, y.Numer)
  Call vfNormalise(q)
  r.Numer = viMod(q.Numer, q.Denom)
  r.Denom = viMul(x.Denom, y.Denom)
  Call vfNormalise(r)

End Sub

'Degrees to Radians
Public Function vfDtoR(ByRef x As vFrac) As vFrac
Attribute vfDtoR.VB_Description = "convert Degrees to Radians"

  vfDtoR.Numer = viMul(x.Numer, vfcPi.Numer)
  vfDtoR.Denom = viMul(x.Denom, viMul(vfcPi.Denom, viFromL(180)))
  Call vfNormalise(vfDtoR)

End Function

'exp(x)  , exp(i+f) = exp(i)*expf(f)  where i=Int(x),f=Frc(x)
Public Function vfExp(ByRef x As vFrac) As vFrac
Attribute vfExp.VB_Description = "return e^x"

 Dim ei As vFrac, ef As vFrac

  vfExp = vfOne
  If vfIsZero(x) Then Exit Function

  Call vfIntFrc(x, ei, ef)
  vfExp = vfMul(vfPow(vfcE, ei), vfExpF(ef))

End Function

'exp(x) = e^(int(x))*[1+x/1(1 + x/2(1+x/3(1+x/4...(1+x/k)]  for all x, here we keep -1<x<1 at all times
Private Function vfExpF(ByRef xf As vFrac) As vFrac

 Dim i As Long

  vfExpF = vfOne
  If vfIsZero(xf) Then Exit Function

  For i = Precision \ 3 To 1 Step -1
    vfExpF.Denom = viMul(viFromL(i), viMul(vfExpF.Denom, xf.Denom))
    vfExpF.Numer = viAdd(vfExpF.Denom, viMul(vfExpF.Numer, xf.Numer))
  Next i
  Call vfNormalise(vfExpF)

End Function

'x=floor(x) , = x for integers, int(x) if x>0, int(x)-1 if x<0
Public Function vfFloor(ByRef x As vFrac) As vFrac
Attribute vfFloor.VB_Description = "return the nearest integer smaller than x"

 Dim z As vFrac

  Call vfIntFrc(x, vfFloor, z)
  If Not viUCmpOne(x.Denom) = EQUAL Then              'integers have denom=1
    If vfIsNeg(x) Then vfFloor = vfSub(vfFloor, vfOne)
  End If
  'already normalised

End Function

'Format a VFrac into a string of width , width <0 then left aligned
Public Function vfFmt(ByRef x As vFrac, ByVal Width As Long) As String
Attribute vfFmt.VB_Description = "Format a vFRac in fractional display form"

  vfFmt = vfToStr(x)
  If Len(vfFmt) < Abs(Width) Then
    If Width < 0 Then
      vfFmt = vfFmt & String$(Abs(Width) - Len(vfFmt), 32)
     Else
      vfFmt = String$(Width - Len(vfFmt), 32) & vfFmt
    End If
  End If

End Function

'Format a VFrac into a string of width , width <0 then left aligned
Public Function vfFmtDec(ByRef x As vFrac, ByVal Width As Long, Optional ByVal ndec As Long = -1) As String
Attribute vfFmtDec.VB_Description = "Format a vFrac in Decimal representation"

  vfFmtDec = vfToDecStr(x, ndec)
  If Len(vfFmtDec) < Abs(Width) Then
    If Width < 0 Then
      vfFmtDec = vfFmtDec & String$(Abs(Width) - Len(vfFmtDec), 32)
     Else
      vfFmtDec = String$(Width - Len(vfFmtDec), 32) & vfFmtDec
    End If
  End If

End Function

Public Function vfFrc(ByRef x As vFrac) As vFrac
Attribute vfFrc.VB_Description = "return the fractional part of a vFRac"

 Dim z As vFrac

  Call vfIntFrc(x, z, vfFrc)
  'already normalised

End Function

'Unroll a double into a VFrac
Public Function vfFromD(ByVal x As Double) As vFrac
Attribute vfFromD.VB_Description = "create a vFrac from a double (accuracy may be variable)"

 Dim i As Long, z As Double
 Dim p0 As vInt, p1 As vInt, q0 As vInt, q1 As vInt, v As vInt

  vfFromD = vfZero
  If x = 0# Then Exit Function

  z = Abs(x)
  If z <= 2147483647# Then
    i = 0
    p1 = viOne
    p0 = viFromL(Int(z))
    q1 = viZero
    q0 = viOne
    vfFromD.Numer = p0
    vfFromD.Denom = q0
    z = z - Int(z)
    Do While z <> 0#                    'unroll continued fraction terms with Division Algorithm
      z = 1# / z
      If z > 2147483647# Then Exit Do   'a subsequent term is just too large (approximate to here only)
      v = viFromL(Int(z))

      vfFromD.Numer = viAdd(viMul(v, p0), p1)
      vfFromD.Denom = viAdd(viMul(v, q0), q1)
      p1 = p0
      q1 = q0
      p0 = vfFromD.Numer
      q0 = vfFromD.Denom

      If viULogB2(q0) > Precision Then
        vfFromD.Numer = p1
        vfFromD.Denom = q1
        Exit Do                       'its big enough so quit
      End If
      z = z - Int(z)
    Loop
    'this is normalised (by defn, but could be very out of BitSize range in the numerator)
   Else
    Err.Raise ErrValTooBig, "RVTvbIM.vfFromD", ErrStrTooBig
  End If
  If x < 0# Then vfFromD = vfNeg(vfFromD)

End Function

'Signed conversion and extension from Decimal String to vFrac    has form of +-aaaaa.bbbbbb
Public Function vfFromDecStr(ByRef x As String) As vFrac
Attribute vfFromDecStr.VB_Description = "create a vFrac from a string of decimal characters +/-nnnn.mmmm"

 Dim z As vFrac, zf As vFrac, s As Long, w As String

  On Error GoTo TooLarge
  s = InStr(1, x, ".")
  If s = 0 Then               'whole number
    z.Numer = viFromStr(x)
    z.Denom = viOne
    vfFromDecStr = vfCopy(z)
   Else
    z.Numer = viFromStr(Left$(x, s - 1))
    z.Denom = viOne
    w = Replace$(Mid$(x, s + 1), " ", "")
    zf.Numer = viFromStr(w)
    zf.Denom = viTenToThe(viFromL(Len(w)))
    vfFromDecStr = vfAdd(z, zf)
  End If

Exit Function

TooLarge:
  vfFromDecStr = vfZero
  Err.Raise ErrValTooBig, "RVTvbIM.vfFromDecStr", ErrStrTooBig
  On Error GoTo 0

End Function

'Signed conversion and extension from Long to vFrac
Public Function vfFromL(ByVal Numer As Long, Optional ByVal Denom As Long = 1) As vFrac
Attribute vfFromL.VB_Description = "create a vfrac from Long values"

  vfFromL = vfZero
  If Denom = 0 Then
    vfFromL = vfcMaxVal
    Err.Raise ErrValDivZero, "RVTvbIM.vfFromL", ErrStrDivZero
  End If

  vfFromL.Numer = viFromL(Numer)
  vfFromL.Denom = viFromL(Denom)
  Call vfNormalise(vfFromL)

End Function

'Signed conversion and extension from String to vFrac    has form of +-aaaaa/bbbbbb
Public Function vfFromStr(ByRef x As String) As vFrac
Attribute vfFromStr.VB_Description = "create a vFrac from a fractional string .  +-nnnnn/mmmmm"

 Dim z As vFrac, s As Long

  On Error GoTo TooLarge
  s = InStr(1, x, "/")
  If s = 0 Then
    z.Numer = viFromStr(x)
    z.Denom = viOne
   ElseIf s = 1 Or s = 2 Then
    z.Numer = viOne
    z.Denom = viFromStr(Mid$(x, s + 1))
   Else
    z.Numer = viFromStr(Left$(x, s - 1))
    z.Denom = viFromStr(Mid$(x, s + 1))
  End If
  vfFromStr = vfCopy(z)

Exit Function

TooLarge:
  vfFromStr = vfZero
  Err.Raise ErrValTooBig, "RVTvbIM.vfFromStr", ErrStrTooBig
  On Error GoTo 0

End Function

'Signed conversion and extension from VInt to vFrac
Public Function vfFromvi(ByRef Numer As vInt, ByRef Denom As vInt) As vFrac
Attribute vfFromvi.VB_Description = "create a vFrac from two VInts"

  vfFromvi.Numer = Numer
  vfFromvi.Denom = Denom
  Call vfNormalise(vfFromvi)

End Function

'A first guess of the value of the root (the value shifted right by 1/k its nr of bits)
Private Function vfGuessRoot(ByRef x As vFrac, ByVal k As Integer) As vFrac

 Dim n As Long

  n = viULogB2(x.Numer) - viULogB2(x.Denom)
  vfGuessRoot.Numer = x.Numer
  vfGuessRoot.Denom = viShift(x.Denom, -n \ k)
  vfGuessRoot.Denom.s(LoB) = vfGuessRoot.Denom.s(LoB) Or 1    ' just in case!!!!
  Call vfNormalise(vfGuessRoot)

End Function

'one half 1/2
Public Property Get vfHalf() As vFrac
Attribute vfHalf.VB_Description = "return 1/2"

  vfHalf.Numer = viOne
  vfHalf.Denom = viTwo

End Property

'return integer portion of x
Public Function vfInt(ByRef x As vFrac) As vFrac
Attribute vfInt.VB_Description = "return the Integer portion of a vFrac"

 Dim z As vFrac

  Call vfIntFrc(x, vfInt, z)
  'already normalised

End Function

'split integer and fractional portions of x
Public Sub vfIntFrc(ByRef x As vFrac, ByRef xInt As vFrac, ByRef xFrc As vFrac)
Attribute vfIntFrc.VB_Description = "split a Vfrac into Integer and Fractional parts"

  If vfIsInteger(x) Then
    xInt = x
    xFrc = vfZero
   Else
    Call viDivMod(x.Numer, x.Denom, xInt.Numer, xFrc.Numer)
    xInt.Denom = viOne
    xFrc = vfSub(x, xInt)
    'already normalised
  End If

End Sub

'return 1/x
Public Function vfInv(ByRef x As vFrac) As vFrac
Attribute vfInv.VB_Description = "return 1/x"

  vfInv.Numer = viAbs(x.Denom)
  vfInv.Denom = viAbs(x.Numer)
  If vfIsNeg(x) Then vfInv.Numer = viNeg(vfInv.Numer)
  'already normalised

End Function

'test if integer
Public Function vfIsInteger(ByRef x As vFrac) As Boolean
Attribute vfIsInteger.VB_Description = "returns TRUE if vFrac is an Integer"

  If viCmpOne(x.Denom) = EQUAL Then vfIsInteger = True

End Function

'test if negative
Public Function vfIsNeg(ByRef x As vFrac) As Boolean
Attribute vfIsNeg.VB_Description = "returns TRUE if VFRac is <0"

  If viIsNeg(x.Numer) Then vfIsNeg = True

End Function

'tests if a vFrac is Positive
Public Function vfIsPos(ByRef x As vFrac) As Boolean
Attribute vfIsPos.VB_Description = "returns TRUE if VFrac >=0"

  If viIsPos(x.Numer) Then vfIsPos = True

End Function

'tests if a vFrac is Zero
Public Function vfIsZero(ByRef x As vFrac) As Boolean
Attribute vfIsZero.VB_Description = "returns TRUE if x=0"

  If viIsZero(x.Numer) Then vfIsZero = True

End Function

'return log(x) by series expression
Public Function vfLog(ByRef x As vFrac) As vFrac
Attribute vfLog.VB_Description = "returns the natural log of x"

 Dim i As Long, q As vFrac, w As vFrac, z As vFrac, z2 As vFrac, p1 As vInt, p2 As vInt

  vfLog = vfZero
  If vfIsNeg(x) Or vfIsZero(x) Then
    Err.Raise ErrValArgNegZero, "RVTvbIM.vfLog", ErrStrArgNegZero
   ElseIf Not (vfCmpOne(x) = EQUAL) Then
    Precision = MAXBITS + 4
    q = vfMul(vfULogB2(x), vfFromL(693147181, 1000000000))    'first estimate of z0 .=. lg(x)*log2
    w = vfDiv(x, vfExp(q))
    z.Numer = viSub(w.Numer, w.Denom)
    z.Denom = viAdd(w.Numer, w.Denom)     '= z= (w-1)/(w+1)
    z2.Numer = viMul(z.Numer, z.Numer)
    z2.Denom = viMul(z.Denom, z.Denom)
    p1 = viFromL(Precision \ 2 - 1)
    p2 = viFromL(Precision \ 2 + 1)
    For i = Precision \ 4 To 1 Step -1      'should do as a max
      vfLog.Denom = viMul3(vfLog.Denom, z2.Denom, p2)
      vfLog.Numer = viAdd(vfLog.Denom, viMul3(vfLog.Numer, z2.Numer, p1))
      p1 = viSub(p1, viTwo)
      p2 = viSub(p2, viTwo)
    Next i
    vfLog.Numer = viMul(viShift(z.Numer, -1), vfLog.Numer)
    vfLog.Denom = viMul(z.Denom, vfLog.Denom)
    Precision = MAXBITS
    vfLog = vfAdd(vfLog, q)
  End If

End Function

'Log(x) base 10
Public Function vfLog10(ByRef x As vFrac) As vFrac
Attribute vfLog10.VB_Description = "return common log base 10 of x"

  vfLog10 = vfDiv(vfLog(x), vfcLog10)

End Function

'Log(x) base 2
Public Function vfLog2(ByRef x As vFrac) As vFrac
Attribute vfLog2.VB_Description = "return log base 2 of x"

  vfLog2 = vfDiv(vfLog(x), vfcLog2)

End Function

'Signed Modulo - result = remainder has sign of divfsor
Public Function vfMod(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfMod.VB_Description = "return Modulus of x/y"

  vfMod.Numer = viMod(viMul(x.Numer, y.Denom), viMul(x.Denom, y.Numer))
  vfMod.Denom = viMul(x.Denom, y.Denom)
  Call vfNormalise(vfMod)

End Function

'Signed multiply of two vFrac Numbers
Public Function vfMul(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfMul.VB_Description = "multiply two vFracs together"

  vfMul.Numer = viMul(x.Numer, y.Numer)
  vfMul.Denom = viMul(x.Denom, y.Denom)
  Call vfNormalise(vfMul)

End Function

'Negate a vFrac  x=-x
Public Function vfNeg(ByRef x As vFrac) As vFrac
Attribute vfNeg.VB_Description = "return the Negative of x"

  vfNeg.Numer = viNeg(x.Numer)
  vfNeg.Denom = x.Denom
  'already normalised

End Function

'remove leading zeroes (in pairs) from a vFrac, and verify SignByte
Private Sub vfNormalise(ByRef v As vFrac)

  Call viNormalise(v.Numer)
  Call viNormalise(v.Denom)

  'make sure zero is correctly identified
  If v.Denom.s(SiB) = SiBZero Then
    If vfIsNeg(v) Then v = vfNeg(vfcMaxVal) Else v = vfcMaxVal
    Err.Raise ErrValDivZero, "RVTvbIM.vfNormalise", ErrStrDivZero
    Exit Sub
  End If

  If viIsNeg(v.Denom) Then      'only the numerator holds a sign
    v.Numer = viNeg(v.Numer)
    v.Denom = viNeg(v.Denom)
  End If

  If Reduce Then                  'this can be independently controlled
    v = vfReduce(v)
    If viIsNeg(v.Denom) Then      'only the numerator holds a sign
      v.Numer = viNeg(v.Numer)
      v.Denom = viNeg(v.Denom)
    End If
  End If

End Sub

'the constant 1
Public Property Get vfOne() As vFrac
Attribute vfOne.VB_Description = "return 1"

  vfOne.Numer = viOne
  vfOne.Denom = viOne

End Property

'return x^y , x>=0, y>=0
Public Function vfPow(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfPow.VB_Description = "return x^y"

 Dim w As vFrac, z As vFrac, p As vFrac

  If vfIsZero(y) Or vfCmpOne(x) = EQUAL Then    'x^0 = 1, 1^y = 1
    vfPow = vfOne
    Exit Function
  End If

  If vfIsInteger(y) Then
    z = x: p = y
    If vfIsNeg(y) Then
      z = vfInv(x)
      p = vfAbs(y)
    End If
    vfPow.Numer = viPow(z.Numer, p.Numer)
    vfPow.Denom = viPow(z.Denom, p.Numer)
   Else
    If vfIsZero(x) Then
      vfPow = vfZero
      Exit Function
    End If

    If vfIsZero(y) Then
      vfPow = vfOne
      Exit Function
    End If

    If vfIsNeg(x) Then
      vfPow = vfZero
      Err.Raise ErrValNegPow, "RVTvbIM.vfPow", ErrStrNegPow
      Exit Function
    End If

    z = x: p = y
    If vfIsNeg(y) Then
      z = vfInv(z)
      p = vfAbs(y)
    End If
    Precision = MAXBITS + 4
    If viUCmp(p.Denom, viTwo) = EQUAL Then
      z.Numer = viPow(z.Numer, p.Numer)
      z.Denom = viPow(z.Denom, p.Numer)
      vfPow = vfSqrt(z)
     Else
      vfPow = vfExp(vfMul(p, vfLog(z)))
    End If
    Precision = MAXBITS
  End If
  Call vfNormalise(vfPow)

End Function

'Polar to Rectangular coordinates x = r.cos(theta), y = r.sin(theta),  theta in radians
Public Sub vfPtoR(ByRef r As vFrac, ByRef theta As vFrac, ByRef x As vFrac, ByRef y As vFrac)
Attribute vfPtoR.VB_Description = "return x=r.Cos(Theta),y=r.Sin(Theta)"

  Call vfSinCos(theta, y, x)
  x = vfMul(x, r)
  y = vfMul(y, r)

End Sub

'return sqrt(x^2+y^2)
Public Function vfPythAdd(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfPythAdd.VB_Description = "return sqrt(x.x+y.y)"

  vfPythAdd = vfSqrt(vfAdd(vfMul(x, x), vfMul(y, y)))

End Function

'return sqrt(x^2-y^2); x>y
Public Function vfPythSub(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfPythSub.VB_Description = "return sqrt(x.x-y.y) , y<=x "

  vfPythSub = vfSqrt(vfMul(vfAdd(x, y), vfSub(x, y)))

End Function

'return a random number of maximally MaxVal
Public Function vfRand(ByRef MaxVal As vFrac) As vFrac
Attribute vfRand.VB_Description = "return a uniform random variable"

  vfRand.Numer = viMul(viFromL(UniformRand), MaxVal.Numer)
  vfRand.Denom = viMul(viFromL(2147483563), MaxVal.Denom)
  Call vfNormalise(vfRand)

End Function

'Unroll and Reroll a VFrac to a given Bitsize (try doing this with a joint)
Private Function vfReduce(ByRef x As vFrac) As vFrac

 Dim i As Long
 Dim a As vInt, b As vInt, v As vInt, w As vInt
 Dim p0 As vInt, p1 As vInt, q0 As vInt, q1 As vInt

  i = 0
  a = viAbs(x.Numer)
  b = x.Denom
  Do While Not viIsZero(b)  'unroll continued fraction terms with Euclids Algorithm
    v = viDiv(a, b)
    w = b
    b = viSub(a, viMul(v, w))
    a = w

    If i = 0 Then           'now reroll it
      p1 = viOne
      p0 = v
      q1 = viZero
      q0 = viOne
      vfReduce.Numer = p0
      vfReduce.Denom = q0
     Else
      vfReduce.Numer = viAdd(viMul(v, p0), p1)
      vfReduce.Denom = viAdd(viMul(v, q0), q1)
      p1 = p0
      q1 = q0
      p0 = vfReduce.Numer
      q0 = vfReduce.Denom

      If viULogB2(q0) > Precision Then
        vfReduce.Numer = p1
        vfReduce.Denom = q1
        Exit Do                     'its big enough so quit
      End If
    End If
    i = i + 1
  Loop
  If vfIsNeg(x) Then vfReduce = vfNeg(vfReduce)
  'this is normalised (by defn, but could be very out of BitSize range in the numerator)

End Function

'Radians to Degrees
Public Function vfRtoD(ByRef x As vFrac) As vFrac
Attribute vfRtoD.VB_Description = "concvert Radians to Degrees"

  vfRtoD.Numer = viMul(x.Numer, viMul(vfcPi.Denom, viFromL(180)))
  vfRtoD.Denom = viMul(x.Denom, vfcPi.Numer)
  Call vfNormalise(vfRtoD)

End Function

'Rectangular to Polar coordinates r= PythAdd(x,y), theta=atan(y,x) in the correct quadrant
Public Sub vfRtoP(ByRef x As vFrac, y As vFrac, ByRef r As vFrac, ByRef theta As vFrac)
Attribute vfRtoP.VB_Description = "return Polar coords for x,y r=PythAdd(x,y); Theta=atan2(y,x)"

  r = vfPythAdd(x, y)
  theta = vfATan2(y, x)   'theta is in the correct quadrant

End Sub

'shift a value ie x = x*2^shift, if shift is negative shift left, positive shift right
Public Function vfShift(ByRef x As vFrac, ByVal shift As Long) As vFrac
Attribute vfShift.VB_Description = "Shift left (*2^k) if k<0 or Shift right (/2^k) if k>0"

  vfShift = x
  If shift <> 0 And Not vfIsZero(vfShift) Then
    If shift < 0 Then vfShift.Numer = viMul(vfShift.Numer, viTwoToThe(viFromL(-shift)))
    If shift > 0 Then vfShift.Numer = viMul(vfShift.Denom, viTwoToThe(viFromL(shift)))
    Call vfNormalise(vfShift)
  End If
  'SIGN is preserved

End Function

Public Function vfSignum(ByRef x As vFrac) As vFrac
Attribute vfSignum.VB_Description = "returns the sign of Vfrac, -1 if <0, 0 if zero, 1 if >0"

  vfSignum.Numer = viSignum(x.Numer)
  vfSignum.Denom = viOne

End Function

'cos(x) = 1(1 - w/2.1(1-w/4.3(1-w/6.5...(1 - w/2k(2k-1)      w=x^2, 0<x<2Pi, for improved convergence properties
'sin(x) = x(1 - w/2.3(1-w/4.5(1-w/6.7...(1 - w/2k(2k+1)      w=x^2, 0<x<2Pi
Public Sub vfSinCos(ByRef x As vFrac, ByRef vSin As vFrac, ByRef vCos As vFrac)
Attribute vfSinCos.VB_Description = "return Sin and Cos of x, -2Pi<x<2Pi"

 Dim i As Long, xm As vFrac, w2 As vFrac, q As Long

  If vfIsZero(x) Then
    vCos = vfOne
    vSin = vfZero
   Else
    Precision = MAXBITS + 4
    xm = vfMod(x, vfc2Pi)                                     'xm -2Pi<x<2Pi (maybe)
    If vfIsNeg(xm) Then xm = vfAdd(xm, vfc2Pi)                'xm now 0<x<2Pi
    If vfCmp(xm, vfcPi) = HIGHER Then xm = vfSub(xm, vfc2Pi)  'now -PI<x<Pi
    If vfIsZero(xm) Then
      vCos = vfOne
      vSin = vfZero
     Else
      w2.Numer = viMul(xm.Numer, xm.Numer)
      w2.Denom = viMul(xm.Denom, xm.Denom)
      vCos = vfOne
      vSin = vfOne
      For i = Precision \ 2 To 2 Step -2    'the 2k factor is in here
        vSin.Denom = viMul3(vSin.Denom, w2.Denom, viFromL(i * (i + 1)))
        vSin.Numer = viSub(vSin.Denom, viMul(w2.Numer, vSin.Numer))

        vCos.Denom = viMul3(vCos.Denom, w2.Denom, viFromL(i * (i - 1)))
        vCos.Numer = viSub(vCos.Denom, viMul(w2.Numer, vCos.Numer))
      Next i
      Precision = MAXBITS
      vSin = vfMul(xm, vSin)
      Call vfNormalise(vCos)
    End If
    Precision = MAXBITS
  End If

End Sub

'Take the Square Root of a vFrac Number >=0, the answer will fit into half as many bits
Public Function vfSqrt(ByRef x As vFrac) As vFrac
Attribute vfSqrt.VB_Description = "return the square Root of x"

 Dim Guess As vFrac, w As vFrac, z As vFrac, s As Integer, q As Long

  If vfIsNeg(x) Then
    vfSqrt = vfZero
    Err.Raise ErrValNegRoot, "RVTvbIM.vfSqrt", ErrStrNegRoot
    Exit Function
  End If

  If vfIsZero(x) Then
    vfSqrt = vfZero
    Exit Function
  End If

  z = x
  If vfCmpOne(z) = LOWER Then
    z = vfInv(z)
    s = 1
  End If

  vfSqrt = vfGuessRoot(z, 2)
  Do                                        'By Newton's root finding method
    Guess = vfSqrt
    w.Numer = viMul(z.Numer, Guess.Denom)
    w.Denom = viMul(z.Denom, Guess.Numer)
    w.Numer = viSub(viMul(w.Numer, Guess.Denom), viMul(w.Denom, Guess.Numer))
    w.Denom = viMul(viShift(w.Denom, -1), Guess.Denom)
    Call vfNormalise(w)
    If vfIsZero(w) Then Exit Do
    vfSqrt = vfAdd(w, Guess)                'Root' = guess + (((x/guess) - guess)/2)
  Loop
  'already normalised
  If s = 1 Then vfSqrt = vfInv(vfSqrt)

End Function

'Signed addition of two vFrac Numbers, Overflow is (nearly) impossible
Public Function vfSub(ByRef x As vFrac, ByRef y As vFrac) As vFrac
Attribute vfSub.VB_Description = "return x-y"

  vfSub.Numer = viSub(viMul(x.Numer, y.Denom), viMul(x.Denom, y.Numer))
  vfSub.Denom = viMul(x.Denom, y.Denom)
  Call vfNormalise(vfSub)

End Function

'tan(x) = sin(x)/cos(x),  if cos(x)=0 then tan(x) = signum(sin(x)) * (2^MAXBITS-1)
Public Function vfTan(ByRef x As vFrac) As vFrac
Attribute vfTan.VB_Description = "return the Tan = Sin/Cos of x,  if Pi/2 or 3PI/2 then return MAXVal"

 Dim vs As vFrac, vc As vFrac

  Call vfSinCos(x, vs, vc)
  If vfIsZero(vc) Then
    vfTan = vfcMaxVal
    If vfIsNeg(vs) Then vfTan = vfNeg(vfTan)
   Else
    vfTan = vfDiv(vs, vc)
  End If

End Function

'return 10
Public Property Get vfTen() As vFrac
Attribute vfTen.VB_Description = "return 10"

  vfTen.Numer = viTen
  vfTen.Denom = viOne

End Property

'return 10^(p/q)
Public Function vfTenToThe(ByRef k As vFrac) As vFrac
Attribute vfTenToThe.VB_Description = "return 10^k"

  vfTenToThe = vfPow(vfTen, k)

End Function

Public Function vfToDecStr(ByRef x As vFrac, Optional ByVal ndec As Long = -1) As String
Attribute vfToDecStr.VB_Description = "convert vFRac to a Decimal String"

 Dim z As vFrac, zm As vInt, zi As vInt, zf As vInt, nd As vFrac, f As vInt

  nd = vfInt(vfMul(vfFromL(Precision + 4), vfFromL(30103, 100000)))
  If ndec < 0 Or ndec > vfToL(nd) Then ndec = vfToL(nd)

  z = vfAbs(x)
  f = viTenToThe(viFromL(ndec))
  zm = viDiv(viAdd(viMul(z.Numer, f), viShift(z.Denom, 1)), z.Denom)  'rounded
  Call viDivMod(zm, f, zi, zf)
  vfToDecStr = viToStr(zi)
  If vfIsNeg(x) Then Mid$(vfToDecStr, 1, 1) = "-"
  If ndec > 0 Then
    vfToDecStr = vfToDecStr & "." & Mid$(viToStr(viAdd(zf, f)), 3)
  End If

End Function

'actually only the lowest 4 bytes
Public Function vfToL(ByRef x As vFrac) As Long
Attribute vfToL.VB_Description = "convert Integer part of VFRa to a Long if small enough"

 Dim i As Long, s As String, HiB As Long, z As vInt

  z = viDiv(x.Numer, x.Denom)
  HiB = UBound(z.s)
  If HiB > 4 Then HiB = 4
  s = "&H"
  For i = HiB To LoB Step -1
    s = s & Right$("00" & Hex$(z.s(i)), 2)
  Next i
  vfToL = CLng(s)
  If viIsNeg(z) Then vfToL = -vfToL     'could overflow if x>2^31

End Function

'Signed conversion and extension from vFrac to String   'as form +nnnnn/ddddddd
Public Function vfToStr(ByRef x As vFrac) As String
Attribute vfToStr.VB_Description = "convert vFRac to a Fractional string"

  On Error GoTo TooLarge
  vfToStr = viToStr(x.Numer) & "/" & Mid$(viToStr(x.Denom), 2)

Exit Function

TooLarge:
  vfToStr = "?"
  Err.Raise ErrValTooBig, "RVTvbIM.vfToStr", ErrStrTooBig
  On Error GoTo 0

End Function

'Signed Test returns -1 if x<0,0 if x=0, +1 if x>0
Public Function vfTst(ByRef x As vFrac) As Long
Attribute vfTst.VB_Description = "Tst a value and return NEGATIVE if <0, ZERO if =0, POSITIVE if >0"

 Dim i As Long

  If viIsNeg(x.Numer) Then
    vfTst = NEGATIVE
   ElseIf viIsZero(x.Numer) Then
    vfTst = ZERO
   Else
    vfTst = POSITIVE
  End If

End Function

'return 2
Public Property Get vfTwo() As vFrac
Attribute vfTwo.VB_Description = "return 2"

  vfTwo.Numer = viTwo
  vfTwo.Denom = viOne

End Property

'return 2^(p/q)
Public Function vfTwoToThe(ByRef k As vFrac) As vFrac
Attribute vfTwoToThe.VB_Description = "return 2^k"

  vfTwoToThe = vfPow(vfTwo, k)

End Function

'unsigned signed compare with One
Public Function vfUCmpOne(ByRef x As vFrac) As Long
Attribute vfUCmpOne.VB_Description = "unsgined compare with one , returns EQUAL if x=1 or x=-1"

  vfUCmpOne = viUCmp(x.Numer, x.Denom)

End Function

'return approximate scale bitsize of vFrac
Private Function vfULogB2(x As vFrac) As vFrac

  vfULogB2 = vfFromL(viULogB2(x.Numer) - viULogB2(x.Denom))

End Function

'UnSigned Test returns 0 if x=0, +1 if x>0
Private Function vfUTst(ByRef x As vFrac) As Long

  If vfIsZero(x) Then vfUTst = ZERO Else vfUTst = POSITIVE

End Function

'return x-1
Public Function vfXMinusOne(ByRef x As vFrac) As vFrac
Attribute vfXMinusOne.VB_Description = "return x-1"

  vfXMinusOne.Numer = viSub(x.Numer, x.Denom)
  vfXMinusOne.Denom = x.Denom
  Call vfNormalise(vfXMinusOne)

End Function

'x+1
Public Function vfXPlusOne(ByRef x As vFrac) As vFrac
Attribute vfXPlusOne.VB_Description = "return x+1"

  vfXPlusOne.Numer = viAdd(x.Numer, x.Denom)
  vfXPlusOne.Denom = x.Denom
  Call vfNormalise(vfXPlusOne)

End Function

'Zero vFrac  (makes it three zero bytes)
Public Property Get vfZero() As vFrac
Attribute vfZero.VB_Description = "return 0"

  vfZero.Numer = viZero                'the smallest vFrac possible 0,0,0/1,1,0
  vfZero.Denom = viOne

End Property

'================================================================================================================
'====================================== VINTS ===================================================================
'================================================================================================================

'Absolute value
Public Function viAbs(ByRef x As vInt) As vInt
Attribute viAbs.VB_Description = "return absolute value of x"

  If viIsNeg(x) Then
    viAbs = viNeg(x)
   Else
    viAbs = x
  End If
  'already normalised

End Function

'Signed addition of two vInt Numbers
Public Function viAdd(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viAdd.VB_Description = "return x+y"

 Dim sx As Integer, sy As Integer

  sx = viTst(x)
  sy = viTst(y)
  If sx = NEGATIVE And sy = NEGATIVE Then
    viAdd = viNeg(viUAdd(viNeg(x), viNeg(y)))
   ElseIf sx = NEGATIVE Then
    viAdd = viSub(y, viNeg(x))
   ElseIf sy = NEGATIVE Then
    viAdd = viSub(x, viNeg(y))
   Else
    viAdd = viUAdd(x, y)
  End If
  'already normalised

End Function

'Taking account of sign --- returns -1 if x<y,0 if x=y, +1 if x>y
'x and y must be Normalised for this to work
Public Function viCmp(ByRef x As vInt, ByRef y As vInt) As Long
Attribute viCmp.VB_Description = "compare x with y, HIGHER if x>y, EQUAL if x=y, LOWER if x<y"

 Dim i As Long

  If viIsPos(x) And viIsNeg(y) Then
    viCmp = HIGHER                            'x+ve, y-ve
   ElseIf viIsNeg(x) And viIsPos(y) Then
    viCmp = LOWER                             'x-ve, y +ve
   Else                                     'either both +ve or both -ve
    If UBound(x.s) > UBound(y.s) Then         'x is longer
      viCmp = HIGHER
     ElseIf UBound(x.s) < UBound(y.s) Then    'y is longer
      viCmp = LOWER
     Else                                     'same length
      For i = UBound(x.s) To LoB Step -1
        If x.s(i) > y.s(i) Then
          viCmp = HIGHER
          Exit For
         ElseIf x.s(i) < y.s(i) Then
          viCmp = LOWER
          Exit For
        End If
      Next i
    End If
    If viIsNeg(x) Then viCmp = -viCmp    'if both negative, reverse sense
  End If
  'EQUAL if we get to here unscathed

End Function

'signed compare with One
Public Function viCmpOne(ByRef x As vInt) As Long
Attribute viCmpOne.VB_Description = "compare x and 1, HIGHER if x>1, EQUAL if x=1, LOWER if x<1"

  If viIsNeg(x) Or viIsZero(x) Then
    viCmpOne = LOWER
   ElseIf UBound(x.s) = 2 Then
    If x.s(LoB) = 1 Then
      If x.s(LoB + 1) = 0 Then viCmpOne = EQUAL
     Else
      viCmpOne = HIGHER
    End If
   Else
    viCmpOne = HIGHER
  End If

End Function

'Duplicates src vInt
Public Function viCopy(ByRef src As vInt) As vInt
Attribute viCopy.VB_Description = "duplicate src into dest"

 ' Dim i As Long

  Call viNormalise(src)
  viCopy = src

End Function

'internal digit conversion
Private Sub viCVDigits(ByRef d() As Byte, ByVal n As Long, ByRef x As vInt)

 Dim i As Long, j As Long, c As Long
 Dim q As vInt, r As vInt

  j = n
  q = x
  Do
    Call viUDivMod(q, viTen, q, r)
    d(j) = r.s(LoB) + 48
    j = j - 1
  Loop Until viTst(q) = ZERO
  j = j + 1

  If j > 1 Then
    i = 1
    For j = j To n
      d(i) = d(j)
      i = i + 1
    Next j
    ReDim Preserve d(0 To i - 1)
  End If

End Sub

'Signed Division - result = arithmetic quotient
Public Function viDiv(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viDiv.VB_Description = "return x/y, y<>0"

 Dim qq As vInt, rr As vInt

  Call viDivMod(x, y, qq, rr)
  viDiv = qq

End Function

'Signed Division and Modulo - q = arithmetic quotient, r= remainder has sign of divisor
Public Sub viDivMod(ByRef x As vInt, ByRef y As vInt, ByRef q As vInt, ByRef r As vInt)
Attribute viDivMod.VB_Description = "return quotient and remainder. remainder takes sign of divisor"

  Call viUDivMod(viAbs(x), viAbs(y), q, r)
  If viIsNeg(x) <> viIsNeg(y) Then q = viNeg(q)
  If viIsNeg(y) Then r = viNeg(r)

End Sub

'Shift the value from bits j.. to bits k.. (z will be made large enough to take the shifted value)
Private Function viDoShift(ByRef x As vInt, ByVal j As Long, ByVal k As Long) As vInt

 Dim i As Long, z As vInt, m1 As Byte, m2 As Byte

  If k < 0 Then
    viDoShift = viZero
    Exit Function
  End If

  i = j Mod 8
  m1 = 1
  For i = 0 To i - 1
    m1 = m1 + m1
  Next i

  i = k Mod 8
  m2 = 1
  For i = 0 To i - 1
    m2 = m2 + m2
  Next i

  ReDim z.s(SiB To 2 * ((k + 31) \ 16))
  z.s(SiB) = x.s(SiB)       'preserve sign of x

  For i = k To 0 Step -1    'bitwise
    If (x.s(LoB + j \ 8) And m1) <> 0 Then z.s(LoB + i \ 8) = z.s(LoB + i \ 8) Or m2
    j = j - 1
    If j < 0 Then Exit For
    m1 = m1 \ 2: If m1 = 0 Then m1 = 128
    m2 = m2 \ 2: If m2 = 0 Then m2 = 128
  Next i
  viDoShift = viCopy(z)

End Function

'Factorial n
Public Function viFact(ByRef n As vInt) As vInt
Attribute viFact.VB_Description = "return Factorial(x)  ie. x!"

 Dim y As vInt

  If viIsZero(n) Then
    viFact = viOne
    Exit Function
  End If

  y = viAbs(n)
  viFact = viOne
  Do While viCmp(y, viOne) <> EQUAL
    viFact = viMul(viFact, y)
    y = viSub(y, viOne)
  Loop
  Call viNormalise(viFact)

End Function

'Format a VInt into a string of width , width <0 then left aligned
Public Function viFmt(ByRef x As vInt, ByVal Width As Long) As String
Attribute viFmt.VB_Description = "Format a vint as a string"

  viFmt = viToStr(x)
  If Len(viFmt) < Abs(Width) Then
    If Width < 0 Then
      viFmt = viFmt & String$(Abs(Width) - Len(viFmt), 32)
     Else
      viFmt = String$(Width - Len(viFmt), 32) & viFmt
    End If
  End If

End Function

'Signed conversion and extension from Long to vInt64
Public Function viFromL(ByVal LongVal As Long) As vInt
Attribute viFromL.VB_Description = "create a Vint from a Long value"

 Dim s As String, i As Long

  ReDim viFromL.s(SiB To 4)                   'the size of a long
  If LongVal < 0 Then
    viFromL.s(SiB) = SiBNeg
    LongVal = -LongVal        'can overflow -2^31
  End If
  For i = 1 To 4
    viFromL.s(i) = LongVal And BASEMAX
    LongVal = LongVal \ BASE
  Next i
  '  s = Right$("00000000" & Hex$(LongVal), 8)
  '  viFromL.s(4) = Val("&H" & Mid$(s, 1, 2))
  '  viFromL.s(3) = Val("&H" & Mid$(s, 3, 2))
  '  viFromL.s(2) = Val("&H" & Mid$(s, 5, 2))
  '  viFromL.s(1) = Val("&H" & Mid$(s, 7, 2))
  Call viNormalise(viFromL)

End Function

'Signed conversion and extension from String to vInt
Public Function viFromStr(ByRef x As String) As vInt
Attribute viFromStr.VB_Description = "creae a Vint from a string +-nnnnnn"

 Dim s As Integer, i As Long, j As Long
 Dim d() As Byte, c As Long, p As Long, HiB As Long

  On Error GoTo TooLarge
  s = POSITIVE
  d = StrConv(x, vbFromUnicode)

  HiB = ((UBound(d) + 1) * 1000) \ 2408 + 1     'an estimate of the number of bytes required for string
  If (HiB And 1) = 1 Then HiB = HiB + 1
  ReDim viFromStr.s(SiB To HiB)

  For i = 0 To UBound(d)
    c = d(i)
    If c = Asc("-") Then
      s = -s
     Else
      If c >= 48 And c <= 57 Then     '0..9
        c = c - 48
        For j = LoB To HiB                ' multiply q = p * q
          c = 10 * CLng(viFromStr.s(j)) + c
          viFromStr.s(j) = c And BASEMAX
          c = c \ BASE
        Next j

        If c > 0 Then
          Err.Raise ErrValOverflow, "RVTvbIM.viFromStr", ErrStrOverflow
          Exit Function
        End If
      End If
    End If
  Next i
  If s = NEGATIVE Then viFromStr = viNeg(viFromStr)
  Call viNormalise(viFromStr)

Exit Function

TooLarge:
  viFromStr = viZero
  Err.Raise ErrValTooBig, "RVTvbIM.FromStr", ErrStrTooBig
  On Error GoTo 0

End Function

'return the Greatest Common Denominator (Factor)
Public Function viGCD(ByRef x As vInt, ByRef y As vInt) As vInt   'Euclids algorithm
Attribute viGCD.VB_Description = "Greatest common Divisor of x and y"

 Dim r As vInt, d As vInt

  viGCD = viAbs(x)
  r = viAbs(y)

  Do While Not viIsZero(r)
    d = r
    r = viMod(viGCD, d)
    viGCD = d
  Loop
  If viIsNeg(x) And viIsNeg(y) Then viGCD = viNeg(viGCD)
  'already normalised

End Function

'A first guess of the value of the root (the value shifted right by 1/k its nr of bits)
Private Function viGuessRoot(ByRef x As vInt, ByVal k As Long) As vInt

 Dim n As Long

  n = viULogB2(x)
  viGuessRoot = viDoShift(x, n, n \ k)
  viGuessRoot.s(LoB) = viGuessRoot.s(LoB) Or 1    ' just in case!!!!
  'normalised in DoShift

End Function

'test if even
Private Function viIsEven(ByRef x As vInt) As Boolean

  If (x.s(LoB) And 1) = 0 Then viIsEven = True

End Function

'test if negative
Private Function viIsNeg(ByRef x As vInt) As Boolean

  If x.s(SiB) = SiBNeg Then viIsNeg = True

End Function

'test if odd
Private Function viIsOdd(ByRef x As vInt) As Boolean

  If (x.s(LoB) And 1) = 1 Then viIsOdd = True

End Function

'tests if a vInt is Positive
Private Function viIsPos(ByRef x As vInt) As Boolean

  If x.s(SiB) <> SiBNeg Then viIsPos = True

End Function

'tests if a vInt is Zero
Private Function viIsZero(ByRef x As vInt) As Boolean

  If x.s(SiB) = SiBZero Then viIsZero = True

End Function

'Take the Kth Root of a vInt Number >=0, the answer will fit into 1/k as many bits
Public Function viKRoot(ByRef x As vInt, ByRef k As vInt) As vInt
Attribute viKRoot.VB_Description = "return kth root of x (nearest only)"

 Dim Guess As vInt, w As vInt, km1 As vInt, z As vInt

  If viIsNeg(x) And viIsEven(k) Then
    viKRoot = viZero
    Err.Raise ErrValERoot, "RVTvbIM.viKRoot", ErrStrERoot
    Exit Function
  End If

  If viIsZero(x) Or viIsNeg(k) Then
    viKRoot = viZero
    Exit Function
  End If

  z = viShift(viAbs(x), -2 * viToL(k))
  km1 = viXMinusOne(k)
  viKRoot = viGuessRoot(z, viToL(k))
  Do                                        'By Newton's root finding method
    Guess = viKRoot
    w = viDiv(viSub(viDiv(z, viPow(Guess, km1)), Guess), k)
    If viIsZero(w) Then Exit Do
    viKRoot = viAdd(w, Guess)               'Root' = guess + (((x/guess^k-1) - guess)/k)
  Loop
  viKRoot = viShift(viKRoot, 2)
  If viIsNeg(x) Then viKRoot = viNeg(viKRoot)
  'already normalised

End Function

'return Lowest common Multiple
Public Function viLCM(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viLCM.VB_Description = "Lowest Common Multiple of x and y"

  viLCM = viDiv(viMul(x, y), viGCD(x, y))
  'alreday normalised

End Function

'Signed Modulo - result = remainder has sign of divisor
Public Function viMod(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viMod.VB_Description = "return remainder of x/y, has sign of y"

 Dim qq As vInt, rr As vInt

  Call viDivMod(x, y, qq, rr)
  viMod = rr
  'already normalised

End Function

'randomize the Rand Generator with the clock
Public Sub vimRandomize()
Attribute vimRandomize.VB_Description = "randomise the Random Number Generator"

  Call Randomize

End Sub

'seed the random number generator
Public Sub vimRandSeed(ByVal Seed1 As Long, ByVal Seed2 As Long)
Attribute vimRandSeed.VB_Description = "Seed the random Number Generator"

  Call SeedRand(Seed1, Seed2)

End Sub

'Signed multiply of two vInt Numbers
Public Function viMul(ByRef x As vInt, ByRef y As vInt) As vInt

  viMul = viUMul(viAbs(x), viAbs(y))
  If viIsNeg(x) <> viIsNeg(y) Then viMul = viNeg(viMul)
  'already normalised

End Function

'Signed multiply of three vInt Numbers
Public Function viMul3(ByRef x As vInt, ByRef y As vInt, ByRef z As vInt) As vInt

  viMul3 = viMul(viMul(x, y), z)
  'already normalised

End Function

'nCr - nr of Combinations of n taken r at a time
Public Function viNCr(ByRef n As vInt, ByRef r As vInt) As vInt
Attribute viNCr.VB_Description = "Return the number of Combinations of n items r at a time "

  viNCr = viUDiv(viFact(n), viUMul(viFact(viSub(n, r)), viFact(r)))
  'already normalised

End Function

'Negate a vInt  x=-x
Public Function viNeg(ByRef x As vInt) As vInt
Attribute viNeg.VB_Description = "negate x"

  viNeg = x
  If viIsNeg(viNeg) Then
    viNeg.s(SiB) = SiBPos
   ElseIf Not viIsZero(viNeg) Then
    viNeg.s(SiB) = SiBNeg
  End If
  'already normalised

End Function

'remove leading zeroes (in pairs) from a vInt, and verify SignByte
Private Sub viNormalise(ByRef v As vInt)

 Dim i As Long, j As Long, HiB As Long

  HiB = UBound(v.s)
  If HiB > 2 Then    'cant get smaller than this SiB=0, LoB=1, HiB=2
    j = HiB
    For i = HiB To LoB + 2 Step -2
      If v.s(i) <> 0 Or v.s(i - 1) <> 0 Then Exit For
      j = j - 2
    Next i
    If j <> HiB Then ReDim Preserve v.s(SiB To j)
  End If

  'make sure zero is correctly identified
  If v.s(SiB) = SiBZero Then v.s(SiB) = SiBPos            'assume positive
  If UBound(v.s) = 2 Then                                 'unless proved otherwise
    If v.s(1) = 0 And v.s(2) = 0 Then v.s(SiB) = SiBZero
  End If

End Sub

'nPr - nr of Permutations of n taken r at a time
Public Function viNPr(ByRef n As vInt, ByRef r As vInt) As vInt
Attribute viNPr.VB_Description = "return the number of Permutations of n items taken r at a time"

  viNPr = viUDiv(viFact(n), viFact(viSub(n, r)))
  'already normalised

End Function

'the constant 1
Public Property Get viOne() As vInt
Attribute viOne.VB_Description = "return 1"

  ReDim viOne.s(SiB To 2)
  viOne.s(SiB) = SiBPos
  viOne.s(LoB) = 1

End Property

'Take the power x^y (y>0)
Public Function viPow(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viPow.VB_Description = "return x^y"

 Dim w As vInt, v As vInt, u As vInt, sx As Integer

  If viIsZero(y) Or viCmp(x, viOne) = EQUAL Then    'x^0 = 1, 1^y = 1
    viPow = viOne
    Exit Function
  End If

  If viIsNeg(y) Or viIsZero(x) Then                 '1/x^y=0, 0^x (x<>0) = 0
    viPow = viZero
    Exit Function
  End If

  If viCmp(x, viNeg(viOne)) = EQUAL Then
    viPow = viOne
    If viIsOdd(y) Then viPow = viNeg(viPow)
    Exit Function
  End If

  w = viOne
  v = viAbs(x)
  u = y
  Do
    If viIsOdd(u) Then w = viUMul(w, v)
    v = viUMul(v, v)
    u = viShift(u, 1)     '/2
  Loop Until viIsZero(u)
  If viIsNeg(x) And viIsOdd(y) Then w = viNeg(w)
  viPow = viCopy(w)
  'already normalised

End Function

'return a random number of maximally Abs(MaxVal)
Public Function viRand(ByRef MaxVal As vInt) As vInt
Attribute viRand.VB_Description = "return a Random Number in Range of x"

  viRand = viUDiv(viUMul(viFromL(UniformRand), MaxVal), viFromL(2147483563))
  'already normalised

End Function

'shift a value ie x = x*2^shift, if shift is negative shift left, positive shift right
Public Function viShift(ByRef x As vInt, ByVal shift As Long) As vInt
Attribute viShift.VB_Description = "Shift left (*2^k) if k<0 or Shift right (/2^k) if k>0"

 Dim j As Long, k As Long

  viShift = x
  If shift <> 0 And Not viIsZero(viShift) Then
    j = viULogB2(x)   'first bit x
    k = j - shift
    If k < 0 Then
      viShift = viZero
     Else
      viShift = viDoShift(x, j, k)
    End If
  End If
  'SIGN is preserved
  'already normalised

End Function

Public Function viSignum(ByRef x As vInt) As vInt
Attribute viSignum.VB_Description = "returns the sign of Vfrac, -1 if <0, 0 if zero, 1 if >0"

  If viIsNeg(x) Then
    viSignum = viNeg(viOne)
   ElseIf viIsZero(x) Then
    viSignum = viZero
   Else
    viSignum = viOne
  End If

End Function

'Take the Square Root of a vInt Number >=0, the answer will fit into half as many bits
Public Function viSqrt(ByRef x As vInt) As vInt
Attribute viSqrt.VB_Description = "return the square Root of x (closest only)"

 Dim Guess As vInt, w As vInt

  If viIsNeg(x) Then
    viSqrt = viZero
    Err.Raise ErrValNegRoot, "RVTvbIM.viSqrt", ErrStrNegRoot
    Exit Function
  End If

  If viIsZero(x) Then
    viSqrt = viZero
    Exit Function
  End If

  viSqrt = viGuessRoot(x, 2)
  Do                                        'By Newton's root finding method
    Guess = viSqrt
    w = viUDiv(x, Guess)
    w = viSub(w, Guess)
    w = viShift(w, 1)       '/ 2
    viSqrt = viAdd(w, Guess)                'Root' = guess + (((x/guess) - guess)/2)
  Loop Until viUCmp(viSqrt, Guess) = ZERO
  'already normalised

End Function

'Signed addition of two vInt Numbers, Overflow is (nearly) impossible
Public Function viSub(ByRef x As vInt, ByRef y As vInt) As vInt
Attribute viSub.VB_Description = "return x-y"

 Dim sx As Integer, sy As Integer

  sx = viTst(x)
  sy = viTst(y)
  If sx = NEGATIVE And sy = NEGATIVE Then
    viSub = viNeg(viSub(viNeg(x), viNeg(y)))    'recursive call!!
   ElseIf sx = NEGATIVE Then
    viSub = viNeg(viUAdd(viNeg(x), y))
   ElseIf sy = NEGATIVE Then
    viSub = viUAdd(x, viNeg(y))
   Else
    If viUCmp(x, y) = HIGHER Then
      viSub = viUSub(x, y)
     Else
      viSub = viNeg(viUSub(y, x))
    End If
  End If
  'already normalised

End Function

'return 10
Public Property Get viTen() As vInt
Attribute viTen.VB_Description = "return 10"

  ReDim viTen.s(SiB To 2)
  viTen.s(SiB) = SiBPos
  viTen.s(LoB) = 10

End Property

'return 10^k
Public Function viTenToThe(ByRef k As vInt) As vInt
Attribute viTenToThe.VB_Description = "return 10^k"

  viTenToThe = viPow(viTen, k)
  'already normalised

End Function

'actually only the lowest 4 bytes
Public Function viToL(ByRef x As vInt) As Long
Attribute viToL.VB_Description = "return a Long from a Vint if its small enough"

 Dim i As Long, s As String, HiB As Long

  HiB = UBound(x.s)
  If HiB > 4 Then HiB = 4
  s = "&H"
  For i = HiB To LoB Step -1
    s = s & Right$("00" & Hex$(x.s(i)), 2)
  Next i
  viToL = CLng(s)
  If viIsNeg(x) Then viToL = -viToL     'could overflow if x>2^31

End Function

'Signed conversion and extension from vInt to String
Public Function viToStr(ByRef x As vInt) As String
Attribute viToStr.VB_Description = "convert a vint to a string"

 Dim d() As Byte, n As Long, z As vInt

  On Error GoTo TooLarge
  n = ((1 + UBound(x.s)) * 2408) \ 1000 + 1   'the approx number of digits to be expected base 10
  ReDim d(0 To n)
  If viIsNeg(x) Then
    d(0) = Asc("-")
   Else
    d(0) = 32
  End If
  Call viCVDigits(d(), n, viAbs(x))
  viToStr = StrConv(d, vbUnicode)

Exit Function

TooLarge:
  viToStr = "?"
  Err.Raise ErrValTooBig, "RVTvbIM.ToStr", ErrStrTooBig
  On Error GoTo 0

End Function

'Signed Test returns -1 if x<0,0 if x=0, +1 if x>0
Public Function viTst(ByRef x As vInt) As Long
Attribute viTst.VB_Description = "Tst a value and return NEGATIVE if <0, ZERO if =0, POSITIVE if >0"

 Dim i As Long

  If viIsNeg(x) Then
    viTst = NEGATIVE
   ElseIf viIsZero(x) Then
    viTst = ZERO
   Else
    viTst = POSITIVE
  End If

End Function

'return 2
Public Property Get viTwo() As vInt
Attribute viTwo.VB_Description = "return 2"

  ReDim viTwo.s(SiB To 2)
  viTwo.s(SiB) = SiBPos
  viTwo.s(LoB) = 2

End Property

'return 2^k
Public Function viTwoToThe(ByRef k As vInt) As vInt
Attribute viTwoToThe.VB_Description = "return 2^k"

  If viIsNeg(k) Then
    viTwoToThe = viZero
   Else
    viTwoToThe = viDoShift(viTwo, 1, viToL(k))
  End If
  'already normalised

End Function

'Unsigned addition of two vInt Numbers
Private Function viUAdd(ByRef x As vInt, ByRef y As vInt) As vInt

 Dim i As Long, j As Long, c As Long
 Dim HiBx As Long, HiBy As Long, HiBq As Long

  HiBx = UBound(x.s)
  HiBy = UBound(y.s)
  If HiBx > HiBy Then HiBq = HiBx + 2 Else HiBq = HiBy + 2

  ReDim viUAdd.s(SiB To HiBq)
  c = 0
  For i = LoB To HiBq                              ' z(HiB..LoB) = x(HiB..LoB) + y(HiB..LoB)
    If i <= HiBx Then c = c + CLng(x.s(i))
    If i <= HiBy Then c = c + CLng(y.s(i))
    viUAdd.s(i) = c And BASEMAX
    c = c \ BASE
  Next i
  Call viNormalise(viUAdd)

End Function

'Unsigned Comparison ---- returns -1 if x<y,0 if x=y, +1 if x>y
Private Function viUCmp(ByRef x As vInt, ByRef y As vInt) As Long

 Dim i As Long

  If UBound(x.s) > UBound(y.s) Then
    viUCmp = HIGHER
    Exit Function
   ElseIf UBound(x.s) < UBound(y.s) Then
    viUCmp = LOWER
    Exit Function
   Else                                    'same length
    For i = UBound(x.s) To LoB Step -1
      If x.s(i) > y.s(i) Then
        viUCmp = HIGHER
        Exit Function
       ElseIf x.s(i) < y.s(i) Then
        viUCmp = LOWER
        Exit Function
      End If
    Next i
  End If
  'EQUAL if we got here

End Function

'Unsigned Comparison with an arbirary long ---- returns -1 if x<y,0 if x=y, +1 if x>y
Private Function viUCmpL(ByRef x As vInt, ByVal LongVal As Long) As Long

  viUCmpL = viUCmp(x, viFromL(LongVal))

End Function

'unsigned compare with one
Public Function viUCmpOne(ByRef x As vInt) As Long
Attribute viUCmpOne.VB_Description = "unsgined compare with one , returns EQUAL if x=1 or x=-1"

  If viIsZero(x) Then
    viUCmpOne = LOWER
   ElseIf UBound(x.s) = 2 Then
    viUCmpOne = HIGHER
    If x.s(LoB) = 1 Then
      If x.s(LoB + 1) = 0 Then
        viUCmpOne = EQUAL
      End If
    End If
   Else
    viUCmpOne = HIGHER
  End If

End Function

'Unsigned Division = quotient
Private Function viUDiv(ByRef x As vInt, ByRef y As vInt) As vInt

 Dim rr As vInt

  Call viUDivMod(x, y, viUDiv, rr)

End Function

'Unsigned Division - q = quotient, r= remainder
Private Sub viUDivMod(ByRef x As vInt, ByRef y As vInt, ByRef q As vInt, ByRef r As vInt)

 Dim i As Long, p As Long, c As Long, HiB As Long
 Dim bx As Integer, by As Integer, bq As Integer        'offset of byte
 Dim px As Integer, py As Integer, pz As Integer        'tmp pointers
 Dim tx As vInt, ty As vInt, tz As vInt, tq As vInt         'tmp variables

  tx = x
  ty = y
  HiB = UBound(tx.s)
  If UBound(ty.s) > HiB Then HiB = UBound(ty.s)

  ReDim Preserve tx.s(SiB To HiB + 4)
  ReDim Preserve ty.s(SiB To HiB + 4)
  ReDim tq.s(SiB To HiB + 4)
  ReDim tz.s(SiB To HiB + 4)

  '=================== tmpvars now ready

  If viUTst(ty) = ZERO Then
    Err.Raise ErrValDivZero, "RVTvbIM.viDiv,DivMod,Mod", ErrStrDivZero
    Exit Sub
  End If

  If viUCmp(ty, viOne) = EQUAL Then     'x div 1 = x, r=0
    q = viCopy(tx)
    r = viZero
    Exit Sub
  End If

  c = viUCmp(tx, ty)
  If c = LOWER Then      'if x<y q=0,r=x
    q = viZero
    r = viCopy(tx)
    Exit Sub
  End If

  If c = EQUAL Then      'if x=y q=1,r=0
    q = viOne
    r = viZero
    Exit Sub
  End If

  'we now know that x>y so q>=0, r>=0
  'setup where to start

  bx = HiB                                  ' at most significant digit of dividend
  Do While (tx.s(bx) = 0)
    bx = bx - 1                             ' to first nonzero digit
    If bx < LoB Then Exit Do
  Loop

  by = HiB                                  ' at msd of divisor
  Do While (ty.s(by) = 0)
    by = by - 1                             ' to first nonzero digit
    If by < LoB Then Exit Do
  Loop

  If bx < LoB Or by < LoB Then by = 1 / 0   '>>>>>>>>>>>>>>> this should NEVER happen

  bq = bx - by + 1                            ' at msd of quotient
  Do While (bx >= LoB And bq >= LoB)
    c = 0
    px = bx
    py = by
    For i = LoB To by                       ' find the next quotient digit
      If (ty.s(py) > tx.s(px)) Then
        c = HIGHER
        bq = bx - by
        Exit For
       ElseIf (ty.s(py) < tx.s(px)) Then
        c = LOWER
        bq = bx - by + 1
        Exit For
      End If
      py = py - 1
      px = px - 1
    Next i                                  ' bq is where the next digit of quotient will be

    If bq < LoB Then Exit Do

    'now ready to go   bx=first digit dividend, by=first digit divisor, bq=first digit quotient
    If c = EQUAL Then
      p = 1
     ElseIf bx = LoB Or c < EQUAL Then
      p = CLng(tx.s(bx)) / CLng(ty.s(by))
     Else
      p = (BASE * CLng(tx.s(bx)) + tx.s(bx - 1)) / ty.s(by) ' first two digits dividend/1 digit divisor
    End If
    If (p > BASEMAX) Then p = BASEMAX                       ' the guess of quotient digit

    ReDim tz.s(SiB To HiB + 4) 'clear accumulator

    c = 0
    pz = LoB
    py = LoB
    For i = LoB To by                                     ' multiply z(by+1..LoB) = p * y(by..LoB)
      c = p * CLng(ty.s(py)) + c
      tz.s(pz) = c And BASEMAX
      c = c \ BASE
      py = py + 1
      pz = pz + 1
    Next i
    tz.s(pz) = c

    c = 0
    pz = LoB
    px = bq
    For i = bq To bq + by + 1                             ' z(by+1..bq) = z(by+1..bq) - x(by+1..bq)
      c = BASE + CLng(tx.s(px)) - CLng(tz.s(pz)) - c
      tz.s(pz) = c And BASEMAX
      If (c - tz.s(pz)) = 0 Then c = 1 Else c = 0
      px = px + 1
      pz = pz + 1
    Next i
    tz.s(pz) = c         'the borrow digit

    Do While (c <> 0)     'we are left with a borrow so we have overdone it, so decrement p, increment difference
      p = p - 1           'lower quotient digit by 1

      c = 0
      pz = LoB
      py = LoB
      For i = bq To bq + by + 1                           ' z(by+1..bq) = z(by+1..bq) + y(by+1..LoB)
        c = CLng(tz.s(pz)) + CLng(ty.s(py)) + c
        tz.s(pz) = c And BASEMAX
        c = c \ BASE
        py = py + 1
        pz = pz + 1
      Next i
      c = tz.s(pz) - c
      tz.s(pz) = c           'the borrow digit now
    Loop

    ' remaining digits of x are now known (copy back to x from z)
    pz = LoB
    px = bq
    For i = bq To bq + by + 1               ' x(HIB64..bq)= z(HIB64..bq)
      tx.s(px) = tz.s(pz)
      px = px + 1
      pz = pz + 1
    Next i

    Do While (tx.s(bx) = 0)
      bx = bx - 1                            'to first nonzero digit
      If bx < LoB Then Exit Do
    Loop
    tq.s(bq) = p                             'the final value of q
    bq = bq - 1                              'next quotient pos
  Loop

  q = viCopy(tq)
  r = viCopy(tx)

End Sub

'Return the first bit set on a Number (sign ignored)
Private Function viULogB2(ByRef x As vInt) As Long

 Dim i As Long, q As Long, j As Long, n As Long

  viULogB2 = -1
  For i = UBound(x.s) To LoB Step -1
    If x.s(i) <> 0 Then
      n = 8 * i - 1               '63,55,47,39,31,23,15,7
      q = x.s(i): j = 128
      Do While (q And j) = 0
        n = n - 1
        j = j \ 2
      Loop
      viULogB2 = n
      Exit Function
    End If
  Next i

End Function

'Unsigned Modulo = remainder
Private Function viUMod(ByRef x As vInt, ByRef y As vInt) As vInt

 Dim qq As vInt

  Call viUDivMod(x, y, qq, viUMod)

End Function

'Unsigned multiply of two vInt Numbers
Private Function viUMul(ByRef x As vInt, ByRef y As vInt) As vInt

 Dim i As Long, j As Long, c As Long
 Dim HiBx As Long, HiBy As Long

  HiBx = UBound(x.s)
  HiBy = UBound(y.s)
  ReDim viUMul.s(SiB To HiBx + HiBy)    '+2

  For j = LoB To HiBy
    c = 0
    For i = LoB To HiBx
      c = CLng(x.s(i)) * CLng(y.s(j)) + CLng(viUMul.s(i + j - 1)) + c
      viUMul.s(i + j - 1) = c And BASEMAX
      c = c \ BASE
    Next i
    viUMul.s(i + j - 1) = c
  Next j
  Call viNormalise(viUMul)

End Function

'Unsigned subtraction of two vInt Numbers
Private Function viUSub(ByRef x As vInt, ByRef y As vInt) As vInt

 Dim i As Long, j As Long, c As Long
 Dim HiBx As Long, HiBy As Long, HiBq As Long

  HiBx = UBound(x.s)
  HiBy = UBound(y.s)
  If HiBx > HiBy Then HiBq = HiBx Else HiBq = HiBy

  ReDim viUSub.s(SiB To HiBq)

  c = 0
  For i = LoB To HiBq                              ' z(HIB64..LoB) = x(HIB64..LoB) - y(HIB64..LoB)
    c = BASE - c
    If i <= HiBx Then c = c + CLng(x.s(i))
    If i <= HiBy Then c = c - CLng(y.s(i))
    viUSub.s(i) = c And BASEMAX
    If (c - viUSub.s(i)) = 0 Then c = 1 Else c = 0
  Next i
  Call viNormalise(viUSub)

End Function

'Signed conversion and extension from vInt to String
Private Function viUToStr(ByRef x As vInt) As String

 Dim d() As Byte, n As Long

  On Error GoTo TooLarge
  n = ((1 + UBound(x.s)) * 2408) \ 1000 + 1   'the approx number of digits to be expected base 10
  ReDim d(0 To n)
  d(0) = 32
  Call viCVDigits(d(), n, viAbs(x))
  viUToStr = StrConv(d, vbUnicode)

Exit Function

TooLarge:
  viUToStr = "?"
  Err.Raise ErrValTooBig, "RVTvbIM.ToStr", ErrStrTooBig
  On Error GoTo 0

End Function

'UnSigned Test returns 0 if x=0, +1 if x>0
Private Function viUTst(ByRef x As vInt) As Long

  If viIsZero(x) Then viUTst = ZERO Else viUTst = POSITIVE

End Function

'x-1
Public Function viXMinusOne(ByRef x As vInt) As vInt
Attribute viXMinusOne.VB_Description = "return x-1"

  viXMinusOne = viSub(x, viOne)

End Function

'x+1
Public Function viXPlusOne(ByRef x As vInt) As vInt
Attribute viXPlusOne.VB_Description = "return x+1"

  viXPlusOne = viAdd(x, viOne)

End Function

'Zero vInt  (makes it three zero bytes)
Public Property Get viZero() As vInt
Attribute viZero.VB_Description = "return 0"

  ReDim viZero.s(SiB To LoB + 1)                   'the smallest vInt possible 0,0,0

End Property

':) Ulli's VB Code Formatter V2.10.8 (01-Jun-02 21:04:29) 98 + 2314 = 2412 Lines
