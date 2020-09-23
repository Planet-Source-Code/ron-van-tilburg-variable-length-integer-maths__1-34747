Attribute VB_Name = "IMRand"
Option Explicit

'Uniform RandomNumberGenerator
'Copyright ©2002 R. van Tilburg
'Public rights for educational use are given. Commercial rights retained by author
'===============================================================================================================

Private mvarSeed1 As Long    'for Uniform
Private mvarSeed2 As Long    'for Uniform

'===================================== METHODS =========================================================================

'See Numerical Recipes in C 2nd Ed p282

Public Function UniformRand() As Double
  
  Const IM1 As Long = 2147483563
  Const IM2 As Long = 2147483399
  Const IA1 As Long = 40014
  Const IA2 As Long = 40692
  Const IQ1 As Long = 53668
  Const IQ2 As Long = 52774
  Const IR1 As Long = 12211
  Const IR2 As Long = 3791
'  Const MXR As Double = (1# - 2.22044604925031E-16)       '(1.0-EPS(double))
'  Const SCF As Double = MXR / 2147483563#                 '/IM1
  
  Dim k As Long
  
  k = mvarSeed1 \ IQ1
  mvarSeed1 = IA1 * (mvarSeed1 - k * IQ1) - k * IR1
  If mvarSeed1 < 0 Then mvarSeed1 = mvarSeed1 + IM1
  
  k = mvarSeed2 \ IQ2
  mvarSeed2 = IA2 * (mvarSeed2 - k * IQ2) - k * IR2
  If mvarSeed2 < 0 Then mvarSeed2 = mvarSeed2 + IM2
  
  k = mvarSeed1 - mvarSeed2
  If k < 1 Then k = k + IM1 - 1
  
  UniformRand = k '* SCF
'  If UniformRand > MXR Then UniformRand = MXR           'smallest 0.0<z<1-eps(double)
End Function

Public Sub SeedRand(ByVal Seed1 As Long, ByVal Seed2 As Long)
  
  Const IM1 As Long = 2147483563
  Const IM2 As Long = 2147483399
  
  mvarSeed1 = Seed1
  If mvarSeed1 >= IM1 Then mvarSeed1 = mvarSeed1 - IM1 + 1
  If mvarSeed1 < 1 Then mvarSeed1 = mvarSeed1 + IM1 - 1
  
  mvarSeed2 = Seed2
  If mvarSeed2 >= IM2 Then mvarSeed2 = mvarSeed2 - IM2 + 1
  If mvarSeed2 < 1 Then mvarSeed2 = mvarSeed2 + IM2 - 1
End Sub

'uses the currnt time to randomise the various seeds    'this is VERY ARBITRARY

Public Sub Randomize()
  SeedRand Timer * 53, Timer * 91
End Sub

