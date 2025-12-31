Sub CalcMaxPower ()
  
'Test Data
  Qes = 4
  Vas = 2.17
  Fs = 31
  Qts = .6
  
  'diameter in inches
  ACDI = 8
  
  'peak to peak driver excursion
  EXCI = .75
  
  'maximum input in watts RMS
  WRMS = 100
  
  'dc resistance of voice coil
  Re = 3.1
'_____________

 
  VRMS = Sqr(WRMS * Re)
  KP = Sqr(Vas * Fs ^ 3 / Qes / Re) / 753
  KD = .3535 * Sqr(Vas / Fs / Qes / Re) / ACDI ^ 2
  
  If TypeofDesign$ = "optimal ported" Or TypeofDesign$ = "non-optimal, ported" Then
    H = FB / Fs
    A = Vas / VB
    Q7 = 7
  Else
    H = .0001
    Q7 = 10000
  End If
 
  H2 = Sqr(H)
  A1 = H2 / Q7 + 1 / Qts / H2
  A2 = 1 / Qts / Q7 + (A + 1 + H ^ 2) / H
  A3 = H2 / Qts + 1 / Q7 / H2
  F1 = H2 * Fs

  For F% = 1 To 200 Step 1
   X = F% / F1
   Y = X ^ 4 / Sqr((X ^ 4 - A2 * X ^ 2 + 1) ^ 2 + (A1 * X ^ 3 - A3 * X) ^ 2)
   Y1 = Y * Sqr((X ^ 2 / H - 1) ^ 2 + (X / Q7) ^ 2 / H) / X ^ 4
   E1 = EXCI / KD / Y1 / 39.37
   If E1 < VRMS Then E2 = E1
    If VRMS <= E1 Then E2 = VRMS
    RESPONSE = 20 * Log(Y) / Log(10)
    SPL = 20 * Log(KP * Y * E2 * 50000!) / Log(10)
    POWER = E2 ^ 2 / Re
  Next F%
     
End Sub

