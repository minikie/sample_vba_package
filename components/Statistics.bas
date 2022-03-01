Function ncdchi(chi2 As Double, df As Double, lambda As Double)
   ncdchi = cumchn(chi2, df, lambda)
End Function


Private Function qsmall(X As Double, sum As Double) As Boolean
' cumchn, cumfn と同じル?ティン
     Dim eps As Double
       eps = 0.000001 '1e-6 '0.000001
'     ..
'     .. Statement Function definitions ..
      qsmall = (sum < 1E-20) Or (X < eps * sum)
End Function


Function cumchn(X As Double, df As Double, pnonc As Double)
'***********************************************************************

'             CUMulative of the Non-central CHi-square distribution
'
'DCDFLIB： http://odin.mdacc.tmc.edu/anonftp/page_2.html#DCDFLIB
'Section of Computer Science, Department of Biomathematics, University of Texas M.D. Anderson Hospital.
'の cumchn.f を excel vba に移植したものである。excel の統計??を使っている。
'移植に?たって，構造化し goto 文をなくした。 excel 97 を使って確認している。2000/5/24

'                               Function
'
'     Calculates     the       cumulative      non-central    chi-square
'     distribution, i.e.,  the probability   that  a   random   variable
'     which    follows  the  non-central chi-square  distribution,  with
'     non-centrality  parameter    PNONC  and   continuous  degrees   of
'     freedom DF, is less than or equal to X.
'
'                              Arguments
'
'     X       --> Upper limit of integration of the non-central
'                 chi-square distribution.非心χ2値
'
'
'     DF      --> Degrees of freedom of the non-central
'                 chi-square distribution.自由度
'
'
'     PNONC   --> Non-centrality parameter of the non-central
'                 chi-square distribution. 非心度パラメ?タ
'
'     CUM <-- Cumulative non-central chi-square distribution.
'                       返値 累積非心χ2分布（確率）
'

'
'                                Method(もとプログラムに書いてあったもの)
'
'     Uses  formula  26.4.25   of  Abramowitz  and  Stegun, Handbook  of
'     Mathematical    Functions,  US   NBS   (1966)    to calculate  the
'     non-central chi-square.
'
'                                Variables
'
'     EPS     --- Convergence criterion.  The sum stops when a
'                 term is less than EPS*SUM.
'
'
'***********************************************************************
'
'

      Dim adj As Double, centaj As Double, centwt As Double, chid2 As Double, dfd2 As Double, eps As Double, lcntaj As Double, lcntwt As Double, lfact As Double, pcent As Double, pterm As Double, sum As Double, sumadj As Double, term As Double, wt As Double, xnonc As Double, xx As Double
      Dim i As Long, icent As Long
'      Dim cum As Double

'
      If (X <= 0#) Then
        cumchn = 0#
        Exit Function
      End If

      If (pnonc <= 0.0000000001) Then
'
'
'     When non-centrality parameter is (essentially) zero,
'     use cumulative chi-square distribution
'
'
        cumchn = 1 - Application.ChiDist(X, df)
        Exit Function
      End If

      xnonc = pnonc / 2#
'***********************************************************************
'
'     The following code calcualtes the weight, chi-square, and
'     adjustment term for the central term in the infinite series.
'     The central term is the one in which the poisson weight is
'     greatest.  The adjustment term is the amount that must
'     be subtracted from the chi-square to move up two degrees
'     of freedom.
'
'***********************************************************************
      icent = xnonc
      If (icent = 0) Then icent = 1
      chid2 = X / 2#
'
'
'     Calculate central weight term
'
'
      lfact = Application.GammaLn(CDbl(icent + 1))
      lcntwt = -xnonc + icent * Log(xnonc) - lfact
      centwt = Exp(lcntwt)
'
'
'     Calculate central chi-square
'
'
'      CALL cumchi(x,dg(icent),pcent,ccum)
       pcent = 1 - Application.ChiDist(X, dg(icent, df))
'
'
'     Calculate central adjustment term
'
'
      dfd2 = dg(icent, df) / 2#
      lfact = Application.GammaLn(1# + dfd2)
      lcntaj = dfd2 * Log(chid2) - chid2 - lfact
      centaj = Exp(lcntaj)
      sum = centwt * pcent
'***********************************************************************
'
'     Sum backwards from the central term towards zero.
'     Quit whenever either
'     (1) the zero term is reached, or
'     (2) the term gets small relative to the sum, or
'
'***********************************************************************
      sumadj = 0#
      adj = centaj
      wt = centwt
      i = icent
'
      Do
        dfd2 = dg(i, df) / 2#
'
'
'     Adjust chi-square for two fewer degrees of freedom.
'     The adjusted value ends up in PTERM.
'
'
        adj = adj * dfd2 / chid2
        sumadj = sumadj + adj
        pterm = pcent + sumadj
'
'
'     Adjust poisson weight for J decreased by one
'
'
        wt = wt * (i / xnonc)
        term = wt * pterm
        sum = sum + term
        i = i - 1
      Loop Until (qsmall(term, sum) Or (i = 0))

      sumadj = centaj
'***********************************************************************
'
'     Now sum forward from the central term towards infinity.
'     Quit when either
'     (1) the term gets small relative to the sum, or
'
'***********************************************************************
      adj = centaj
      wt = centwt
      i = icent
'
'
'     Update weights for next higher J
'
'
      Do
        wt = wt * (xnonc / (i + 1))
'
'
'     Calculate PTERM and add term to sum
'
'
        pterm = pcent - sumadj
        term = wt * pterm
        sum = sum + term
'
'
'     Update adjustment term for DF for next iteration
'
'
        i = i + 1
        dfd2 = dg(i, df) / 2#
        adj = adj * chid2 / dfd2
        sumadj = sumadj + adj
      Loop Until (qsmall(term, sum))

      cumchn = sum

End Function


Function dg(i As Long, df)
     dg = df + 2# * CDbl(i)
End Function

Function chi_density(X As Double, V As Double) As Double

Dim res As Double
Dim V_2 As Double

V_2 = 0.5 * V
res = X ^ (V_2 - 1) * Exp(-0.5 * X)
chi_density = res / ((2 ^ V_2) * WorksheetFunction.Gamma(V_2))

End Function


Function ncx2pdf(X As Double, V As Double, DELTA As Double) As Double

Dim res As Double
Dim value As Double
Dim ln_z_2 As Double

Dim s As Long

res = 0

For s = 0 To 50

value = Exp(-0.5 * DELTA) * (DELTA / 2) ^ s
value = value * chi_density(X, V + 2 * s)
value = value / WorksheetFunction.Fact(s)

res = res + value

Next s

ncx2pdf = res


End Function
