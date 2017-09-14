Attribute VB_Name = "mChgPt"
Option Explicit

'========================================
'Bayesian Online Change Point Detection
'========================================
'=======================================================================
'Main reference is "Bayesian Online Changepoint Detection, RP Adam, D MacKay (2007)"
'The authors had a Matlab implementation here:
'http://hips.seas.harvard.edu/content/bayesian-online-changepoint-detection, but that
'implementaion does not benefit from the online capability of the algorithm, and
'requires a large (n_T x n_T) array in memory. It's re-written here so that data point
'can be fed in one by one.
'Default conjugate prior used is normal-inverse-gamma which is suitable for gaussian
'process with unknown mean and variance.
'=======================================================================

'Init(R_len() As Long, R_prob() As Double, Optional boundary As Long = 0)
'    Desc:      Initialize probability of different run lengths at t=0. By default
'               (the only choice at the moment) a change point is assumed to have
'               occured before first data point. So R_len() and R_prob() will be
'               vectors of size 1, with R_len(1)=0 and R_Prob(1)=1
'    Output:    R_len(), empty integer vector
'               R_prob(), empty real vector


'Init_Priors(priors As Variant, init_values As Variant, Optional PriorType As String = "NIG")
'    Desc:      Initialize hyperparameters of prior probablily distribution.
'    Input:     init_values, a 1-D vector holding the desired value of each parameter
'                   normal-inverse-gamma: 4 parameters    : mu, kappa, alpha, beta
'                   gamma               : 2+1 parameters  : alpha, beta, known mu
'               PriorType,  default is normal-inverse-gamma ("NIG"), the other chose is gamma ("GAMMA")
'    Output:    priors, empty variant. Jagged array holding vectors for each parameter.


'ChgPt_Chk(x As Double, R_len() As Long, R_prob() As Double, priors As Variant, max_r As Long, _
'            Optional lambda As Double = 200, Optional tol As Double = 0.0001, Optional PriorType As String = "NIG")
'    Desc:      Givent a new data point x, update the probability of each run length.
'               Also update the hyperparameters of priors at different run lengths.
'    Input:     x, a new data point
'               R_len(), an integer vector storing run lengths from last time-step.
'               R_prob(), a real vector storing probablilty of each run length from last time_step.
'               priors, jagged array storing the prior parameters at different run length
'               lambda, hazard function time scale, the larger this is, the less likely a change point has occured
'               tol, probability smaller than this level will be removed
'               PriorType, distribution of prior. Can be normal-inverse-gamma ("NIG") or gamma ("GAMMA")
'    Output:    max_r, an integer that represents the most probable run length
'               R_len(), R_prob() and priors are replaced by udpated values


'ChgPt_Series(x() As Double, R_len() As Long, R_prob() As Double, priors As Variant, R_t() As Long, _
            Optional lambda As Double = 200, Optional tol As Double = 0.0001, _
            Optional PriorType As String = "NIG", Optional saveR As Variant)
'    Desc:      Apply ChgPt_Chk sequentially on time series x().
'    Input:     x(), time series of size 1:N
'               R_len(), an integer vector storing run lengths from last time-step.
'               R_prob(), a real vector storing probablilty of each run length from last time_step.
'               priors, jagged array storing the prior parameters at different run length
'               lambda, hazard function time scale, the larger this is, the less likely a change point has occured
'               tol, probability smaller than this level will be removed
'               PriorType, distribution of prior. Can be normal-inverse-gamma ("NIG") or gamma ("GAMMA")
'    Output:    R_t(), an integer vector of size 1:N, holding most probably run length at each time step
'               R_len(), R_prob() and priors are replaced by last udpated values
'               savePriors, if provided will store the estimated hyper-parameters at each time step.


'=== Initialize P(r_0)
Sub Init(R_len() As Long, R_prob() As Double, Optional boundary As Long = 0)
    '=== Initialize R()
    ReDim R_len(1 To 1)
    ReDim R_prob(1 To 1)
    If boundary = 0 Then 'Assume change point occurred before first data
        R_len(1) = 0    'run length is zero
        R_prob(1) = 1   'P(r=0)=1
    ElseIf boundary = 1 Then
        Debug.Print "mChgPt: Failed:" & boundary & " boundary condition not implemented."
    End If
End Sub

'=== Initialize parameters of conjugate prior
'NIG    : normal-inverse-gamma: 4 parameters    : mu, kappa, alpha, beta
'GAMMA  : gamma               : 2+1 parameters  : alpha, beta, mu is known
Sub Init_Priors(priors As Variant, init_values As Variant, Optional PriorType As String = "NIG")
Dim i As Long, m As Long, n As Long
Dim vec1() As Double
If PriorType = "NIG" Or PriorType = "GAMMA" Then
    m = LBound(init_values, 1)
    n = UBound(init_values, 1)
    ReDim priors(1 To n - m + 1)
    ReDim vec1(1 To 1)
    For i = 1 To n - m + 1
        vec1(1) = init_values(m + i - 1)
        priors(i) = vec1
    Next i
    Erase vec1
Else
    Debug.Print "mChgPt: Failed: prior " & PriorType & " not implemented."
End If
End Sub


'=== Run algorithm over a time series
'Input: x(), 1:n real vector signal
'       R_len() & R_prob(), initialized by Init()
'       priors, initialized by Init_Priors()
'       lambda, hazard function time scale
'       tol, probability smaller than this level will be removed
'       PriorType, distribution of prior
'Output: R_t(), 1:n integer vector tracing out the run length envelop
'        self, returns an integer array of change points index positions
'        R_len(),R_prob() and priors are replaced by udpated values
Sub ChgPt_Series(x() As Double, R_len() As Long, R_prob() As Double, priors As Variant, R_t() As Long, _
            Optional lambda As Double = 200, Optional tol As Double = 0.0001, _
            Optional PriorType As String = "NIG", Optional savePriors As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_R As Long
Dim tmp_vec() As Double, iArr() As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(x, 1)
    ReDim R_t(1 To n)
    If IsMissing(savePriors) = False Then ReDim savePriors(1 To n, 1 To UBound(priors))
    For i = 1 To n
        If i Mod 500 = 0 Then
            DoEvents
            Application.StatusBar = "Scanning signal: " & i & "/" & n
        End If
        Call ChgPt_Chk(x(i), R_len, R_prob, priors, k, lambda, tol, PriorType)
        R_t(i) = k
        If IsMissing(savePriors) = False Then
            n_R = UBound(R_prob, 1)
            For k = 1 To UBound(priors)
                tmp_vec = priors(k)
                tmp_x = 0: tmp_y = 0
                For j = 1 To n_R
                    tmp_x = tmp_x + tmp_vec(j) * R_prob(j)
                    tmp_y = tmp_y + R_prob(j)
                Next j
                savePriors(i, k) = tmp_x / tmp_y
            Next k
        End If
        
    Next i
    
    Application.StatusBar = False
End Sub

'=== Main Rountine:
'Input: x, a new data point
'       R_len(), an integer vector storing run lengths from last time-step,i.e. r_{t-1}
'       R_prob(), a real vector storing probablilty of each run length from last time_step P(r_{t-1},x(1:t-1))
'       priors, jagged array storing the prior parameters at different run length
'       e.g. for normal-inverse-gama prior, there are 4 parameters, so priors has dimension 1:4,
'       each dimension is a real vector holding a parameter's value at different run lenght,e.g. mu(r_t)
'       lambda, hazard function time scale
'       tol, probability smaller than this level will be removed
'       PriorType, distribution of prior
'Output:  max_r, an integer that represents the most probable run length, i.e. r_t that maximize P(r_t,x(1:t))
'         R_len(),R_prob() and priors are replaced by udpated values
Sub ChgPt_Chk(x As Double, R_len() As Long, R_prob() As Double, priors As Variant, max_r As Long, _
            Optional lambda As Double = 200, Optional tol As Double = 0.0001, Optional PriorType As String = "NIG")
Dim i As Long, j As Long, k As Long, n As Long, n_R As Long
Dim tmp_x As Double, tmp_y As Double
Dim R_prob_new() As Double, predprobs() As Double, h() As Double
Dim R_len_new() As Long, keeper() As Long

    n_R = UBound(R_len, 1)

    Call Calc_Predictive(predprobs, x, priors, PriorType)
    Call Calc_hazard_func(h, R_len, lambda)
    
    ReDim R_prob_new(1 To n_R + 1)
    ReDim R_len_new(1 To n_R + 1)
    
    R_len_new(1) = 0    'change point has occured, reset run length
    For i = 1 To n_R    'no change point, increment run lenghts by 1
        R_len_new(i + 1) = R_len(i) + 1
    Next i
    
    'Calculate growth and change point probabilities
    tmp_x = 0
    For i = 1 To n_R
        R_prob_new(i + 1) = R_prob(i) * (1 - h(i)) * predprobs(i)
        tmp_x = tmp_x + R_prob(i) * h(i) * predprobs(i)
    Next i
    R_prob_new(1) = tmp_x
    
    Erase predprobs, h
    
    'Normalize probabilities to 1
    tmp_x = 0
    For i = 1 To n_R + 1
        tmp_x = tmp_x + R_prob_new(i)
    Next i
    For i = 1 To n_R + 1
        R_prob_new(i) = R_prob_new(i) / tmp_x
    Next i
    
    'Prune small probabilities for better efficiency
    Call Prune_R(R_len_new, R_prob_new, keeper, tol)
    n_R = UBound(R_len_new)
    
    'Update prior for next time step
    Call Update_Prior(x, priors, keeper, PriorType)
    
    R_len = R_len_new
    R_prob = R_prob_new
    
    'Identify the most probable run length
    tmp_x = -1
    j = 0
    For i = 1 To n_R
        If R_prob_new(i) > tmp_x Then
            tmp_x = R_prob_new(i)
            max_r = R_len_new(i)
            j = i
        End If
    Next i
    
    Erase R_len_new, R_prob_new, keeper
End Sub




'=== Calculate P(x_t | r_{t-1})
Private Sub Calc_Predictive(predprobs() As Double, x As Double, priors As Variant, Optional PriorType As String = "NIG")
Dim i As Long, m As Long, n As Long
Dim mu() As Double, kappa() As Double, alpha() As Double, beta() As Double
Dim nu() As Double, var() As Double
    
    If PriorType = "NIG" Then
        '=== normal-inverse-gamma
        mu = priors(1)
        kappa = priors(2)
        alpha = priors(3)
        beta = priors(4)
        n = UBound(alpha, 1)
        ReDim nu(1 To n)
        ReDim var(1 To n)
        For i = LBound(nu, 1) To UBound(nu, 1)
            nu(i) = 2 * alpha(i)
            var(i) = beta(i) * (kappa(i) + 1) / (alpha(i) * kappa(i))
        Next i
        predprobs = student_pdfs(x, nu, mu, var)
        Erase nu, var, mu, kappa, alpha, beta
    
    ElseIf PriorType = "GAMMA" Then
        '=== Gamma
        alpha = priors(1)
        beta = priors(2)
        mu = priors(3)
        n = UBound(alpha, 1)
        ReDim nu(1 To n)
        ReDim var(1 To n)
        For i = LBound(nu, 1) To UBound(nu, 1)
            nu(i) = 2 * alpha(i)
            var(i) = beta(i) / alpha(i)
        Next i
        predprobs = student_pdfs(x, nu, mu, var)
        Erase nu, var, mu, alpha, beta
        
    Else
        '=== To implement other types of conjugate prior?
        Debug.Print "Calc_Predictive: Failed: " & PriorType & " not implemented."
    End If
End Sub

'Hazard Function
'Probabiltiy that a run length will drop to zero
Private Sub Calc_hazard_func(h() As Double, R_len() As Long, lambda As Double)
Dim i As Long, m As Long, n As Long
    m = LBound(R_len, 1)
    n = UBound(R_len, 1)
    ReDim h(m To n)
    For i = m To n
        h(i) = 1 / lambda  'Assume to be constant rate in this implementation
    Next i
End Sub


'Remove run lengths that have little probaabilities.
'keeper() is an integer array to keep track of which run length is being kept.
Private Sub Prune_R(R_len() As Long, R_prob() As Double, keeper() As Long, _
        Optional tol As Double = 0.0001, Optional min_len As Long = 1)
Dim i As Long, j As Long, k As Long, n_R As Long
    n_R = UBound(R_len)
    k = 0
    ReDim keeper(1 To n_R)
    For i = 1 To n_R
        If R_prob(i) >= tol Or R_len(i) <= min_len Then
            k = k + 1
            keeper(k) = i
        End If
    Next i
    ReDim Preserve keeper(1 To k)
    
    For i = 1 To k
        j = keeper(i)
        R_len(i) = R_len(j)
        R_prob(i) = R_prob(j)
    Next i
    
    ReDim Preserve R_len(1 To k)
    ReDim Preserve R_prob(1 To k)
End Sub


'Update prior for next time step
'Input: priors of different run lengths from current time step
'       keeper(), integer array indicating run lenth to be kept after pruning
Private Sub Update_Prior(x As Double, priors As Variant, keeper() As Long, _
                PriorType As String)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_R As Long
Dim mu() As Double, kappa() As Double, alpha() As Double, beta() As Double
Dim mu_tmp() As Double, kappa_tmp() As Double, alpha_tmp() As Double, beta_tmp() As Double
Dim mu0 As Double
    n_R = UBound(keeper, 1)
    
    If PriorType = "NIG" Then
        '=== Normal-inverse-gamma
        ReDim mu_tmp(1 To n_R)
        ReDim kappa_tmp(1 To n_R)
        ReDim alpha_tmp(1 To n_R)
        ReDim beta_tmp(1 To n_R)
        
        mu = priors(1)
        kappa = priors(2)
        alpha = priors(3)
        beta = priors(4)
        
        mu_tmp(1) = mu(1)
        kappa_tmp(1) = kappa(1)
        alpha_tmp(1) = alpha(1)
        beta_tmp(1) = beta(1)
        
        For k = 2 To n_R
            i = keeper(k) - 1
            mu_tmp(k) = (kappa(i) * mu(i) + x) / (kappa(i) + 1)
            kappa_tmp(k) = kappa(i) + 1
            alpha_tmp(k) = alpha(i) + 0.5
            beta_tmp(k) = beta(i) + (kappa(i) * ((x - mu(i)) ^ 2)) / (2 * (kappa(i) + 1))
        Next k
        
        priors(1) = mu_tmp
        priors(2) = kappa_tmp
        priors(3) = alpha_tmp
        priors(4) = beta_tmp
    
        Erase mu, kappa, alpha, beta
        Erase mu_tmp, kappa_tmp, alpha_tmp, beta_tmp

    ElseIf PriorType = "GAMMA" Then
        '=== gamma
        mu0 = 0
        ReDim alpha_tmp(1 To n_R)
        ReDim beta_tmp(1 To n_R)
        ReDim mu_tmp(1 To n_R)
        
        alpha = priors(1)
        beta = priors(2)
        mu = priors(3)
        
        alpha_tmp(1) = alpha(1)
        beta_tmp(1) = beta(1)
        mu_tmp(1) = mu(1)
        
        For k = 2 To n_R
            i = keeper(k) - 1
            alpha_tmp(k) = alpha(i) + 0.5
            beta_tmp(k) = beta(i) + ((x - mu(i)) ^ 2) / 2
            mu_tmp(k) = mu(i)
        Next k

        priors(1) = alpha_tmp
        priors(2) = beta_tmp
        priors(3) = mu_tmp
        Erase alpha, beta, mu
        Erase alpha_tmp, beta_tmp, mu_tmp

    Else
        '=== To implement other types of conjugate prior?
        Debug.Print "Update_Prior: Prior " & PriorType & " not implemented."
    End If
    
    
End Sub




'=============================================
'Predictive Posteriors
'=============================================

'pdf of student-t distribution
'nu = degree of freedom
Private Function student_pdf(x As Double, nu As Double, Optional mu As Double = 0, Optional var As Double = 1) As Double
    student_pdf = Exp(gammaln((nu + 1) * 0.5) - gammaln(nu * 0.5)) / Sqr(nu * 3.14159265358979 * var)
    student_pdf = student_pdf * ((1 + ((x - mu) ^ 2) / (nu * var)) ^ (-(nu + 1) * 0.5))
End Function

'To speed up calculation of Student's pdf of a single value of x for multiple parameters
Private Function student_pdfs(x As Double, nu() As Double, mu() As Double, var() As Double) As Double()
Dim i As Long, m As Long, n As Long
Dim tmp_x As Double
Dim vec1() As Double, vec2() As Double
    m = LBound(nu, 1)
    n = UBound(nu, 1)
    ReDim vec1(m To n)
    ReDim vec2(m To n)
    For i = m To n
        vec1(i) = (nu(i) + 1) * 0.5
        vec2(i) = nu(i) * 0.5
    Next i
    vec1 = gammalns(vec1)
    vec2 = gammalns(vec2)
    For i = m To n
        tmp_x = Exp(vec1(i) - vec2(i)) / Sqr(nu(i) * 3.14159265358979 * var(i))
        vec1(i) = tmp_x * ((1 + ((x - mu(i)) ^ 2) / (nu(i) * var(i))) ^ (-(nu(i) + 1) * 0.5))
    Next i
    student_pdfs = vec1
    Erase vec1, vec2
End Function



'=============================================
'Special Functions
'=============================================

'Returns ln|gamma(x)| for x>0
'Lanczos approximation from Numerical Recipe FORTRAN77 Chapter 6.1
'x can either be a single or a vector of real positive numbers
Private Function gammalns(x As Variant) As Variant
Dim i As Long, j As Long, m As Long, n As Long
Dim ser As Double, stp As Double, tmp As Double, z As Double
Dim cof() As Double, y() As Double
    ReDim cof(1 To 6)
    cof(1) = 76.1800917294715       '76.18009172947146d0
    cof(2) = -86.5053203294168      '-86.50532032941677d0
    cof(3) = 24.0140982408309       '24.01409824083091d0
    cof(4) = -1.23173957245015      '-1.231739572450155d0
    cof(5) = 1.20865097386618E-03   '.1208650973866179d-2
    cof(6) = -5.395239384953E-06    '-.5395239384953d-5
    stp = 2.506628274631            '2.5066282746310005d0
    If IsArray(x) = False Then
        tmp = x + 5.5
        tmp = (x + 0.5) * Log(tmp) - tmp
        ser = 1.00000000019001          '1.000000000190015d0
        For j = 1 To 6
            ser = ser + cof(j) / (x + j)
        Next j
        gammalns = tmp + Log(stp * ser / x)
    Else
        m = LBound(x, 1)
        n = UBound(x, 1)
        ReDim y(m To n)
        For i = m To n
            z = x(i)
            tmp = z + 5.5
            tmp = (z + 0.5) * Log(tmp) - tmp
            ser = 1.00000000019001      '1.000000000190015d0
            For j = 1 To 6
                ser = ser + cof(j) / (z + j)
            Next j
            y(i) = tmp + Log(stp * ser / z)
        Next i
        gammalns = y
        Erase y
    End If
    Erase cof
End Function
