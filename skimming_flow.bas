
Option Explicit

Public Function SKIMMING_FLOW(parametro As String, B As Double, Hdam As Double, S As Double, l As Double, Q As Double) As Variant

'********************************************************************************************************
'o código a seguir descreve a metodologia para cálculo de descidas d'água em degraus em regime
'skimming flow conforme OHTSU, I., YASUDA, Y., TAKAHASHI (2004). a metodologia é aplicável para
'ângulos da escada (Patamar/espelho) entre 5.7° e 55°.
'Autor: Pedro H. Lucena
'Data:21/05/2024
'********************************************************************************************************

'--------------------------------------------------------------------------------------------------------
' Variáveis de entrada da função SKIMMING_FLOW()
'--------------------------------------------------------------------------------------------------------

'Parâmetro de saída da variável, pondendo ser: Profundidade final, Velocidade Final, Energia Resi-
'dual e altura de referência da parede;
'B é a largura da escada em (m);
'Hdam é a altura total da escada em (m);
'S é o espelho da escada em (m);
'l é o patamar da escada em (m);
'Q é a vazão afluente à escada em (m3/s).

'--------------------------------------------------------------------------------------------------------
'cálculo da profundidade crítica, considerando seção retangular e Hdam/dc
'--------------------------------------------------------------------------------------------------------

Dim qw As Double, dc As Double, p1 As Double

qw = Q / B ' vazão específica da seção da escada
dc = (((qw) ^ 2) / 9.81) ^ (1 / 3)

'Cálculo do parâmetro não-dimensional Hdam/dc
p1 = Hdam / dc

'--------------------------------------------------------------------------------------------------------
'limite superior da altura relativa do degrau (S/dc)s por meio da variável Lm e theta
'--------------------------------------------------------------------------------------------------------

Dim Lm As Double, theta As Double, LmA As Double

theta = WorksheetFunction.Degrees(Atn(S / l))
Lm = (7 / 6) * (Tan(WorksheetFunction.Radians(theta))) ^ (1 / 6)

LmA = 13 * Tan(WorksheetFunction.Radians(theta)) ^ 2 - 2.73 * Tan(WorksheetFunction.Radians(theta)) + 0.373

'--------------------------------------------------------------------------------------------------------
'classificação do caso em questão em SK A e SK B e cálculo da do angulo de inclinação da descida d'água.
'--------------------------------------------------------------------------------------------------------

Dim Classif_skimming As String

If (theta < 5.7) Or (theta > 55) Or ((S / dc) > (Lm)) Then
    GoTo ErrorHandler1
Else
    If theta > 19 Then
        Classif_skimming = "B"
    Else
        If ((S / dc) > (LmA)) Then
            Classif_skimming = "A"
        Else
            Classif_skimming = "B"
        End If
    End If
End If

'--------------------------------------------------------------------------------------------------------
'verifica se o espelho (S) e o patamar (l) atendem aos limites previstos por OHTSU, I., YASUDA,
'Y., TAKAHASHI (2004). Caso não atenda, deve-se alterar os valores de l e/ou S para ficar dentro
'da faixa.
'--------------------------------------------------------------------------------------------------------

If (0.25 > (S / dc)) Or ((S / dc) > Lm) Then GoTo ErrorHandler1

'--------------------------------------------------------------------------------------------------------
'Cálculo do parâmetro não-dimensional altura de queda (Hdam/dc) da escada a ser dimensionada.
'--------------------------------------------------------------------------------------------------------

Dim dh1 As Double
dh1 = Hdam / dc

'--------------------------------------------------------------------------------------------------------
'Cálculo do parâmetro não-dimensional altura de queda (He/dc) de referência para verificação da
'da ocorrência de fluxo quasi-uniforme (Ohtsu et al., 2004, p.862, eq.3).
'--------------------------------------------------------------------------------------------------------

Dim e As Double 'Variável que armazena o número neperiano até a 14° casa decimal.
Dim dh2 As Double

e = 2.71828182845904

dh2 = ((-1.21 * (10 ^ (-5)) * (theta ^ 3) + 1.6 * (10 ^ (-3)) * (theta ^ 2) - 7.13 * (10 ^ (-2)) * theta + 1.3) ^ (-1)) * (5.7 + 6.7 * e ^ (-6.5 * (S / dc)))

'--------------------------------------------------------------------------------------------------------
'Cálculo do fator de fricção (Ohtsu et al., 2004, p. 864, eq. 12a e eq. 12b).
'--------------------------------------------------------------------------------------------------------

Dim a As Double
Dim fmax As Double
Dim f As Double

Select Case theta
    
    'Regime de fluxo não uniforme
    Case Is < 19
        If (0.1 <= S / dc) And (S / dc <= 0.5) Then
            a = -1.7 * (10 ^ (-3)) * (theta ^ 2) + 6.4 * (10 ^ (-2)) * theta - 1.5 * 10 ^ (-1)
            fmax = -4.2 * (10 ^ (-4)) * (theta ^ 2) + 1.6 * (10 ^ (-2)) * theta + 3.2 * (10 ^ (-2))
            f = fmax - a * (0.5 - S / dc) ^ 2
        Else
            If (0.5 <= S / dc) And S / dc <= Lm Then
               fmax = -4.2 * (10 ^ (-4)) * (theta ^ 2) + 1.6 * (10 ^ (-2)) * theta + 3.2 * (10 ^ (-2))
               f = fmax
            Else: GoTo ErrorHandler1
            End If
        End If
        
        fmax = -4.2 * (10 ^ (-4)) * (theta ^ 2) + 1.6 * (10 ^ (-2)) * theta + 3.2 * (10 ^ (-2))
    
    'Regime de fluxo quasi-uniforme
    Case Is > 19
        If (0.1 <= S / dc) And (S / dc <= 0.5) Then
            a = 0.452
            fmax = 2.32 * 10 ^ (-5) * (theta ^ 2) - 2.75 * 10 ^ (-3) * (theta) + 2.31 * (10 ^ (-1))
            f = fmax - a * (0.5 - S / dc) ^ 2
        Else
            If (0.5 <= S / dc) And S / dc <= Lm Then
               fmax = 2.32 * 10 ^ (-5) * (theta ^ 2) - 2.75 * 10 ^ (-3) * (theta) + 2.31 * (10 ^ (-1))
               f = fmax
            Else: GoTo ErrorHandler1
            End If
        End If
End Select

'--------------------------------------------------------------------------------------------------------
'Há uma diferenciação no fator de fricção (f) em função de S/dc para ângulos <=19° e para valores
'entre 19° e 55°. Verifica-se, portanto, qual o caso em questão.
'--------------------------------------------------------------------------------------------------------

Dim Estimativa_Eres As Double 'Energia residual por altura crítica Eres/dc estimada

Dim m As Double 'variável presente na equação da energia residual para fluxo não-uniforme
Dim dw As Double 'variável que armazena a profundidade nomal no final da escada
Dim U As Double

m = -(theta / 25) + 4

Select Case theta
    
    Case Is <= 19
    
    
    If (dh1 >= dh2) Then
    'regime de fluxo quasi-uniforme
        'dw = dc * (f / (8 * Sin(WorksheetFunction.Radians(theta))) ^ (1 / 3))
        
        If Classif_skimming = "A" Then
            
            U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) * Cos(WorksheetFunction.Radians(theta)) + _
            (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
            Estimativa_Eres = U * dc
            dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
        Else
            U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) + _
            (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
            Estimativa_Eres = U * dc
            dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
        End If
    
    Else
    'regime de fluxo não-uniforme
        If Classif_skimming = "A" Then
            'Tipo A
            U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) * Cos(WorksheetFunction.Radians(theta)) + _
            (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
            Estimativa_Eres = (1.5 + (U - 1.5) * (1 - (1 - (dh1 / dh2)) ^ m)) * dc
            dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
        Else
            'Tipo B
            U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) + _
            (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
            Estimativa_Eres = (1.5 + (U - 1.5) * (1 - (1 - (dh1 / dh2)) ^ m)) * dc
            dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
        End If
        
    End If

    Case Is > 19
    
        If (dh1 >= dh2) Then
            'dw = dc * (f / (8 * Sin(WorksheetFunction.Radians(theta))) ^ (1 / 3))
            If Classif_skimming = "A" Then
                
                U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) * Cos(WorksheetFunction.Radians(theta)) + _
                (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
                Estimativa_Eres = U * dc
                dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
            Else
                U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) + _
                (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
                Estimativa_Eres = U * dc
                dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
            End If
        Else 'regume de fluxo não-uniforme
            U = ((f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (1 / 3)) * Cos(WorksheetFunction.Radians(theta)) + _
            (1 / 2) * (f / (8 * Sin(WorksheetFunction.Radians(theta)))) ^ (-2 / 3)
            Estimativa_Eres = (1.5 + (U - 1.5) * (1 - (1 - (dh1 / dh2)) ^ m)) * dc
            dw = Find_dw(Classif_skimming, Estimativa_Eres, qw, theta)
        End If
End Select

'--------------------------------------------------------------------------------------------------------
'Cálculo da concentração média de ar no regime skimming flow
'--------------------------------------------------------------------------------------------------------
    
Dim Cmean As Double, D As Double, Ya As Double, Hw As Double

If (theta <= 19) Then
    D = 0.3
Else
    D = -2 * (10 ^ (-4)) * (theta ^ 2) + 2.14 * (10 ^ (-2)) * theta - 3.57 * (10 ^ (-2))
End If

Cmean = D - 0.3 * (e ^ (-5 * ((S / dc) ^ 2) - (4 * S / dc)))

Ya = dw / (1 - Cmean)
    
Hw = 1.4 * Ya

'--------------------------------------------------------------------------------------------------------
'Cálculo da energia dissipada
'--------------------------------------------------------------------------------------------------------

Dim Hss As Double

Hss = Hdam + 1.5 * dc - Estimativa_Eres

'--------------------------------------------------------------------------------------------------------
'Definição do resultado da função conforme escolha do usuário
'--------------------------------------------------------------------------------------------------------

'Parâmetro de saída da variável, pondendo ser: Profundidade final, Velocidade Final, Energia Resi-
'dual e altura de referência da parede;

Select Case parametro
    
    Case "Profundidade final"
    
    SKIMMING_FLOW = dw
    
    Case "Velocidade final"
    
    SKIMMING_FLOW = (qw / dw)
    
    Case "Energia residual"
    
    SKIMMING_FLOW = Estimativa_Eres
    
    Case "Energia dissipada"
    
    SKIMMING_FLOW = Hss
    
    Case "Altura da parede"
    
    SKIMMING_FLOW = Hw
    
End Select

Exit Function

ErrorHandler1:
    SKIMMING_FLOW = "Verif. s/l"
    Exit Function

ErrorHandler2:
    SKIMMING_FLOW = "diminuir H"
    Exit Function

End Function

Public Function Find_dw(Classif_skimming As String, Eres As Double, qw As Double, theta As Double) As Double


Dim i As Long

Dim y As Double, ynext As Double

Dim funcao As Double, derivada As Double


'1º palpite

y = 0.1

For i = 1 To 1500
    
    Select Case Classif_skimming
    
        Case "A"
        funcao = Eres - y - (1 / (2 * 9.81)) * (qw / y) ^ 2
        derivada = -1 + (qw ^ 2) / (9.81 * (y ^ 3))
        
        Case "B"
        funcao = Eres - y * Cos(WorksheetFunction.Radians(theta)) - (1 / (2 * 9.81)) * (qw / y) ^ 2
        derivada = -1 * Cos(WorksheetFunction.Radians(theta)) + (qw ^ 2) / (9.81 * (y ^ 3))

    End Select
    
    ynext = y - funcao / derivada
    
    If (Abs(y - ynext) <= 0.000000000000001) Then
        Exit For
    Else
        y = ynext
    End If
Next i

Find_dw = y

End Function

