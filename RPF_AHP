

Type Member
m As Double
v As Double
End Type

Type RRough
u As Double
l As Double
End Type

Type Rough
u As Member
l As Member
End Type









Sub RPF_AHP()

Dim x, z As Integer


'x = InputBox("Kriter sayısını giriniz")
'z = InputBox("Uzman sayısını giriniz")
x = 8
z = 9

Dim EXP() As Member
ReDim EXP(z, x, x) As Member

Dim Karar() As Rough
ReDim Karar(z, x, x) As Rough

Dim rkarar() As Rough
ReDim rkarar(x, x) As Rough

Dim difmat() As RRough
ReDim difmat(x, x) As RRough

Dim mulmat() As RRough
ReDim mulmat(x, x) As RRough

Dim detmat() As Double
ReDim detmat(x, x) As Double

Dim wemat() As Double
ReDim wemat(x, x) As Double

Dim up As Member
Dim lo As Member
Dim sayu As Member
Dim sayl As Member

Dim Weight() As Double
ReDim Weight(x) As Double




For k = 1 To z
For j = 1 To x

    For i = 1 To x
    
        
           
        If Cells(i + (k - 1) * x, j) > 0 Then
        
        
        
        
        EXP(k, i, j).m = 0.5 + (Cells(i + (k - 1) * x, j) - 1) * 0.05
        EXP(k, i, j).v = 0.5 - (Cells(i + (k - 1) * x, j) - 1) * 0.05
        EXP(k, j, i).m = EXP(k, i, j).v
        EXP(k, j, i).v = EXP(k, i, j).m
        
        
                    
        
        Cells(i + (k - 1) * x, x + 2 + j) = CStr(EXP(k, i, j).m) & ";" & CStr(EXP(k, i, j).v)
        Cells(j + (k - 1) * x, x + 2 + i) = CStr(EXP(k, j, i).m) & ";" & CStr(EXP(k, j, i).v)
                
        End If
       
 
     
            
            
    Next i

Next j
Next k

Top = 0
For i = 1 To x
    For j = 1 To x
    
        rkarar(i, j).l.m = 1
        rkarar(i, j).u.m = 1
        
        rkarar(i, j).l.v = 1
        rkarar(i, j).u.v = 1
    
        For t = 1 To z
        up.m = 0
        lo.m = 0
       
        sayu.m = 0
        sayl.m = 0
        
        up.v = 0
        lo.v = 0
       
        sayu.v = 0
        sayl.v = 0
        
            For s = 1 To z
                If EXP(t, i, j).m <= EXP(s, i, j).m Then
                
                sayu.m = sayu.m + 1
                up.m = up.m + EXP(s, i, j).m
                
                End If
                
                
                If EXP(t, i, j).m >= EXP(s, i, j).m Then
                
                sayl.m = sayl.m + 1
                lo.m = lo.m + EXP(s, i, j).m
                
                End If
                
                If EXP(t, i, j).v <= EXP(s, i, j).v Then
                
                sayu.v = sayu.v + 1
                up.v = up.v + EXP(s, i, j).v
                
                End If
                
                
                If EXP(t, i, j).v >= EXP(s, i, j).v Then
                
                sayl.v = sayl.v + 1
                lo.v = lo.v + EXP(s, i, j).v
                
                End If
                
              
            Next s
                
           Karar(t, i, j).l.m = lo.m / sayl.m
           Karar(t, i, j).u.m = up.m / sayu.m
           
           Karar(t, i, j).l.v = lo.v / sayl.v
           Karar(t, i, j).u.v = up.v / sayu.v
        
        Cells(i + (t - 1) * x, 2 * x + 4 + j) = "[" & CStr(Round(Karar(t, i, j).l.m, 3)) & ";" & CStr(Round(Karar(t, i, j).u.m, 3)) & "]" & ";" & "[" & CStr(Round(Karar(t, i, j).l.v, 3)) & ";" & CStr(Round(Karar(t, i, j).u.v, 3)) & "]"
        
        rkarar(i, j).l.m = rkarar(i, j).l.m * Karar(t, i, j).l.m
        rkarar(i, j).u.m = rkarar(i, j).u.m * Karar(t, i, j).u.m
        
        rkarar(i, j).l.v = rkarar(i, j).l.v * Karar(t, i, j).l.v
        rkarar(i, j).u.v = rkarar(i, j).u.v * Karar(t, i, j).u.v
        
        
        Next t
    
    rkarar(i, j).l.m = rkarar(i, j).l.m ^ (1 / z)
    rkarar(i, j).u.m = rkarar(i, j).u.m ^ (1 / z)
    
    rkarar(i, j).l.v = rkarar(i, j).l.v ^ (1 / z)
    rkarar(i, j).u.v = rkarar(i, j).u.v ^ (1 / z)
    
    Cells(i, 3 * x + 6 + j) = "(" & "[" & CStr(Round(rkarar(i, j).l.m, 3)) & ";" & CStr(Round(rkarar(i, j).u.m, 3)) & "]" & ";" & "[" & CStr(Round(rkarar(i, j).l.v, 3)) & ";" & CStr(Round(rkarar(i, j).u.v, 3)) & "]" & ")"
    
    difmat(i, j).l = (rkarar(i, j).l.m) ^ 2 - (rkarar(i, j).u.v) ^ 2
    difmat(i, j).u = (rkarar(i, j).u.m) ^ 2 - (rkarar(i, j).l.v) ^ 2
    
    Cells(i + 1, 4 * x + 7 + j) = "[" & CStr(Round(difmat(i, j).l, 3)) & ";" & CStr(Round(difmat(i, j).u, 3)) & "]"
    
    mulmat(i, j).l = ((1000) ^ difmat(i, j).l) ^ (1 / 2)
    mulmat(i, j).u = ((1000) ^ difmat(i, j).u) ^ (1 / 2)
    
    Cells(i + 1, 5 * x + 9 + j) = "[" & CStr(Round(mulmat(i, j).l, 3)) & ";" & CStr(Round(mulmat(i, j).u, 3)) & "]"
    
    detmat(i, j) = 1 - ((rkarar(i, j).u.m) ^ 2 - (rkarar(i, j).l.m) ^ 2) - ((rkarar(i, j).u.v) ^ 2 - (rkarar(i, j).l.v) ^ 2)
    
    Cells(i + 1, 6 * x + 11 + j) = Round(detmat(i, j), 3)
    
    wemat(i, j) = ((mulmat(i, j).l + mulmat(i, j).u) / 2) * detmat(i, j)
    
    Cells(i + 1, 7 * x + 13 + j) = Round(wemat(i, j), 3)
    
    Weight(i) = Weight(i) + wemat(i, j)
    
    Next j
    Top = Top + Weight(i)
Next i

For i = 1 To x

Weight(i) = Weight(i) / Top

Cells(i + 1, 8 * x + 15) = Weight(i)
Next i



Cells(1, 4 * x + 7) = " The differences matrix"
Cells(1, 5 * x + 9) = " The interval multiplicative matrix"
Cells(1, 6 * x + 11) = "The determinacy value matrix"
Cells(1, 7 * x + 13) = " the matrix of weights"
Cells(1, 8 * x + 15) = "Weights"






End Sub
