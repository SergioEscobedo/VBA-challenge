Attribute VB_Name = "Module1"
Sub Ticker()
    Dim datosA() As Variant
    Dim datosI() As Variant
    Dim i As Long, j As Long, k As Long
    Dim encontrado As Boolean
    
    datosA = Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row).Value
    
    ReDim datosI(1 To UBound(datosA, 1), 1 To 1)
    
    ' Repeated data
    k = 1
    datosI(k, 1) = datosA(1, 1)
    
    For i = 2 To UBound(datosA, 1)
        encontrado = False
        
        For j = 1 To k
            If datosA(i, 1) = datosI(j, 1) Then
                encontrado = True
                Exit For
            End If
        Next j
        
        If Not encontrado Then
            k = k + 1
            datosI(k, 1) = datosA(i, 1)
        End If
    Next i
    
    Range("I1:I" & k).Value = datosI
    
End Sub

Sub YearlyChange()
    Dim filaApertura As Long
    Dim filaCierre As Long
    Dim filaResultado As Long
    
    filaApertura = 2
    filaCierre = 252
    filaResultado = 2
    
    For i = 1 To 3000
        Dim precioApertura As Double
        Dim precioCierre As Double
        Dim cambioAnual As Double
        
        precioApertura = Cells(filaApertura, 3).Value
        precioCierre = Cells(filaCierre, 6).Value
        
        cambioAnual = precioCierre - precioApertura
        
        Cells(filaResultado, 10).Value = cambioAnual
        
        filaApertura = filaApertura + 251
        filaCierre = filaCierre + 251
        filaResultado = filaResultado + 1
    Next i

End Sub

Sub percentChange()
    Dim filaApertura As Long
    Dim filaCierre As Long
    Dim filaResultado As Long
    
    filaApertura = 2
    filaCierre = 252
    filaResultado = 2
    
    For i = 1 To 3000
        Dim precioApertura As Double
        Dim precioCierre As Double
        Dim percentChange As Double
        

        precioApertura = Cells(filaApertura, 3).Value
        precioCierre = Cells(filaCierre, 6).Value
        
        percentChange = ((precioCierre - precioApertura) / precioApertura) * 100

        Cells(filaResultado, 11).Value = percentChange
        
        filaApertura = filaApertura + 251
        filaCierre = filaCierre + 251
        filaResultado = filaResultado + 1
    Next i
End Sub

Sub TotalStock()
    Dim filaInicial As Long
    Dim filaResultado As Long
    Dim i As Long
    
    filaInicial = 2 '
    filaResultado = 2
    
    For i = 1 To 3000
        Dim suma As Double
        suma = 0

        Dim fila As Long
        For fila = filaInicial To filaInicial + 250
            suma = suma + Cells(fila, 7).Value
        Next fila
        
        Cells(filaResultado, 12).Value = suma
        

        filaInicial = filaInicial + 251
        filaResultado = filaResultado + 1
    Next i
End Sub

Sub MaxMin()
    Dim maxIndexK As Long
    Dim minIndexK As Long
    Dim maxIndexL As Long
    

    maxIndexK = Application.Match(WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
    minIndexK = Application.Match(WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)

    maxIndexL = Application.Match(WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
    Range("Q2").Value = WorksheetFunction.Max(Range("K:K"))
    
    Range("Q2").Offset(0, 1).Value = Cells(maxIndexK, 9).Value
    Range("Q3").Value = WorksheetFunction.Min(Range("K:K"))

    Range("Q3").Offset(0, 1).Value = Cells(minIndexK, 9).Value
    Range("Q4").Value = WorksheetFunction.Max(Range("L:L"))
    
    Range("Q4").Offset(0, 1).Value = Cells(maxIndexL, 9).Value
End Sub

Sub EjecutarEnHoja(hoja As String)
    Sheets(hoja).Activate
    Ticker
    YearlyChange
    percentChange
    TotalStock
    MaxMin
End Sub

Sub EjecutarEnTodasLasHojas()
    EjecutarEnHoja "2018"
    EjecutarEnHoja "2019"
    EjecutarEnHoja "2020"
End Sub

