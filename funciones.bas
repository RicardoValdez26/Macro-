Option VBASupport 1
'------------Luis Ricardo Valdez Pacheco-----------------------
'--------------------------------------------------------------


Sub Eliminar()
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
End Sub
Sub Formato()
'
' Formato Macro
' Formato para la tabla de licencias
'
' Acceso directo: CTRL+h
'
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1:F1").Select
    Range("F1").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("G1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = ""
    Range("G1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Comentario"
    Range("G3").Select
    Columns("G:G").ColumnWidth = 22.29
    Range("G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("F:F").ColumnWidth = 13.43
    Range("E1").Select
    '---------nombre-----------
    ActiveWorkbook.Worksheets("Control").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Control").Sort.SortFields.Add Key:=Range _
        ("E2:E489"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
        
    Range("B1").Select
    Columns("B:B").ColumnWidth = 27.71
    Columns("B:B").ColumnWidth = 35.43
    Columns("C:C").ColumnWidth = 23.43
    Columns("F:F").ColumnWidth = 17.43
    'Contador
    Dim UltLinea As Long
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
    'Colocar margen a todas las celdas
        For i = 1 To UltLinea Step 1
            Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        Next
End Sub
Sub Fechas()
    '------------------Acomodar el contendo de la celda en fecha---------------
    Dim D, M, A As String
    Dim Fecha As String
    Dim FechaFinal As Date
    Dim UltLinea As Long
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
 
    For i = 2 To UltLinea Step 1
    
        D = Left(Range("D" & Trim(Str(i))), 2)
        M = Mid(Range("D" & Trim(Str(i))), 4, 2)
        A = Right(Range("D" & Trim(Str(i))), 4)
    
        Fecha = M & "/" & D + "/" + A
        
        Range("D" & Trim(Str(i))) = Fecha
        
        FechaFinal = DateValue(Range("D" & Trim(Str(i))))
        
        Range("D" & Trim(Str(i))) = FechaFinal
    
    Next

    For i = 2 To UltLinea Step 1
    
        D = Left(Range("E" & Trim(Str(i))), 2)
        M = Mid(Range("E" & Trim(Str(i))), 4, 2)
        A = Right(Range("E" & Trim(Str(i))), 4)
    
        Fecha = M & "/" & D + "/" + A
        
        Range("E" & Trim(Str(i))) = Fecha
        
        FechaFinal = DateValue(Range("E" & Trim(Str(i))))
        
        Range("E" & Trim(Str(i))) = FechaFinal
        
    Next
    
End Sub
Sub Colorear()
'--------Resaltar la Fecha de vencimiento con colores semaforo--------
    
    Dim RangoFechas As Range
    Dim CeldaFecha As Range
    Dim FechaActual As Date
    
    
    FechaActual = Date
    HojaActual = ActiveSheet.Name
    Set currentsheet = ActiveWorkbook.Sheets(HojaActual)
    Set RangoFechas = currentsheet.Range("E2:E999")
    
    
    
    For Each CeldaFecha In RangoFechas
        If Not CeldaFecha = Empty And CeldaFecha.Value < FechaActual Then
            CeldaFecha.Interior.ColorIndex = 6
        End If
    Next
    
    For Each CeldaFecha In RangoFechas
        If Not CeldaFecha = Empty And CeldaFecha.Value > FechaActual Then
            CeldaFecha.Interior.ColorIndex = 4
        End If
    Next

End Sub
Sub OrdenarFechas()

            Range("E2").Sort Key1:=Range("E3"), Order1:=xlAscending, Header:=xlYes
End Sub
Sub Meses()
    Dim UltLinea As Long
    Dim MesVenc As String
    Dim indicador As Integer
    
    indicador = 1
    Dim mes As Date
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To (UltLinea + 1) Step 1
        mes = Range("E" & Trim(Str(i)))
        MesVenc = Month(mes)
        Select Case MesVenc
            Case Is = "1" And indicador = 1
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Enero "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    indicador = 2
            Case Is = "2" And indicador = 2
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Febrero "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 3
            Case Is = "3" And indicador = 3
                 Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Marzo "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 4
            Case Is = "4" And indicador = 4
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Abril"
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 5
            Case Is = "5" And indicador = 5
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Mayo "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 6
            Case Is = "6" And indicador = 6
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Junio "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
               indicador = 7
            Case Is = "7" And indicador = 7
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Julio "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
               indicador = 8
            Case Is = "8" And indicador = 8
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Agosto "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 9
            Case Is = "9" And indicador = 9
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Septiembre "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
               indicador = 10
            Case Is = "10" And indicador = 10
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Octubre "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 11
            Case Is = "11" And indicador = 11
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Noviembre "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 12
            Case Is = "12" And indicador = 12
                Rows(i).EntireRow.Insert
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Range("G" & Trim(Str(i))).Activate
                    With Selection.Interior
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent6
                        .TintAndShade = 0.399975585192419
                        .PatternTintAndShade = 0
                    End With
                    With Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    Selection.Merge
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    ActiveCell.FormulaR1C1 = "Diciembre "
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                    Selection.Font.Bold = True
                    Selection.Font.Size = 18
                    Range("A" & Trim(Str(i)) & ":G" & Trim(Str(i))).Select
                indicador = 1
            End Select
            
    Next
End Sub
Sub Formulario()
'
' Formulario Macro
'

'
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Numero"
    Range("I2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Nombre "
    Range("I3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("J3:M3").Select
    Range("M3").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("K2").Select
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveCell.FormulaR1C1 = "Número ID"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "Número"
    Range("L2:M2").Select
    Range("M2").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "Fecha exp."
    Range("K2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("I5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("K5").Select
    ActiveCell.FormulaR1C1 = "Fecha Venc."
    Range("K5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Range("I2:M8").Select
    Range("I7").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "Comentario"
    Range("J7:M7").Select
    Range("M7").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I2:M10").Select
    Range("L10").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("I8:M10").Select
    Range("M10").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I4:M4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("I6:M6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I8:M10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("M5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("M5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("L5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("I2:M10").Select
    Range("I8").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("I8:M10").Select
    ActiveCell.Offset(2, -3).Range("A1:E1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    ActiveCell.Select
    
    
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    '----------------------Botones------------------------------------
    '------Buscar
    ActiveSheet.Buttons.Add(781.5, 115.5, 72.75, 23.25).Select
    Selection.OnAction = "PERSONAL.XLSB!Buscar"
    Selection.Characters.Text = "Buscar"
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    ActiveCell.Offset(1, -1).Range("A1").Select
    '-----Baja
     ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveSheet.Buttons.Add(990, 114.75, 70, 25).Select
    Selection.OnAction = "PERSONAL.XLSB!baja"
    Selection.Characters.Text = "Baja"
    With Selection.Characters(Start:=1, Length:=4).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    
    '-----Actualizar
      ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveSheet.Buttons.Add(890, 114.75, 70, 25).Select
    Selection.OnAction = "PERSONAL.XLSB!Actualizar"
    Selection.Characters.Text = "Actualizar"
    With Selection.Characters(Start:=1, Length:=4).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
    End With

    Range("I6:M6").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Range("I6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("J6:M6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "Puesto"
    Range("I7").Select
    
End Sub
Sub Buscar()
    
    Dim contador As Long
    Dim UltLinea As Long
    Dim Numero As Variant
    Dim Nombre As Variant
    Dim Puesto As Variant
    Dim Exped As Date
    Dim Vencim As Date
    Dim Id As Variant
    Dim Comentario As Variant
    Dim rango As Variant
    Dim FechaActual As Date
    
    FechaActual = Date
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Range("A2:G" & Trim(Str(UltLinea)))
    Numero = Range("J2")
    
    For contador = 2 To UltLinea
         If Range("A" & Trim(Str(contador))) = Numero Then
            Nombre = Application.VLookup(Numero, rango, 2, False)
            Range("J3") = Nombre
            Puesto = Application.VLookup(Numero, rango, 3, False)
            Range("J6") = Puesto
            Exped = Application.VLookup(Numero, rango, 4, False)
            Range("J5") = Exped
            Vencim = Application.VLookup(Numero, rango, 5, False)
            Range("L5") = Vencim
            Id = Application.VLookup(Numero, rango, 6, False)
            Range("L2") = Id
            Comentario = Application.VLookup(Numero, rango, 7, False)
            Range("J7") = Comentario
        End If
        If Range("L5") > FechaActual Then Range("L5").Interior.ColorIndex = 4
        If Range("L5") < FechaActual Then Range("L5").Interior.ColorIndex = 6
        
    Next
    
      
End Sub
Sub baja()
    Dim contador As Long
    Dim UltLinea As Long
    Dim Numero As Variant
    Dim rango As Variant
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Range("A2:G" & Trim(Str(UltLinea)))
    Numero = Range("J2")
    For contador = 2 To UltLinea
        If Range("A" & Trim(Str(contador))) = Numero Then
            Range("A" & Trim(Str(contador)) & ":G" & Trim(Str(contador))).Interior.Color = 255
        End If
            
    Next

End Sub
Sub Actualizar()
    Dim contador As Long
    Dim UltLinea As Long
    Dim Numero As Variant
    Dim Nombre As Variant
    Dim Puesto As Variant
    Dim Exped As Date
    Dim Vencim As Date
    Dim Id As Variant
    Dim Comentario As Variant
    Dim rango As Variant
    Dim FechaActual As Date
    FechaActual = Date
    
    UltLinea = Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Range("A2:G" & Trim(Str(UltLinea)))
    Numero = Range("J2")
    Nombre = Range("J3")
    Puesto = Range("J6")
    Exped = Range("J5")
    Vencim = Range("L5")
    Id = Range("L2")
    Comentario = Range("J7")
    
    
    For contador = 2 To UltLinea
        If Range("A" & Trim(Str(contador))) = Numero Then
            Range("A" & Trim(Str(contador))) = Numero
            Range("B" & Trim(Str(contador))) = Nombre
            Range("C" & Trim(Str(contador))) = Puesto
            Range("D" & Trim(Str(contador))) = Exped
            Range("E" & Trim(Str(contador))) = Vencim
            Range("F" & Trim(Str(contador))) = Id
            Range("G" & Trim(Str(contador))) = Comentario
             If Range("E" & Trim(Str(contador))) > FechaActual Then Range("L5").Interior.ColorIndex = 4
             If Range("E" & Trim(Str(contador))) < FechaActual Then Range("L5").Interior.ColorIndex = 6
             If Range("E" & Trim(Str(contador))) > FechaActual Then Range("E" & Trim(Str(contador))).Interior.ColorIndex = 4
             If Range("E" & Trim(Str(contador))) < FechaActual Then Range("E" & Trim(Str(contador))).Interior.ColorIndex = 6
              '----------mandar al final de la hoja-------------
              If Range("L5") > Range("E" & Trim(Str(UltLinea - 20))) Then
                Range("A" & Trim(Str(contador)) & ":G" & Trim(Str(contador))).Select
                Range("G" & Trim(Str(contador))).Activate
                Selection.Cut
                ActiveWindow.SmallScroll Down:=519
                Range("A" & Trim(Str(UltLinea + 1)) & ":G" & Trim(Str(UltLinea + 1))).Select
                Range("G" & Trim(Str(UltLinea + 1))).Activate
                ActiveSheet.Paste
                Range("E" & Trim(Str(UltLinea + 1))).Select
                ActiveWindow.SmallScroll Down:=-12
                Range("A" & Trim(Str(contador)) & ":G" & Trim(Str(contador))).Select
                Range("G" & Trim(Str(UltLinea + 1))).Activate
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                    Range("A" & Trim(Str(contador)) & ":G" & Trim(Str(contador))).Select
                    Range("G" & Trim(Str(contador))).Activate
                    Selection.Delete Shift:=xlUp
                '---------bordes--------------
                    Range("A" & Trim(Str(UltLinea + 1)) & ":G" & Trim(Str(UltLinea + 1))).Select
                    Range("G" & Trim(Str(UltLinea + 1))).Activate
                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    With Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Range("C" & Trim(Str(UltLinea + 1))).Select
            End If
       End If
    Next
    
End Sub





