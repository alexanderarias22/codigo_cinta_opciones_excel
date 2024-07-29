Attribute VB_Name = "balance_prueba_gral_siigo_nube"

sub Depurar_bp_siigo_nube

' primero elimino las filas de la 1 a la 6
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp

    'ahora lo columna A queda de 10 de ancho
    Columns("A:A").ColumnWidth = 10
    Columns("B:B").ColumnWidth = 5
    Columns("C:C").ColumnWidth = 12
    Columns("D:D").ColumnWidth = 40
    Columns("E:H").ColumnWidth = 17

    'ahora elimino las 3 ultimas filas que no me sirven
    
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Establecer la hoja de trabajo activa
    Set ws = ActiveSheet

    ' Encontrar la última fila con contenido
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Eliminar las últimas tres filas con contenido
    ws.Rows(lastRow - 2 & ":" & lastRow).Delete

    'ahora borramos el contenido de la columna A desde A2 hasta el final
    Range("A2:A" & lastRow).ClearContents

    'ahora eliminamos las filas que en la columna b tienen la palabra "No"
    dim i as Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    for i = lastRow to 2 step -1
        if ws.Cells(i, 2).Value = "No" then
            ws.Rows(i).Delete
        end if
    next i
    
    'ahora insertamos una columna en la columna A
    Columns("A:A").Insert Shift:=xlToRight
    Range("A1").Value = "Clase"
    range("B1").Value = "Grupo"

    'ahora hacemos un ciclo para que en la columna A se ponga los 2 primeros caracteres de la columna C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    for i = 2 to lastRow
        ws.Cells(i, 1).Value = left(ws.Cells(i, 4).Value, 2)
    next i
    'ahora hacemos un ciclo para que en la columna B se ponga los 4 primeros caracteres de la columna C
    for i = 2 to lastRow
        ws.Cells(i, 2).Value = left(ws.Cells(i, 4).Value, 4)
    next i

'ahora inmovilizamos solamente la primera fila que es la de los titulos

    range("A1").Select
'
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

'ahora ponemos filtro a la hoja
    Rows("1:1").Select
    Selection.AutoFilter

  range("A1").Select  

'ahora insertamos una tabla dinámica en una segunda hoja con los datos de la hoja actual

'' ahora si montemos las tablas dinamicas

    Dim wks As Worksheet
    Dim pvc As PivotCache
    Dim pvt As PivotTable
    Dim x As String
    Dim y As String
    Dim z As String
    activeSheet.Name = "datos"
    Sheets("datos").Select
    Range("A1").End(xlToRight).Select
    x = ActiveCell.Column
    Range("A1").End(xlDown).Select
    y = ActiveCell.Row
'
    'Inserta una nueva hoja:
    Set wks = Worksheets.Add
    'Crea el pivot cache
    Set pvc = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:="datos!R1C1:R" & y & "C" & x & "")
    'Crea la tabla dinamica:
    Set pvt = pvc.CreatePivotTable(TableDestination:=wks.Range("A3"), _
    DefaultVersion:=xlPivotTableVersion12)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Montamos los campos de la tabla dinamica:
    With pvt
    'Montamos lo que va por el lado de las filas:
    With .PivotFields("Código cuenta contable")
         .Orientation = xlRowField
         .Position = 1
         .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    With .PivotFields("Nombre cuenta contable")
         .Orientation = xlRowField
         .Position = 2
         .Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    'Montamos los campos de valor:
    'los saldos de los productos:
    With .PivotFields("Saldo inicial")
         .Orientation = xlDataField
         .Position = 1
         .Function = xlSum
         .NumberFormat = "#,##0;[Red]#,##0"
    End With
    With .PivotFields("Movimiento débito")
         .Orientation = xlDataField
         .Position = 2
         .Function = xlSum
         .NumberFormat = "#,##0;[Red]#,##0"
    End With
    With .PivotFields("Movimiento crédito")
         .Orientation = xlDataField
         .Position = 3
         .Function = xlSum
         .NumberFormat = "#,##0;[Red]#,##0"
    End With
    With .PivotFields("Saldo final")
         .Orientation = xlDataField
         .Position = 4
         .Function = xlSum
         .NumberFormat = "#,##0;[Red]#,##0"
    End With
     ' With .PivotFields("Saldo inicial")
     ' .Orientation = xlPageField
     ' .Position = 1
      '.CurrentPage = "Florencia"
    'End With
    
    .RowAxisLayout xlTabularRow
    
    End With
   
    'ponermos el ancho de la columna A a 15
    Columns("A:A").ColumnWidth = 15
    
    

End Sub