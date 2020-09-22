Attribute VB_Name = "Módulo1"



' separarValores()
'   -> separarValorFilaEntera()
'       -> rellenarValorFactura()
'
'
' separarValorUnaColumna()
'   -> rellenarValorFactura()
'
'
' pruebaNumeroATexto()
'
'
' imprimirCaratula()
'   -> prepararRegistros()
'       -> estaEnVector()
'       -> imprimirRangoCaratula()
'           -> caratulaEnWord()
'
'
' imprimirVisto()
'   -> prepararRegistros()
'       -> estaEnVector()
'       -> estaEnVector()imprimirRangoVisto()
'           -> vistoEnWord()
'               -> numeroATexto()
'                   -> convierteCifra()
'
'
' imprimirRemito()
'   -> prepararRegistros()
'       -> estaEnVector()
'       -> imprimirRangoRemito()
'           -> remitoEnExcel()
'
'
' eliminarDuplicados()
'
'



'Separa todos los valores encontrados en una celda seleccionada y un registro por cada valor encontrado
'PRECONDICIONES:
' > Al iniciar se debe ubicar en la primer fila de la columna con los valores a separar
Function separarValores()

    Dim cantReg As Integer
    cantReg = Selection.CurrentRegion.Rows.Count 'Determina cuantos registros (filas) hay en el archivo
    
    Dim i As Integer
    For i = 1 To cantReg
        separarValorFilaEntera ("/")
    Next
    
    mensaje = MsgBox("Listo!" + Chr(13) + " ", vbInformation, "Info")
    
End Function

'Toma el valor de una celda y lo divide usando como delimitador el valor <separador>
'Luego cada valor separado es insertado como nueva fila en la celda inferior al original,
'copiando los valores de las demás columnas de ese registro
Public Function separarValorFilaEntera(ByVal separador As String)

    ' Declaro las Variables
    Dim str As String
    Dim vecStr() As String
    Dim i As Integer
    Dim posCelda As String
    
    ' Inicializo
    str = Selection.Value                        'Comienzo tomando el valor seleccionado
    vecStr = Split(str, separador)               'Separo el string en un vector, usando "/" como delimitador
    rellenarValorFactura vecStr
    i = 0                                        'Inicio la variable para el ciclo while
    posCelda = ActiveCell.Address                'Guarda la posición de la celda activa
    
    ' Ciclo while
    Do While i < UBound(vecStr)         'UBound(vecStr) devuelve el valor máximo que puede tener el índice del vector
        Range(posCelda).Select          'Me ubico en la celda <posCelda>
        Selection.Value = vecStr(i)     'Pega el valor de vecStr(i) en la celda seleccionada
        ActiveCell.EntireRow.Select     'Selecciona toda la fila seleccionada
        Selection.Copy                  'Copia toda la fila
        Range(posCelda).Select          'Vuelvo a parame solo en esa celda
        ActiveCell.Offset(1, 0).Select  'Selecciona la celda que se encuentra debajo de la activa
        posCelda = ActiveCell.Address   'Guarda la posición de la celda activa
        ActiveCell.EntireRow.Select     'Selecciona toda la fila seleccionada
        Selection.Insert Shift:=xlDown  'Inserta la fila copiada anteriomente y desplaza la vieja hacia abajo
        i = i + 1
    Loop
    
    Range(posCelda).Select              'Me ubico en la celda <posCelda>
    Selection.Value = vecStr(i)         'Pega el valor de vecStr(i) en la celda seleccionada
    ActiveCell.Offset(1, 0).Select      'Selecciona la celda que se encuentra debajo de la activa
    
End Function

'Toma el valor de una celda y lo divide usando como delimitador  el valor <separador>
'Luego cada valor separado es insertado en la celda inferior al original, afectando a una sola columna
Public Function separarValorUnaColumna(ByVal separador As String)

    ' Declaro las Variables
    Dim str As String
    Dim vecStr() As String
    Dim i As Integer
    Dim posCelda As String
    
    ' Inicializo
    str = Selection.Value                        'Comienzo tomando el valor seleccionado
    vecStr = Split(str, separador)               'Separo el string en un vector, usando "/" como delimitador
    rellenarValorFactura vecStr
    i = 0                                        'Inicio la variable para el ciclo while

    ' Ciclo while
    Do While i < UBound(vecStr)         'UBound(vecStr) devuelve el valor máximo que puede tener el índice del vector
        Selection.Value = vecStr(i)     'Pega el valor de vecStr(i) en la celda seleccionada
        ActiveCell.Offset(1, 0).Select  'Selecciona la celda que se encuentra debajo de la activa
        Selection.Insert Shift:=xlDown  'Inserta una celda nueva y desplaza la vieja hacia abajo
        i = i + 1
    Loop

    Selection.Value = vecStr(i)         'Pega el valor de vecStr(i) en la celda seleccionada
    ActiveCell.Offset(1, 0).Select      'Selecciona la celda que se encuentra debajo de la activa
    
End Function

'Rellena los valores de numero de factura a 8 dígitos
'Si el primer valor del vector tiene mas dígitos que los siguientes valores, entonces
'Esos valores serán completados con los valores del primer valor
Public Function rellenarValorFactura(ByRef vecStr() As String)
Dim valorPatron As String
Dim valorActual As String
Dim ceros As String
Dim temp As String
Dim i As Integer
Dim largo As Integer

'Le doy formato al valor patrón, con 8 dígitos, completando con ceros
valorPatron = vecStr(0)
ceros = "00000000"
If Len(valorPatron) < 8 Then
    temp = Left(ceros, 8 - Len(valorPatron)) & valorPatron
    valorPatron = temp
    vecStr(0) = valorPatron
End If

i = 1
Do While i <= UBound(vecStr)         'UBound(vecStr) devuelve el valor máximo que puede tener el índice del vector
    temp = ""
    valorActual = vecStr(i)
    If Len(valorActual) < 8 Then
        temp = Left(valorPatron, 8 - Len(valorActual)) & valorActual
    End If
    vecStr(i) = temp
    i = i + 1
Loop
End Function

'Imprime una carátula en Word con los datos de <iniciador> , <nrosFacturas> y de <importe> pasados por parámetro
'La carátula se crea a partir de un documento de Word preexistente, el cual tiene marcadores, el mismo se usa de plantilla
'para la impresión
Public Function caratulaEnWord(ByVal iniciador As String, ByVal nrosFacturas As String, ByVal importeNum As Double, ByVal ordCompra As String, ByVal detalle As String, ByVal destino As String)

    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim importe As String
    Dim facturacion As String

    'Paso el valor double a string
    importe = Format(importeNum, "Standard")
    facturacion = ""

    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True                                                       'Pone visible a la aplicación Word
    Set wrdDoc = wrdApp.Documents.Open(ThisWorkbook.Path & "/" & "caratula.doc") 'Abre un documento existente
    
    'Anexo la orden de compra (Si es que tiene)
    If ordCompra <> "NO/NO" Then
        wrdDoc.Bookmarks("oc").Range.Text = "OC " + ordCompra
    End If
    
    'Acá se genera toda la rutina de guardado
    wrdDoc.Bookmarks("tipoDoc").Range.Text = "Facturación"    'Escribe en el marcador del documento Word
    wrdDoc.Bookmarks("iniciador").Range.Text = iniciador
    If Not (nrosFacturas = "") Then
        facturacion = " Facturación N° " + nrosFacturas
    End If
    If detalle <> "" Then
        detalle = " " + detalle
    End If
    wrdDoc.Bookmarks("tema").Range.Text = facturacion + detalle + "  Importe $ " + importe
    wrdDoc.Bookmarks("destino").Range.Text = destino

    wrdDoc.PrintOut                                         'Imprime desde la impresora estándar
    wrdDoc.Close False                                      'Sale sin guardar los cambios
    
    'wrdApp.ActiveDocument.SaveAs ThisWorkbook.Path & "/" & "SalidaExcel.doc"    'Guarda el documento Word
    'wrdApp.Documents.Close                                                      'Cierra el documento Word
    wrdApp.Quit                                                                 'Cierra la aplicación Word
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

End Function

'Imprime un visto en Word con los datos de <iniciador> , <nrosFacturas> y de <importe> pasados por parámetro
'La carátula se crea a partir de un documento de Word preexistente, el cual tiene marcadores, el mismo se usa de plantilla
'para la impresión
Public Function vistoEnWord(ByVal expediente As String, ByVal iniciador As String, ByVal nrosFacturas As String, ByVal importeNum As Double, ByVal detalle As String)

    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim importe As String
    
    'Paso el valor double a string
    importe = Format(importeNum, "Standard")
    
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True                                                       'Pone visible a la aplicación Word
    Set wrdDoc = wrdApp.Documents.Open(ThisWorkbook.Path & "/" & "visto.doc") 'Abre un documento existente
    
    If Not (detalle = "") Then
        detalle = " " + detalle
    End If
    
    'Acá se genera toda la rutina de guardado
    wrdDoc.Bookmarks("expediente").Range.Text = expediente
    wrdDoc.Bookmarks("empresa").Range.Text = iniciador
    wrdDoc.Bookmarks("nroFacturas").Range.Text = nrosFacturas + detalle
    wrdDoc.Bookmarks("importe").Range.Text = importe                 'Inserto el valor en numeros
    wrdDoc.Bookmarks("enLetras").Range.Text = numeroATexto(importe)
    
    wrdDoc.PrintOut                                             'Imprime desde la impresora estándar
    wrdDoc.Close False                                          'Sale sin guardar los cambios
   
    'wrdApp.ActiveDocument.SaveAs ThisWorkbook.Path & "/" & "salidaVisto.doc"   'Guarda el documento Word
    'wrdApp.Documents.Close                                                     'Cierra el documento Word
    wrdApp.Quit                                                                 'Cierra la aplicación Word
    Set wrdDoc = Nothing
    Set wrdApp = Nothing

End Function

'Imprime un remito en excel con los datos de <iniciador> , <nrosFacturas> y de <importe> pasados por parámetro
'La carátula se crea a partir de un documento de Excel preexistente, el cual tiene marcadores, el mismo se usa de plantilla
'para la impresión
Public Function remitoEnExcel(ByRef iniciador() As String, ByRef nrosFacturas() As String)
' >>> El rango va de la celda G-12 a la G-39 (28 lineas contiguas)
    Dim xlsApp As Excel.Application
    Dim xlsDoc As Excel.Workbook
    Dim texto As String
    Dim cantidadDeFilas As Integer
    Dim largoRenglon As Integer
    Dim MAX_RENGLON As Integer
    Dim i As Integer
    Dim j As Integer
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = True                                                      'Pone visible a la aplicación Word
    MAX_RENGLON = 120
    i = 0                                   'El vector empieza en 0
    cantidadDeFilas = 28
    Do While i <= UBound(iniciador)         'El primer while itera con la cantidad de datos ingredsados en el vector
    
        Set xlsDoc = xlsApp.Workbooks.Open(ThisWorkbook.Path & "/" & "remito.xls") 'Abre un documento existente
        xlsDoc.Activate
        
        j = 0               'Para controlar la cantidad de filas
        
            Do While (j < cantidadDeFilas) And (i <= UBound(iniciador)) 'El segundo while tambien contempla la cantidad de filas del excel
                texto = " + " + iniciador(i) + " " + nrosFacturas(i)
                
                If (Len(texto) > MAX_RENGLON) Then
                    texto = Left(texto, MAX_RENGLON) + "..."
                End If
                
'                '' ***COMIENZO ZONA DE PRUEBAS***
'                '' >>> TODO: Código que trata de repartir los números de facturas en varios renglones
'                '' >>> en el caso de no entrar en uno solo (Código incompleto)
'                '' Para buscar substrings: InStr( [start], string_being_searched, string2, [compare] )
'
'                Dim MAX_RENGLON As Integer
'                Dim indice As Integer
'                Dim renglon As String
'                Dim nrosFactAux As String
'                Dim hayDatos As Boolean
'                Dim auxConEspacio As Boolean
'
'                MAX_RENGLON = 90
'                indice = 0
'                renglon = ""
'                nrosFactAux = ""
'                hayDatos = True
'                auxConEspacio = True
'
'                Do While (hayDatos) ' Ciclo que itera mientras haya datos que agregar en el vector nrosFacturas(i)
'
'                    Do While (auxConEspacio) And (hayDatos) ' Ciclo que itera mientras el renglón a imprimir tenga espacio, y...
'                                                            ' ...mientras haya datos que agregar en el vector nrosFacturas(i)
'                        If Len(nrosFacturas(i)) = 0 Then
'                            hayDatos = False
'                        Else
'                            indice = InStr(indice + 1, nrosFacturas(i), "/")
'                            If (indice = 0) Then
'                                nrosFactAux = iniciador(i) + " " + Left(nrosFacturas(i), indice)
'                                if len(nrosFactAux) <
'                            Else
'
'                            End If
'
'                        End If
'
'
'
'                    Loop
'
'                Loop
'
'                '' ***FIN ZONA DE PRUEBAS***
                
                xlsDoc.Sheets("Hoja1").Range("B12").Offset(j, 0).Value = texto
                i = i + 1
                j = j + 1
            Loop
        
        ' Imprime el documento y luego lo cierra
        xlsDoc.PrintOut                                         'Imprime desde la impresora estándar
        xlsDoc.Close False                                      'Sale sin guardar los cambios
'
        ' Solo para probar cuando no tengo imresora
'        xlsApp.ActiveDocument.SaveAs ThisWorkbook.Path & "/" & "salidaRemito.xls"   'Guarda el documento Excel
'        xlsApp.Documents.Close                                                      'Cierra el documento Excel

    Loop

    xlsApp.Quit                                                                 'Cierra la aplicación Excel
    Set xlsDoc = Nothing
    Set xlsApp = Nothing

End Function

'Recorre todo el rango de filas seleccionado e imprime una carátula por cada conjunto de registros de la misma empresa
'con una misma orden de compra
'Si existen varios registros contiguos de la misma empresa suma los importes, concatena los números de facturas y los
'imprime en una sóla carátula
'PRECONDICIONES:
' > Se debe elegir un rango "sólido". Es decir que no se pueden elegir filas salteadas. Siempre una contigua a la otra.
' > No importa si se seleccionan todas la columnas de las filas o solo una columna cualquiera
Function imprimirRangoCaratula()
    ' variable de tipo Range para hacer referencia a las celdas
    Dim obj_Cell As Range
    Dim iniciador As String
    Dim nroFact As String
    Dim ordCompra As String
    Dim importe As Double
    Dim detalle As String
    Dim destino As String
    Dim ret As Integer
    
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Columns(1).Select        'Elijo solo la primer columna de lo seleccionado
    iniciador = ""
    nroFact = ""
    ordCompra = ""
    detalle = ""
    destino = ""
    importe = 0
    
    'Recorrer todas las celdas seleccionadas en el rango actual
    For Each obj_Cell In Selection.Cells
        With obj_Cell
            If (.Offset(0, 1).Text = iniciador) And (.Offset(0, 5).Text + "/" + .Offset(0, 6).Text = ordCompra) And Not (.Offset(0, 12).Text = "Nuevo Documento") Then
                If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                    nroFact = nroFact + " / " + .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                End If
                importe = importe + .Offset(0, 4).Value
            Else
                If iniciador <> "" Then
                    ret = caratulaEnWord(iniciador, nroFact, importe, ordCompra, detalle, destino)
                End If
                iniciador = .Offset(0, 1).Text
                importe = .Offset(0, 4).Value
                ordCompra = .Offset(0, 5).Text + "/" + .Offset(0, 6).Text
                
                'Si pide no incluir el detalle, lo salteo. Por defecto lo incluyo.
                If Not (.Offset(0, 12).Text = "No Incluir Detalle") Then
                    detalle = .Offset(0, 11).Text
                End If
                
                'Imprime solo cuando tengo un número de factura válido.
                If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                    nroFact = .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                End If
                
                Select Case .Offset(0, 12).Text
                    Case "Rto A Sector B": destino = "Sector B"
                    Case Else: destino = "Sector A"
                End Select

            End If
        End With
    Next
    
    'Imprimo el último proveedor
    ret = caratulaEnWord(iniciador, nroFact, importe, ordCompra, detalle, destino)
    
End Function

'Recorre todo el rango de filas seleccionado e imprime una carátula por cada conjunto de registros de la misma empresa
'con una misma orden de compra
'Si existen varios registros contiguos de la misma empresa suma los importes, concatena los números de facturas y los
'imprime en una sóla carátula
'PRECONDICIONES:
' > Se debe elegir un rango "sólido". Es decir que no se pueden elegir filas salteadas. Siempre una contigua a la otra.
' > No importa si se seleccionan todas la columnas de las filas o solo una columna cualquiera
Function imprimirRangoVisto()
    ' Variable de tipo Range para hacer referencia a las celdas
    Dim obj_Cell As Range
    Dim iniciador As String
    Dim nroFact As String
    Dim expediente As String
    Dim detalle As String
    Dim importe As Double
    Dim ret As Integer
    Dim hayOC As Boolean
    
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Columns(1).Select        'Elijo solo la primer columna de lo seleccionado
    iniciador = ""
    nroFact = ""
    expediente = ""
    detalle = ""
    importe = 0
    hayOC = False
    
    'Recorrer todas las celdas seleccionadas en el rango actual
    For Each obj_Cell In Selection.Cells
    
        If obj_Cell.Offset(0, 5).Text + "/" + obj_Cell.Offset(0, 6).Text = "NO/NO" Then
            
            With obj_Cell
            
                If (.Offset(0, 9).Text + "/" + .Offset(0, 10).Text = expediente) Then
               'If (.Offset(0, 1).Text = iniciador) And (.Offset(0, 9).Text + "/" + .Offset(0, 10).Text = expediente) Then
                    If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                        If (.Offset(0, 1).Text = iniciador) Then
                            nroFact = nroFact + " / " + .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                        Else
                            nroFact = nroFact + " || " + .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                        End If
                    End If
                    
                    'Agrega varios proveedores en el mismo expediente
                    If (.Offset(0, 1).Text <> iniciador) Then
                        iniciador = iniciador + " || " + .Offset(0, 1).Text
                    End If
                    
                    importe = importe + .Offset(0, 4).Value
                Else
                    If iniciador <> "" Then
                        ret = vistoEnWord(expediente, iniciador, nroFact, importe, detalle)
                    End If
                    expediente = .Offset(0, 9).Text + "/" + .Offset(0, 10).Text
                    iniciador = .Offset(0, 1).Text
                    importe = .Offset(0, 4).Value
                    
                    'Si pide no incluir el detalle, lo salteo. Por defecto lo incluyo.
                    If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                        nroFact = .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                    End If
                    
                    'Si pide no incluir el detalle, lo salteo. Por defecto lo incluyo.
                    If Not (.Offset(0, 12).Text = "No Incluir Detalle") Then
                        detalle = .Offset(0, 11).Text
                    End If
                                   
                End If
            End With
    
        Else
            If iniciador <> "" Then
                ret = vistoEnWord(expediente, iniciador, nroFact, importe, detalle)
            End If
            hayOC = True
            iniciador = ""
            nroFact = ""
            expediente = ""
            importe = 0
        End If
    Next

    'Imprimo el último proveedor
    If iniciador <> "" Then
        ret = vistoEnWord(expediente, iniciador, nroFact, importe, detalle)
    End If
    
    If hayOC = True Then
        mensaje = MsgBox("Se encontraron registros para imprimir visto con Orden de Compra. Los registros se omitieron.", vbInformation, "Atención")
    End If

End Function


'Recorre todo el rango de filas seleccionado e imprime un remito por cada cierta cantidad de registros determinada
'por la cantidad de celdas libres que tiene el remito
'Si existen varios registros contiguos de la misma empresa suma los importes, concatena los números de facturas y los
'imprime en una sóla carátula
'PRECONDICIONES:
' > Se debe elegir un rango "sólido". Es decir que no se pueden elegir filas salteadas. Siempre una contigua a la otra.
' > No importa si se seleccionan todas la columnas de las filas o solo una columna cualquiera
Function imprimirRangoRemito()
    ' Variable de tipo Range para hacer referencia a las celdas
    Dim obj_Cell As Range
    Dim iniciador As String
    Dim nroFact As String
    Dim ordCompra As String
    Dim detalle As String
    Dim vecIniciador() As String
    Dim vecNroFact() As String
    Dim ret As Integer
    Dim i As Integer

    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Columns(1).Select                         'Elijo solo la primer columna de lo seleccionado
    iniciador = ""
    nroFact = ""
    ordCompra = ""
    detalle = ""
    i = 0

    'Recorrer todas las celdas seleccionadas en el rango actual
    For Each obj_Cell In Selection.Cells
        With obj_Cell
            If (.Offset(0, 1).Text = iniciador) And (.Offset(0, 5).Text + "/" + .Offset(0, 6).Text = ordCompra) And Not (.Offset(0, 12).Text = "Nuevo Documento") Then
                'Imprime solo cuando tengo un número de factura válido.
                If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                    nroFact = nroFact + " / " + .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                End If
            Else
                If iniciador <> "" Then
                    'Agrego un nuevo dato
                    ReDim Preserve vecIniciador(0 To i)
                    ReDim Preserve vecNroFact(0 To i)
                    vecIniciador(i) = iniciador
                    If detalle <> "" Then
                        detalle = " " + detalle 'Si hay contenido le agrego un espacio para pegarlo con el nroFact
                    End If
                    vecNroFact(i) = nroFact + detalle
                    i = i + 1
                End If
                iniciador = .Offset(0, 1).Text
                ordCompra = .Offset(0, 5).Text + "/" + .Offset(0, 6).Text
                
                detalle = ""
                'Si pide no incluir el detalle, lo salteo. Por defecto lo incluyo.
                If Not (.Offset(0, 12).Text = "No Incluir Detalle") Then
                    detalle = .Offset(0, 11).Text
                End If
                
                'Imprime solo cuando tengo un número de factura válido.
                If Not (.Offset(0, 2).Text + "-" + .Offset(0, 3).Text = "0000-00000000") Then
                    nroFact = .Offset(0, 2).Text + "-" + .Offset(0, 3).Text
                End If
                
            End If
        End With
    Next

    'Agrego el último dato
    ReDim Preserve vecIniciador(0 To i)
    ReDim Preserve vecNroFact(0 To i)
    vecIniciador(i) = iniciador
    vecNroFact(i) = nroFact + detalle

    'Mando a hacer el remito con todo el vector cargado
    ret = remitoEnExcel(vecIniciador, vecNroFact)

End Function

'Prueba de numero a texto
Function pruebaNumeroATexto()
    Dim num1 As String
    Dim letra1 As String
    num1 = "22.023.026,25"
    letra1 = numeroATexto(num1)
    mensaje = MsgBox(num1 + " en letras es: " + letra1, vbInformation, "+<=D")
End Function

'Convierte un número a texto. Toma siempre 2 posiciones decimales.
'Límites: El número mas chico posible es 0,00. El número mas grande posible es 999.999.999,99. No admite números negativos.
Function numeroATexto(Numero)
    Dim texto
    Dim Millones
    Dim Miles
    Dim Cientos
    Dim Decimales
    Dim Cadena
    Dim CadMillones
    Dim CadMiles
    Dim CadCientos
    texto = Numero
    texto = FormatNumber(texto, 2)
    texto = Right(Space(14) & texto, 14)
    Millones = Mid(texto, 1, 3)
    Miles = Mid(texto, 5, 3)
    Cientos = Mid(texto, 9, 3)
    Decimales = Mid(texto, 13, 2)
    CadMillones = convierteCifra(Millones, 1)
    CadMiles = convierteCifra(Miles, 1)
    CadCientos = convierteCifra(Cientos, 0)
    
    'Armo el texto para los millones
    If Trim(CadMillones) > "" Then
        If Trim(CadMillones) = "UN" Then
            Cadena = CadMillones & " MILLÓN"
        Else
            Cadena = CadMillones & " MILLONES"
        End If
    End If
    
    'Armo el texto para los miles
    If Trim(CadMiles) > "" Then
        Cadena = Cadena & " " & CadMiles & " MIL"
    End If
    
    'Armo el texto para los centenas y decimales
    If Trim(CadCientos) > "" Then
        Cadena = Cadena & " " & Trim(CadCientos) & " CON " & Decimales & "/100"
    Else
        If Trim(CadMillones) = "" And Trim(CadMiles) = "" Then
            Cadena = Cadena & Decimales & "/100"
        Else
            Cadena = Cadena & " CON " & Decimales & "/100"
        End If
    End If
    
    numeroATexto = Trim(Cadena)    'Devuelvo el valor obtenido
End Function

'Toma un numero de tres cifras (centena,decena,unidad) y lo convierte en texto.
Function convierteCifra(texto, SW)
    Dim Centena
    Dim Decena
    Dim Unidad
    Dim txtCentena
    Dim txtDecena
    Dim txtUnidad
    Centena = Mid(texto, 1, 1)
    Decena = Mid(texto, 2, 1)
    Unidad = Mid(texto, 3, 1)
    Select Case Centena
        Case "1"
            txtCentena = "CIEN"
            If Decena & Unidad <> "00" Then
                txtCentena = "CIENTO"
            End If
        Case "2"
            txtCentena = "DOSCIENTOS"
        Case "3"
            txtCentena = "TRESCIENTOS"
        Case "4"
            txtCentena = "CUATROCIENTOS"
        Case "5"
            txtCentena = "QUINIENTOS"
        Case "6"
            txtCentena = "SEISCIENTOS"
        Case "7"
            txtCentena = "SETECIENTOS"
        Case "8"
            txtCentena = "OCHOCIENTOS"
        Case "9"
            txtCentena = "NOVECIENTOS"
    End Select

    Select Case Decena
        Case "1"
            txtDecena = "DIEZ"
            Select Case Unidad
                Case "1"
                    txtDecena = "ONCE"
                Case "2"
                    txtDecena = "DOCE"
                Case "3"
                    txtDecena = "TRECE"
                Case "4"
                    txtDecena = "CATORCE"
                Case "5"
                    txtDecena = "QUINCE"
                Case "6"
                    txtDecena = "DIECISÉIS"
                Case "7"
                    txtDecena = "DIECISIETE"
                Case "8"
                    txtDecena = "DIECIOCHO"
                Case "9"
                    txtDecena = "DIECINUEVE"
            End Select
        Case "2"
            txtDecena = "VEINTE"
            If Unidad <> "0" Then
                txtDecena = "VEINTI"
            End If
        Case "3"
            txtDecena = "TREINTA"
            If Unidad <> "0" Then
                txtDecena = "TREINTA Y "
            End If
        Case "4"
            txtDecena = "CUARENTA"
            If Unidad <> "0" Then
                txtDecena = "CUARENTA Y "
            End If
        Case "5"
            txtDecena = "CINCUENTA"
            If Unidad <> "0" Then
                txtDecena = "CINCUENTA Y "
            End If
        Case "6"
            txtDecena = "SESENTA"
            If Unidad <> "0" Then
                txtDecena = "SESENTA Y "
            End If
        Case "7"
            txtDecena = "SETENTA"
            If Unidad <> "0" Then
                txtDecena = "SETENTA Y "
            End If
        Case "8"
            txtDecena = "OCHENTA"
            If Unidad <> "0" Then
                txtDecena = "OCHENTA Y "
            End If
        Case "9"
            txtDecena = "NOVENTA"
            If Unidad <> "0" Then
                txtDecena = "NOVENTA Y "
            End If
    End Select

    If Decena <> "1" Then
        Select Case Unidad
            Case "1"
                If SW Then
                    txtUnidad = "UN"
                Else
                    txtUnidad = "UNO"
                End If
            Case "2"
                If Decena = "2" Then
                    txtUnidad = "DÓS"
                Else
                    txtUnidad = "DOS"
                End If
            Case "3"
                If Decena = "2" Then
                    txtUnidad = "TRÉS"
                Else
                    txtUnidad = "TRES"
                End If
            Case "4"
                txtUnidad = "CUATRO"
            Case "5"
                txtUnidad = "CINCO"
            Case "6"
                If Decena = "2" Then
                    txtUnidad = "SÉIS"
                Else
                    txtUnidad = "SEIS"
                End If
            Case "7"
                txtUnidad = "SIETE"
            Case "8"
                txtUnidad = "OCHO"
            Case "9"
                txtUnidad = "NUEVE"
        End Select
    End If
    convierteCifra = txtCentena & " " & txtDecena & txtUnidad
End Function

Sub imprimirCaratula()
    prepararRegistros ("imprimirCaratula")
End Sub

Sub imprimirVisto()
    prepararRegistros ("imprimirVisto")
End Sub
Sub imprimirRemito()
    prepararRegistros ("imprimirRemito")
End Sub
Function prepararRegistros(ByVal opcion As String)

Dim hojaDatos As String
hojaDatos = ActiveSheet.Name

Application.DisplayAlerts = False
Sheets.Add.Name = "temp0000"
Application.DisplayAlerts = True

Sheets(hojaDatos).Select

Dim vecFilas() As Long
ReDim Preserve vecFilas(0 To 0)
Dim tamanio As Integer
Dim i As Integer

tamanio = 0
i = 0
For Each obj_Cell In Selection.Cells
    If estaEnVector(obj_Cell.Row, vecFilas) = False Then
        ReDim Preserve vecFilas(0 To tamanio)
        tamanio = tamanio + 1
        vecFilas(tamanio - 1) = obj_Cell.Row
        Worksheets(hojaDatos).Range(obj_Cell.End(xlToLeft), obj_Cell.End(xlToRight).End(xlToRight).End(xlToRight)).Copy Worksheets("temp0000").Range("A1").Offset(i, 0)
        i = i + 1
    End If
Next

Sheets("temp0000").Select


Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
If i > 1 Then
    Range(Selection, Selection.End(xlDown)).Select
End If

' Acá hay que llamar al procedimiento que trata los registros
Select Case opcion
    Case "imprimirCaratula": imprimirRangoCaratula
    Case "imprimirVisto": imprimirRangoVisto
    Case "imprimirRemito": imprimirRangoRemito
    Case Else: mensaje = MsgBox("No se eligió una opción válida", vbInformation, "prepararRegistros")
End Select


Sheets(hojaDatos).Select

Application.DisplayAlerts = False
Sheets("temp0000").Delete
Application.DisplayAlerts = True

End Function



'Devuelve true si encuentra el valor <fila> en el el vector <vecFilas>
Function estaEnVector(ByVal fila As Long, ByRef vecFilas() As Long)
    
Dim i As Integer

i = 0   'El vector empieza en 0
Do While i <= UBound(vecFilas)         'UBound(vecStr) devuelve el valor máximo que puede tener el índice del vector
    If vecFilas(i) = fila Then
        estaEnVector = True
        Exit Function
    End If
    i = i + 1
Loop

estaEnVector = False
End Function

'Elimina los registros duplicados ubicados en una columna, los cuales se deben ordenar alfabeticamente de antemano.
'Copia el resultado en la columna inmediata derecha.
Function eliminarDuplicados()
    
    Dim cantReg As Integer
    Dim distanciaColumna As Integer
    Dim i As Integer
    cantReg = Selection.CurrentRegion.Rows.Count 'Determina cuantos registros (filas) hay en el archivo
    distanciaColumna = 2    '<<< CAMBIAR el valor según a la distancia que se encuentren las columnas
    
    
    ActiveCell.Offset(0, distanciaColumna).Value = ActiveCell.Text
    For i = 1 To cantReg
        If (ActiveCell.Offset(i, 0).Text <> ActiveCell.Offset(i - 1, 0).Text) Then
            ActiveCell.Offset(i, distanciaColumna).Value = ActiveCell.Offset(i, 0).Text
        End If
    Next
        
End Function

'----------------------------------------------
'Public Function timeStampDdMmAa() As Date
'    timeStampDdMmAa = Now
'End Function
'
'Public Function timeStampAaaa() As Date
'    timeStampAaaa = Year(Now)
'End Function
'----------------------------------------------


