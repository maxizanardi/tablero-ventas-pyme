// Limpieza equivalente en Power Query (Excel o Power BI)
// Origen: "Online Retail.xlsx" de UCI (https://archive.ics.uci.edu/dataset/352/online+retail)
// Misma lógica que analisis.py, para quien quiera rehacer el tablero dentro de Excel o Power BI.
let
    Origen = Excel.Workbook(File.Contents("C:\datos\Online Retail.xlsx"), null, true),
    Hoja = Origen{[Item = "Online Retail", Kind = "Sheet"]}[Data],
    Encabezados = Table.PromoteHeaders(Hoja, [PromoteAllScalars = true]),
    Tipos = Table.TransformColumnTypes(Encabezados, {
        {"InvoiceNo", type text}, {"StockCode", type text}, {"Description", type text},
        {"Quantity", Int64.Type}, {"InvoiceDate", type datetime}, {"UnitPrice", type number},
        {"CustomerID", type text}, {"Country", type text}}),

    // 1. Duplicados exactos
    SinDuplicados = Table.Distinct(Tipos),

    // 2. Doce meses completos (diciembre 2011 está incompleto)
    Periodo = Table.SelectRows(SinDuplicados, each [InvoiceDate] >= #datetime(2010, 12, 1, 0, 0, 0)
                                                and [InvoiceDate] <  #datetime(2011, 12, 1, 0, 0, 0)),

    // 3. Códigos que no son productos (portes, comisiones, ajustes manuales)
    NoProducto = {"POST", "D", "M", "BANK CHARGES", "AMAZONFEE", "DOT", "CRUK", "S", "PADS", "B", "C2"},
    SoloProductos = Table.SelectRows(Periodo, each not List.Contains(NoProducto, Text.Upper([StockCode]))
                                                and not Text.StartsWith([StockCode], "gift_")),

    // 4. Separar devoluciones (facturas que empiezan con "C")
    EsDevolucion = Table.AddColumn(SoloProductos, "Devolucion", each Text.StartsWith([InvoiceNo], "C"), type logical),
    Ventas = Table.SelectRows(EsDevolucion, each not [Devolucion] and [Quantity] > 0 and [UnitPrice] > 0),

    // 5. Columnas para el análisis
    Importe = Table.AddColumn(Ventas, "Importe", each [Quantity] * [UnitPrice], type number),
    Mes = Table.AddColumn(Importe, "Mes", each Date.StartOfMonth(DateTime.Date([InvoiceDate])), type date),
    DiaSemana = Table.AddColumn(Mes, "DiaSemana", each Date.DayOfWeek(DateTime.Date([InvoiceDate]), Day.Monday) + 1, Int64.Type),
    Hora = Table.AddColumn(DiaSemana, "Hora", each Time.Hour(DateTime.Time([InvoiceDate])), Int64.Type),
    ClienteLimpio = Table.ReplaceValue(Hora, null, "(sin dato)", Replacer.ReplaceValue, {"CustomerID"})
in
    ClienteLimpio
