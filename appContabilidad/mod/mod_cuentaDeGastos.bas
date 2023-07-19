Attribute VB_Name = "mod_cuentaDeGastos"


Sub mainCuentasDeGastos()

    usfCuentasDeGastos.Show

    
End Sub



Sub reporteSAP()

    Application.DisplayAlerts = False
    
    '--------------------------------------------
    '       SAP
    '--------------------------------------------
    '       VARIABLES
    '--------------------------------------------
    usr = CStr(usfCuentasDeGastos.txtUser.Value)
    pwd = CStr(usfCuentasDeGastos.txtPass.Value)
    
    
    Dim SapGuiAuto As Object, app As Object, session As Object
    
    Dim usuario As String
            If usr = "" Then
                usuario = "vdigesu"
            Else
                usuario = usr
            End If
    
    Dim pass As String
            If pwd = "" Then
                pass = "rarolon0"
            Else
                pass = pwd
            End If
            
            
    Dim rutaDestino As String: rutaDestino = CStr(usfCuentasDeGastos.txtCarpetaDestino.Value)
    Dim archivo As String
    
    Dim sociedadCO As String: sociedadCO = CStr(usfCuentasDeGastos.cmbSociedad.Value)
    Dim ejercicio As String: ejercicio = CStr(usfCuentasDeGastos.txtEjercicio.Value)
    Dim mesInicio As String: mesInicio = CStr(usfCuentasDeGastos.txtMesInicio.Value)
    Dim mesFinal As String: mesFinal = CStr(usfCuentasDeGastos.txtMesFinal.Value)
    Dim cuentaDe As String: cuentaDe = CStr(usfCuentasDeGastos.txtCuentaDe.Value)
    Dim cuentaHasta As String: cuentaHasta = CStr(usfCuentasDeGastos.txtCuentaHasta.Value)
    
    If ruta = "" Then
        ruta = sociedadCO & "_" & ejercicio & "(de_" & cuentaDe & "_a_" & cuentaHasta & ")"
    End If

    
    Dim trx As String: trx = "S_ALR_87013611"
    
    
    ambiente = CStr(usfCuentasDeGastos.cmbAmbiente.Value)
                'PRD = "CCN PRD"
                'QAS = "CCN QAS"
    
     
     
    
    '--------------------------------------------
    '       CONEXION
    '--------------------------------------------
    
    Set SAPShell = CreateObject("WScript.Shell")
    SAPShell.Run ("""C:\Program Files\sap\FrontEnd\SAPgui\saplogon.exe""")
   
      SAP_BIN = "saplogon.exe"
      SAP_GUI_PATH = "C:\Program Files\sap\FrontEnd\SAPgui\" & SAP_BIN
   

  
    
    Application.Wait (Now + TimeValue("00:00:03"))
    
      
       Set SapGuiAuto = GetObject("SAPGUI")
           Application.Wait (Now + TimeValue("00:00:03"))
       Set app = SapGuiAuto.GetScriptingEngine
           Application.Wait (Now + TimeValue("00:00:03"))
       Set Connection = app.OpenConnection(ambiente, True)
           Application.Wait (Now + TimeValue("00:00:03"))
       Set session = Connection.Children(0)
           Application.Wait (Now + TimeValue("00:00:03"))


'iniciar sesion
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = pass
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
    session.findById("wnd[0]").sendVKey 0

    Application.Wait (Now + TimeValue("00:00:03"))
    
   For i = CInt(mesInicio) To CInt(mesFinal)
    
        session.StartTransaction (trx)
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxt$1KOKRE").Text = sociedadCO
        session.findById("wnd[0]/usr/txt$1GJAHR").Text = ejercicio
        session.findById("wnd[0]/usr/ctxt$1PERIV").Text = i
        session.findById("wnd[0]/usr/ctxt$1PERIB").Text = i
        session.findById("wnd[0]/usr/ctxt$1VERP").Text = "0"
        session.findById("wnd[0]/usr/ctxt_1KOSET-LOW").Text = cuentaDe
        session.findById("wnd[0]/usr/ctxt_1KOSET-HIGH").Text = cuentaHasta
        session.findById("wnd[0]/usr/ctxt_1KOSET-HIGH").SetFocus
        session.findById("wnd[0]/usr/ctxt_1KOSET-HIGH").caretPosition = 7
        session.findById("wnd[0]").sendVKey 8
        
        
        '------------------------------------------------------------------
        ' GUARDAR REPORTE
        
        '--------------------------------------
        '   Generar nombre de archivo
        
        If i < 10 Then
            archivo = sociedadCO & "_" & ejercicio & "_0" & i & ".csv"
        Else
            archivo = sociedadCO & "_" & ejercicio & "_" & i & ".csv"
        End If
        
        session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[1]").Select
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = rutaDestino
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
        Application.Wait (Now + TimeValue("00:00:02"))
    Next i
     
        Application.Wait (Now + TimeValue("00:00:02"))
   
    SAPShell.Run "taskkill /f /im saplogon.exe", 0, True
End Sub


Sub convertirCSV()

    Application.DisplayAlerts = False
    
    Dim base As String: base = ActiveWorkbook.Name
    Dim ruta As String: ruta = CStr(usfCuentasDeGastos.txtCarpetaDestino.Value)
    Dim arrArchivos() As Variant: arrArchivos() = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    Dim i As Integer
    
    Workbooks(base).Activate
    Worksheets(1).Activate
    

    Dim sociedadCO As String: sociedadCO = CStr(usfCuentasDeGastos.cmbSociedad.Value)
    Dim ejercicio As String: ejercicio = CStr(usfCuentasDeGastos.txtEjercicio.Value)
    Dim mesInicio As String: mesInicio = CStr(usfCuentasDeGastos.txtMesInicio.Value)
    Dim mesFinal As String: mesFinal = CStr(usfCuentasDeGastos.txtMesFinal.Value)
    Dim cuentaDe As String: cuentaDe = CStr(usfCuentasDeGastos.txtCuentaDe.Value)
    Dim cuentaHasta As String: cuentaHasta = CStr(usfCuentasDeGastos.txtCuentaHasta.Value)
       
       
    If ruta = "" Then
        ruta = sociedadCO & "_" & ejercicio & "(de_" & cuentaDe & "_a_" & cuentaHasta & ")"
    End If
        
    'Para procesar los archivos csv

    For i = CInt(mesInicio) - 1 To CInt(mesFinal) - 1
         nombre = sociedadCO & "_" & ejercicio & "_" & arrArchivos(i)
                
         Dim archivo As String: archivo = ruta & nombre & ".csv"
         Dim salida As String: salida = nombre & ".xlsx"
         Dim hoja As String: hoja = sociedadCO & "_" & ejercicio & "_" & arrArchivos(i)
         
         '----------------------------------------------------------------------------------
         '   ABRIR ARCHIVO CSV POR CADA MES
         '----------------------------------------------------------------------------------
         
         '   Procesar datos en columnas
         '-------------------------------------------
         Workbooks.Open Filename:= _
             archivo
         Columns("A:A").Select
         Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
             TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
             Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
             :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
             TrailingMinusNumbers:=True
             
         '   Titulos
         '--------------------------------------------------------------------------------
         Range("B1").Select
         ActiveCell.FormulaR1C1 = "Clases de Coste"
         Range("C1").Select
         Columns("B:B").EntireColumn.AutoFit
         ActiveCell.FormulaR1C1 = "Cst.reales"
         Range("D1").Select
         ActiveCell.FormulaR1C1 = "Mes"
         Columns("B:D").Select
         
         
         uf = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
         
         '   en todas las filas coloco el mes correspondiente
         '--------------------------------------------------------------------------------
         Range("D2").Value = arrArchivos(i)
         Range("D2").Copy: Range("D3:D" & uf).PasteSpecial xlPasteValues
         
         
         
         ActiveWorkbook.Worksheets(hoja).Sort.SortFields.Clear
         ActiveWorkbook.Worksheets(hoja).Sort.SortFields.Add2 Key:=Range( _
             "B2:B" & uf), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
             xlSortNormal
             
         With ActiveWorkbook.Worksheets(hoja).Sort
             .SetRange Range("B1:D" & uf)
             .Header = xlYes
             .MatchCase = False
             .Orientation = xlTopToBottom
             .SortMethod = xlPinYin
             .Apply
         End With
         
         
         '   Selecciono A1 y coloco formula ESPACIOS(B1)
         '--------------------------------------------------------------------------------
         Range("A1").Select
         ActiveCell.FormulaR1C1 = "=TRIM(RC[1])"
         Range("A1").Copy
            
         '   copio formula hasta ultima fila
         '--------------------------------------------------------------------------------
         
         Range("A1").Copy: Range("A2:A" & uf).PasteSpecial xlPasteFormulas
         
           
         '   copio columna A y pego valores en columna B
         '--------------------------------------------------------------------------------
         
         Range("A:A").Copy: Range("B:B").PasteSpecial xlPasteValues
         
         
         '   borrar datos extras de columnas
         '--------------------------------------------------------------------------------
         Columns("A:A").Select
         Selection.ClearContents
         Columns("E:BE").Select
         Selection.ClearContents
         ActiveWindow.ScrollColumn = 2
         ActiveWindow.ScrollColumn = 1
         
         
         '   dejar solo las cuentas, valores y mes
         '--------------------------------------------------------------------------------
         Range("A1").Select
         ActiveCell.FormulaR1C1 = "=MID(RC[1],1,1)"      'Extrae primer caracter del texto en columna B
         
         
         Range("A1").Copy: Range("A2:A" & uf).PasteSpecial xlPasteFormulas
         
         Selection.AutoFilter
         ActiveSheet.Range("$A$1:$F$114").AutoFilter Field:=1, Criteria1:=Array("-", _
             "*", "A", "C", "D", "E", "G", "I", "P", "R", "S", "V", "="), Operator:=xlFilterValues
         Rows("2:250").Select
         Selection.Delete Shift:=xlUp
         Selection.AutoFilter
         Range("B2").Select
         Columns("A:A").Select
         Selection.ClearContents
         Selection.AutoFilter
            
         ChDir ruta
       
         ActiveWorkbook.SaveAs Filename:= _
             salida, FileFormat:= _
             xlOpenXMLWorkbook, CreateBackup:=False
             
         ActiveWorkbook.Close
       
       Next i
       
       '--------------------------------------------------------------------------------
       '   MOVER CSV A OTRA CARPETA
       '--------------------------------------------------------------------------------
       'Dim carpeta As String: carpeta = sociedadCO & "-" & ejercicio & "-CSV"
       'Dim archivoCSV As String: archivoCSV = Dir(ruta & "*.csv")
       'Dim nuevaRuta As String: nuevaRuta = ruta & carpeta
        '
        ' MkDir nuevaRuta       ' crear nueva carpeta
        '
        ' If archivoCSV = "" Then
        '     MsgBox "No hay archivos a mover.", vbExclamation, "IT"
        ' Else
        '     Do Until archivoCSV = ""
        '         Name ruta & archivoCSV As nuevaRuta & "\" & archivoCSV
        '         archivoCSV = Dir
        '     Loop
        ' End If
       
       Kill ruta & "*.csv"

End Sub



Sub copiarData()

    Application.DisplayAlerts = False

    Dim base As String: base = ActiveWorkbook.Name
    
    Dim ruta As String: ruta = CStr(usfCuentasDeGastos.txtCarpetaDestino.Value)
    Dim archivo As String: archivo = "consolidado.xlsx"
    Dim nombreHoja As String: nombreHoja = "Resumen"
    
    '-----------------------------------------------------------------
    
    Workbooks(base).Activate
    Worksheets(1).Activate
    
    Dim sociedadCO As String: sociedadCO = CStr(usfCuentasDeGastos.cmbSociedad.Value)
    Dim ejercicio As String: ejercicio = CStr(usfCuentasDeGastos.txtEjercicio.Value)
    Dim mesInicio As String: mesInicio = CStr(usfCuentasDeGastos.txtMesInicio.Value)
    Dim mesFinal As String: mesFinal = CStr(usfCuentasDeGastos.txtMesFinal.Value)
    Dim cuentaDe As String: cuentaDe = CStr(usfCuentasDeGastos.txtCuentaDe.Value)
    Dim cuentaHasta As String: cuentaHasta = CStr(usfCuentasDeGastos.txtCuentaHasta.Value)

    If ruta = "" Then
        ruta = sociedadCO & "_" & ejercicio & "(de_" & cuentaDe & "_a_" & cuentaHasta & ")"
    End If

    '-----------------------------------------------------------------
    ' nuevo libro
    '-----------------------------------------------------------------
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:= _
        ruta & archivo, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    
    Windows(archivo).Activate
    Sheets("Hoja1").Name = "Resumen"        'cambio nombre a hoja
      
    '-----------------------------------------------------------------
    ' Vuelvo a Script para procesar los archivos
    '-----------------------------------------------------------------
    
    For i = CInt(mesInicio) - 1 To CInt(mesFinal) - 1
    
        arrArchivos = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
        arch = sociedadCO & "_" & ejercicio & "_" & arrArchivos(i) & ".xlsx"
        procesar = ruta & arch
        
    '-----------------------------------------------------------------------------------------------------------------
        Windows(base).Activate
        Workbooks.Open Filename:=procesar
        ufPr = Trim(CStr(ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row))
                
        Range("B1:D" & ufPr).Select
        Selection.Copy
    '-----------------------------------------------------------------------------------------------------------------
    Debug.Print (archivo)
    
        Workbooks(archivo).Activate
        Sheets(nombreHoja).Activate
        
        uf = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
        
        If uf = 1 Then
            uf = ufPr
            Range("A1:A" & uf).Select
            ActiveSheet.Paste
        Else
            uf = uf + 1
            Range("A" & uf & ":A" & ufPr + uf).Select
            ActiveSheet.Paste
        End If
        
        Application.CutCopyMode = False
    
        Workbooks(arch).Activate
        
        ActiveWorkbook.Close SaveChanges:=False

    Next i
    
    
    '-----------------------------------------------------------------------------------------------------------------
    '       CONSOLIDADO
    '-----------------------------------------------------------------------------------------------------------------
    
    Windows(archivo).Activate
    Sheets(nombreHoja).Activate
    
    Dim nmbTabla As String: nmbTabla = "tblData"        'nombre de data
    Dim hojaTD As String: hojaTD = "Tabla Resumen"      'nombre hoja de la tabla dinamica
    Dim dataTD As String: dataTD = "TablaDinamica"      'nombre de la tabla dinamica

    
    Dim pt As PivotTable
    Dim pCache As PivotCache
    
    
    'selecciono el rango para tabla
    '-----------------------------------------------------------------------------------------------------------------
    ufC = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    
    ActiveWorkbook.Sheets(nombreHoja).ListObjects.Add(xlSrcRange, Range("A1:C" & ufC), , xlYes).Name = nmbTabla
    
    'creao tabla dinamica
    '-----------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Worksheets(hojaTD).Delete      'elimino si existe
    
    Worksheets.Add(Before:=ActiveSheet).Name = hojaTD     'creo hoja para tabla dinamica
    

    Set pCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=nmbTabla)

    Sheets(hojaTD).Activate
    
    Set pt = ActiveSheet.PivotTables.Add( _
        PivotCache:=pCache, _
        tabledestination:=Range("A1"))
    
    
    'Columnas
    With pt.PivotFields("Clases de coste")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    'Filas
    With pt.PivotFields("Mes")
        .Orientation = xlColumnField
        .Position = 1
    End With
       
    'Valores
    With pt.PivotFields("Cst.reales")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "#,##0"
    End With
               
    Range("A1").Value = sociedadCO & "-" & ejercicio & "-" & cuentaDe & "-" & cuentaHasta
    Range("A2").Value = "Clases de coste"
    Range("B1").Value = "Mes"


    Workbooks(archivo).Close SaveChanges:=True
 
    
    '--------------------------------------------------------------------------------
    '   MOVER XLSX (meses) A OTRA CARPETA
    '--------------------------------------------------------------------------------
    ' arrArchivos = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    '--------------------------------------------------------------------------------
    
    'Dim carpetaXLSX As String: carpetaXLSX = sociedadCO & "-" & ejercicio & "-Meses"
    'Dim nuevaRuta As String: nuevaRuta = ruta & carpetaXLSX
    
    'MkDir nuevaRuta
    
    For i = CInt(mesInicio) - 1 To CInt(mesFinal) - 1
    
        seArc = sociedadCO & "_" & ejercicio & "_" & arrArchivos(i) & ".xlsx"
    
        Kill ruta & seArc
        
        

    
        'FileCopy ruta & seArc, nuevaRuta & "\" & seArc
        'Kill ruta & seArc
    Next i
  
    Name ruta & archivo As _
            ruta & "Consolidado_" & sociedadCO & "_" & ejercicio & "_(" & mesInicio & "-" & mesFinal & ")_(" & cuentaDe & "-" & cuentaHasta & ").xlsx"
    

End Sub




