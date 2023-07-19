# ETL

CuentasDeGastos.xlsm se conecta a SAP, genera reportes sobre la trx S_ALR_87013611 segun criterios ingresados, los descarga individualmente y los procesa de manera de generar un consolidado en formato xlsx

nBalance.xlsm se conecta a SAP, genera reportes sobre la trx S_ALR_87012284 segun criterios ingresados, los descarga individualmente y los procesa de manera de generar dos consolidados en formato xlsx (uno es sobre los segmentos y el otro sobre los centro de beneficio)

---------------
Scripting SAP
```
Set SapGuiAuto = GetObject("SAPGUI")
       Set app = SapGuiAuto.GetScriptingEngine
       Set Connection = app.OpenConnection(ambiente, True)
       Set session = Connection.Children(0)

        
        'iniciar sesion
            'session.findById("wnd[0]").maximize
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = pass
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
            session.findById("wnd[0]").sendVKey 0
        
            Application.Wait (Now + TimeValue("00:00:03"))
            
        session.StartTransaction (trx)
                

        'Seleccion cuenta Mayor
        session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").Text = planDeCuentas
        session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").Text = sociedad
        
        '---------------------------------------------------------------------------------------------
        'Otras delimitaciones
        tblDO = "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1"
        
        tblDOSub = tblDO & "/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/"
        'session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM30").select ' evaluaciones especiales
 ```


VBS



VBA

 ```
Workbooks.Add.SaveAs Filename:=consolidadoCeBe
    
    ChDir rutaDestino
    
    
    'levantar archivos xls y preparar
    For cb = 0 To 4
        archivo = "CeBe_" & centroBeneficio(cb) & ".xls"
        
        Workbooks.OpenText Filename:= _
            rutaDsetino & archivo, Origin:=932, _
            StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
            , Space:=False, Other:=True, OtherChar:="|", FieldInfo:=Array(Array(1, 1 _
            ), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
            Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
            , 1), Array(16, 1)), TrailingMinusNumbers:=True
    Next cb
    
    '------------------------------------------------------------------------------------
    'mover hojas a consolidado
    For cb = 0 To 4
    
        archivo = "CeBe_" & centroBeneficio(cb) & ".xls"
        hoja = "CeBe_" & centroBeneficio(cb)
        
        Windows(archivo).Activate
        Sheets(hoja).Select
        Sheets(hoja).Move Before:=Workbooks(consolidadoCeBe).Sheets(1)
     
    Next cb
 ``` 
    Windows(consolidadoCeBe).Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close
'''
