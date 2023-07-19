Attribute VB_Name = "mod_balance"
Sub mainBalance()

    usfBalance.Show
    
End Sub




Sub balanceSAP()

    Application.DisplayAlerts = False
    
    '--------------------------------------------
    '       SAP
    '--------------------------------------------
    '       VARIABLES
    '--------------------------------------------
    
    
    
    
    Dim SapGuiAuto As Object, app As Object, session As Object
    
    '---- USUARIO / CONTRASEÑA ------------------------------------------------------
    usr = usfBalance.txtUser.Value
    pwd = usfBalance.txtPass.Value
    
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
            
            
    Dim rutaDestino As String: rutaDestino = CStr(usfBalance.txtCarpetaDestino.Value)
    Dim archivo As String
    
    '---- DATOS CUENTAS MAYOR ------------------------------------------------------
    Dim planDeCuentas As String: planDeCuentas = "PCCN"
    Dim sociedad As String: sociedad = CStr(usfBalance.cmbSociedad.Value)
    
    
    
    '---- OTRAS DELIMITACIONES ------------------------------------------------------
    '---
    '---- DATOS PERIODO ACTUAL ------------------------------------------------------
    Dim anio As String: anio = CStr(usfBalance.txtAnio.Value)
    Dim periodoDe As String: periodoDe = CStr(usfBalance.txtPeriodoDe.Value)
    Dim periodoHasta As String: periodoHasta = CStr(usfBalance.txtPeriodoHasta.Value)
    
        
    '---- DATOS PERIODO COMPARACION ------------------------------------------------------
    Dim anioCmp As String: anioCmp = CStr(usfBalance.txtAnioComp.Value)
    Dim periodoDeCmp As String: periodoDeCmp = CStr(usfBalance.txtPeriodoDeComp.Value)
    Dim periodoHastaCmp As String: periodoHastaCmp = CStr(usfBalance.txtPeriodoHastaComp.Value)
    
    Dim estructuraBalance As String: estructuraBalance = "ZCN1"
    
    '---- DATOS EVALUACIONES ESPECIALES ------------------------------------------------------
    Dim tipoDeBalance As String: tipoDeBalance = "3"
    Dim segmentos As Variant: segmentos = Array("CORPORATIV", "LADRILLO_C", "LADRILLO_O", "OTROS", "PISOS_CBA", "PISOS_OLAV", "TEJAS_OLAV", "VIDRIOS")
    Dim centroBeneficio As Variant: centroBeneficio = Array("CCN", "CERAMIROJA", "EXTRUIDO", "REVESTIMIE", "PORCELANAT")
        
    '---- TRANSACCION ------------------------------------------------------
    Dim trx As String: trx = "S_ALR_87012284"
    
    
    ambiente = CStr(usfBalance.cmbAmbiente.Value)
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
        'session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2").select 'control de salida
        
        
        session.findById(tblDO).Select ' Otras delimitaciones
        
       
        session.findById(tblDOSub & "ctxtBILAVERS").Text = estructuraBalance
        session.findById(tblDOSub & "txtBILBJAHR").Text = anio
        session.findById(tblDOSub & "txtB-MONATE-LOW").Text = periodoDe
        session.findById(tblDOSub & "txtB-MONATE-HIGH").Text = periodoHasta
        session.findById(tblDOSub & "txtBILVJAHR").Text = anioCmp
        session.findById(tblDOSub & "txtV-MONATE-LOW").Text = periodoDeCmp
        session.findById(tblDOSub & "txtV-MONATE-HIGH").Text = periodoHastaCmp
        
        
        '---------------------------------------------------------------------------------------------
        'Evaluaciones especiales
        'Dim tipoDeBalance as String:
 
        
        tblEE = "wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM30"
        tblEESub = tblEE & "/ssub%_SUBSCREEN_TABBL1:RFBILA00:0030/"
        
        
        session.findById(tblEE).Select
        
        session.findById(tblEESub & "ctxtBILABTYP").Text = tipoDeBalance
        session.findById(tblEESub & "chkBILANULL").SetFocus
        session.findById(tblEESub & "chkBILANULL").Selected = True
                
        'ejecutar reporte consolidado
        
        archivo = "consolidado.xls"
        
        session.findById("wnd[0]").sendVKey 8   'F8
        
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = rutaDestino
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = archivo
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
        session.findById("wnd[1]/tbar[0]/btn[11]").press
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[0]/btn[0]").press
                
                
    '--------------------------------------------
    '       SEGMENTACIONES
    '--------------------------------------------
        
        For sg = 0 To 7     'sobre array segmentos

            session.findById("wnd[0]/tbar[0]/btn[3]").press
        
            archivo = "Seg_" & segmentos(sg) & ".xls"
        
            tblSgm = "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN011-LOW"
                
            session.findById("wnd[0]/tbar[1]/btn[16]").press        'delimitaciones opcionales
    
            'elegir segmento
            session.findById(tblSgm).Text = segmentos(sg)
            session.findById(tblSgm).SetFocus
            session.findById(tblSgm).caretPosition = 3
            session.findById("wnd[0]").sendVKey 0   'ENTER
            session.findById("wnd[0]").sendVKey 8   'F8
            
            'guardar archivo de segmento
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
      
            session.findById("wnd[1]/usr/ctxtDY_PATH").Text = rutaDestino
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = archivo
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
            session.findById("wnd[1]/tbar[0]/btn[11]").press
            
        
        Next sg
                  
        
    '--------------------------------------------
    '       CENTRO DE BENEFICIOS
    '--------------------------------------------
                  
        For cb = 0 To 4
        
            archivo = "CeBe_" & centroBeneficio(cb) & ".xls"
        
            tblCeBe = "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN009-LOW"
        
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/tbar[1]/btn[16]").press
            
            'elegir centro de beneficio
            session.findById(tblCeBe).Text = centroBeneficio(cb)
            session.findById(tblCeBe).SetFocus
            session.findById(tblCeBe).caretPosition = 3
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").sendVKey 8
            
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
            session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            
            'guardar centro de beneficio
            session.findById("wnd[1]/usr/ctxtDY_PATH").Text = rutaDestino
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = archivo
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 7
            session.findById("wnd[1]/tbar[0]/btn[11]").press
        
        Next cb
                  
        SAPShell.Run "taskkill /f /im saplogon.exe", 0, True
                  
 End Sub
 


Sub consolidarCeBe()
    
    Application.DisplayAlerts = False
    '--------------------------------------------
    '       VARIABLES
    '--------------------------------------------
    Dim rutaDestino As String: rutaDestino = CStr(usfBalance.txtCarpetaDestino.Value)

    Dim segmentos As Variant: segmentos = Array("CORPORATIV", "LADRILLO_C", "LADRILLO_O", "OTROS", "PISOS_CBA", "PISOS_OLAV", "TEJAS_OLAV", "VIDRIOS")
    Dim centroBeneficio As Variant: centroBeneficio = Array("CCN", "CERAMIROJA", "EXTRUIDO", "REVESTIMIE", "PORCELANAT")

    Dim archivo As String
    Dim hoja As String
    
    '--------------------------------------------------------------------------------
    Dim sociedad As String: sociedad = CStr(usfBalance.cmbSociedad.Value)

    Dim anio As String: anio = CStr(usfBalance.txtAnio.Value)
    Dim periodoDe As String: periodoDe = CStr(usfBalance.txtPeriodoDe.Value)
    Dim periodoHasta As String: periodoHasta = CStr(usfBalance.txtPeriodoHasta.Value)

    Dim consolidadoCeBe As String
    consolidadoCeBe = "PisosOLAVCeBe_" & sociedad & "_" & anio & "(" & periodoDe & "-" & periodoHasta & ").xlsx"
    '--------------------------------------------------------------------------------
    
    
    '--------------------------------------------
    '       CONSOLIDAR Y BORRAR
    '--------------------------------------------
    
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
 
    Windows(consolidadoCeBe).Activate
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
End Sub



Sub consolidarSegmentos()


    Application.DisplayAlerts = False
    '--------------------------------------------
    '       VARIABLES
    '--------------------------------------------
    Dim rutaDestino As String: rutaDestino = CStr(usfBalance.txtCarpetaDestino.Value)

    Dim segmentos As Variant: segmentos = Array("CORPORATIV", "LADRILLO_C", "LADRILLO_O", "OTROS", "PISOS_CBA", "PISOS_OLAV", "TEJAS_OLAV", "VIDRIOS")
    Dim centroBeneficio As Variant: centroBeneficio = Array("CCN", "CERAMIROJA", "EXTRUIDO", "REVESTIMIE", "PORCELANAT")
    
    Dim cons As String: cons = "consolidado"
   
   
   '--------------------------------------------------------------------------------
    Dim sociedad As String: sociedad = CStr(usfBalance.cmbSociedad.Value)

    Dim anio As String: anio = CStr(usfBalance.txtAnio.Value)
    Dim periodoDe As String: periodoDe = CStr(usfBalance.txtPeriodoDe.Value)
    Dim periodoHasta As String: periodoHasta = CStr(usfBalance.txtPeriodoHasta.Value)

    Dim consolidadoSegmentos As String
    consolidadoSegmentos = "consolidadoSEGM_" & sociedad & "_" & anio & "(" & periodoDe & "-" & periodoHasta & ").xlsx"
    '--------------------------------------------------------------------------------
     
     
     
    '--------------------------------------------
    '       FORMULAS DE CONTROL
    '--------------------------------------------
    formulaCtl = "=IFERROR(FIXED(INDEX(Seg_VIDRIOS!C[-4],MATCH(consolidado!RC[-10],Seg_VIDRIOS!C[-10],0))+INDEX(Seg_TEJAS_OLAV!C[-4],MATCH(consolidado!RC[-10],Seg_TEJAS_OLAV!C[-10],0))+INDEX(Seg_PISOS_OLAV!C[-4],MATCH(consolidado!RC[-10],Seg_PISOS_OLAV!C[-10],0))+INDEX(Seg_PISOS_CBA!C[-4],MATCH(consolidado!RC[-10],Seg_PISOS_CBA!C[-10],0))+INDEX(Seg_OTROS!C[-4],MATCH(consolidado!RC[-10],Seg_OTROS!C[-10],0))+INDEX(Seg_LADRILLO_O!C[-4],MATCH(consolidado!RC[-10],Seg_LADRILLO_O!C[-10],0))+INDEX(Seg_LADRILLO_C!C[-4],MATCH(consolidado!RC[-10],Seg_LADRILLO_C!C[-10],0))+INDEX(Seg_CORPORATIV!C[-4],MATCH(consolidado!RC[-10],Seg_CORPORATIV!C[-10],0))),0)"
    formulaADecimal = "=IFERROR(FIXED(CONCAT(""-"",LEFT(TRIM(RC[-7]),FIND(""-"",TRIM(RC[-7]),1)-1))),0)"
    formulaIgual = "=IFERROR(EXACT(RC[-2],RC[-1]),0)"



        '"=IFERROR(RC[-3],0)"
    '--------------------------------------------
    '       CONSOLIDAR Y BORRAR
    '--------------------------------------------

    Workbooks.Add.SaveAs Filename:=consolidadoSegmentos
    
    ChDir rutaDestino
    
    
    'levantar archivos xls y preparar
    For sg = 0 To 7
        archivo = "Seg_" & segmentos(sg) & ".xls"

    Workbooks.OpenText Filename:= _
        rutaDestino & archivo, Origin:= _
        932, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=True, OtherChar:="|", FieldInfo:=Array(Array(1, 1 _
        ), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1)), TrailingMinusNumbers:=True
    Next sg
    
    '------------------------------------------------------------------------------------
    'mover hojas a consolidado
    For sg = 0 To 7
    
        archivo = "Seg_" & segmentos(sg) & ".xls"
        hoja = "Seg_" & segmentos(sg)
        
        
        
        Windows(archivo).Activate
        Sheets(hoja).Select
        Sheets(hoja).Move Before:=Workbooks(consolidadoSegmentos).Sheets(1)
        
        
    Next sg
    
    
        'agregar consolidado
        archivo = cons & ".xls"
        hoja = cons
        
        Workbooks.OpenText Filename:= _
        rutaDestion & archivo, Origin:=932 _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=True, OtherChar:="|", FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
    
        'mover a consolidado de segmentos
        Windows(archivo).Activate
        Sheets(hoja).Select
        Sheets(hoja).Move Before:=Workbooks(consolidadoSegmentos).Sheets(1)
    
        'cerrar consolidado de segmentos
        
        Windows(consolidadoSegmentos).Activate
        Sheets("consolidado").Activate
        
        uf = Trim(CStr(ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row))
        
        'separar codigo de descripcion para buscar en las demas hojas
        Columns("F:H").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("E:E").Select
        Selection.TextToColumns Destination:=Range("E1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(6, 1), Array(8, 1), Array(11, 1)), _
        TrailingMinusNumbers:=True
    
        
        Windows(consolidadoSegmentos).Activate
        Sheets("consolidado").Activate
        Range("O10").Select
        ActiveCell.FormulaR1C1 = formulaCtl
        
        Range("P10").Select
        ActiveCell.FormulaR1C1 = formulaADecimal

        
        Range("Q10").Select
        ActiveCell.FormulaR1C1 = formulaIgual
        
        Range("O10:Q10").Copy Range("O11:Q" & uf)
        
        Columns("Q:Q").EntireColumn.AutoFit
        Columns("O:Q").Select
    
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        '"=INDEX(Seg_VIDRIOS!C[-4],MATCH(consolidado!RC[-10],Seg_VIDRIOS!C[-10],0))"



        Windows(consolidadoSegmentos).Activate
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    
    
End Sub

Sub borrarXLS()

    
    Dim ruta As String: ruta = CStr(usfBalance.txtCarpetaDestino.Value)
    
    Dim cebe As String
    Dim seg As String
    
    cebe = Dir(ruta & "CeBe*.xls")
    seg = Dir(ruta & "Seg*.xls")
    
    Do Until cebe = ""
    
        Kill ruta & cebe
        cebe = Dir(ruta & "CeBe*.xls")
        
    Loop
    
    Do Until seg = ""
        
        Kill ruta & seg
        seg = Dir(ruta & "Seg*.xls")
        
    Loop
    
End Sub
