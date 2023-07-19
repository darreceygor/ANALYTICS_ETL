VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCuentasDeGastos 
   Caption         =   "CUENTAS DE GASTOS (S_ALR_87013611)"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "usfCuentasDeGastos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfCuentasDeGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEjecutar_Click()

    Call killSAP
    
     Application.Wait (Now + TimeValue("00:00:03"))

    Call llamadasCdG
    
End Sub

Sub llamadasCdG()
    
    Call reporteSAP
    
    Call convertirCSV
    
    Call copiarData
    
    MsgBox "Proceso terminado...", vbExclamation
    usfCuentasDeGastos.Hide

End Sub

Sub killSAP()
 
    Dim kSAP As String

    kSAP = "TASKKILL /F /IM saplogon.exe"
    Shell kSAP, vbHide
End Sub


Private Sub UserForm_Initialize()

    usfCuentasDeGastos.cmbAmbiente.AddItem "CCN QAS RPV"
    usfCuentasDeGastos.cmbAmbiente.AddItem "CCN PRD RPV"
    usfCuentasDeGastos.cmbAmbiente.AddItem "CCN PRD Publica"
    
    usfCuentasDeGastos.cmbSociedad.AddItem "CCN"
    usfCuentasDeGastos.cmbSociedad.AddItem "CSG"
End Sub
