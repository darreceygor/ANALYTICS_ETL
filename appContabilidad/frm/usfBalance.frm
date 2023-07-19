VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfBalance 
   Caption         =   "BALANCE (S_ALR_87012284)"
   ClientHeight    =   8385.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   OleObjectBlob   =   "usfBalance.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usfBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEjecutar_Click()

    Call killSAP
    
     Application.Wait (Now + TimeValue("00:00:03"))
     
    Call llamadasBalance
   
End Sub

Sub llamadasBalance()

    Call balanceSAP
    
    Call consolidarCeBe
    
    Call consolidarSegmentos
    
    Call borrarXLS
    
    MsgBox "Proceso terminado...", vbExclamation
    usfBalance.Hide
    
    
End Sub

Sub killSAP()
 
    Dim kSAP As String

    kSAP = "TASKKILL /F /IM saplogon.exe"
    Shell kSAP, vbHide
End Sub


Private Sub UserForm_Initialize()

    usfBalance.cmbAmbiente.AddItem "CCN QAS RPV"
    usfBalance.cmbAmbiente.AddItem "CCN PRD RPV"
    usfBalance.cmbAmbiente.AddItem "CCN PRD Publica"
    
    usfBalance.cmbSociedad.AddItem "CCN"
    usfBalance.cmbSociedad.AddItem "CSG"
    
End Sub
