' PRUEBA DE 2 MINUTOS - ¿puede este ordenador hablar con Google y ver la carpeta ALBTIEND?
' Doble clic. No escribe nada en el servidor ni cambia nada en Google.
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbysLH9zV2Tss0fjnkv9USLkm6HZx7y_sBxEPf9_XG_EzxTCydsWcHPoEPY6uGp51qk8LA/exec"
Const DESTINO = "\\Servidor-i7\servidor\SERVIDORW10\ALBTIEND"   ' <-- CAMBIAR por la ruta real (ver instrucciones)

Dim h, fso, r1, r2
On Error Resume Next
Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
h.setTimeouts 15000, 15000, 30000, 60000
h.open "GET", URL, False
h.send
If Err.Number <> 0 Then
  r1 = "GOOGLE: FALLA" & vbCrLf & Err.Description
  Err.Clear
ElseIf h.status = 200 And InStr(h.responseText, "PEDIDOS TIENDA") > 0 Then
  r1 = "GOOGLE: BIEN"
Else
  r1 = "GOOGLE: FALLA (respuesta " & h.status & ")" & vbCrLf & Left(h.responseText, 200)
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(DESTINO) Then
  r2 = "CARPETA ALBTIEND: BIEN" & vbCrLf & DESTINO
Else
  r2 = "CARPETA ALBTIEND: NO LA ENCUENTRO" & vbCrLf & DESTINO
End If

MsgBox r1 & vbCrLf & vbCrLf & r2, 64, "Prueba del puente"
