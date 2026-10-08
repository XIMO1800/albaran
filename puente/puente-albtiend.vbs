' PUENTE ALBTIEND - deja en el servidor los albaranes de la tienda (T*.TXT) que emite la app de pedidos.
' Lo lanza una tarea programada cada 5 minutos. No enseña ventanas: todo lo apunta en envios.log (al lado de este fichero).
' Pide a Google los pendientes -> escribe cada uno con nombre temporal -> lo renombra -> solo entonces confirma "entregado".
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbysLH9zV2Tss0fjnkv9USLkm6HZx7y_sBxEPf9_XG_EzxTCydsWcHPoEPY6uGp51qk8LA/exec"
Const CLAVE = "PEGAR_AQUI_LA_CLAVE"                        ' <-- la que da INSTALAR_PUENTE() en el Apps Script
Const DESTINO = "\\Servidor-i7\servidor\SERVIDORW10\ALBTIEND"     ' <-- la misma ruta que funcionó en la prueba

Dim fso, carpeta, logPath, gTxt, gBytes
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
logPath = carpeta & "\envios.log"

Sub Apunta(t)
  On Error Resume Next
  If fso.FileExists(logPath) Then
    If fso.GetFile(logPath).Size > 1000000 Then fso.CopyFile logPath, carpeta & "\envios-anterior.log", True: fso.DeleteFile logPath
  End If
  Dim f: Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & t
  f.Close
End Sub

' Pide algo a Google. Deja el texto en gTxt y los bytes exactos en gBytes. Devuelve True si ha ido bien.
Function Pide(params)
  Dim h
  Pide = False
  On Error Resume Next
  Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  h.setTimeouts 15000, 15000, 30000, 60000
  h.open "GET", URL & "?" & params & "&clave=" & CLAVE & "&r=" & Int(Timer * 100), False
  h.send
  If Err.Number <> 0 Then Apunta "ERROR conexion con Google: " & Err.Description: Err.Clear: Exit Function
  If h.status <> 200 Then Apunta "ERROR Google respondio " & h.status: Exit Function
  gTxt = h.responseText
  gBytes = h.responseBody
  If Left(gTxt, 6) = "ERROR:" Then Apunta "ERROR del Apps Script (" & params & "): " & Left(gTxt, 200): Exit Function
  Pide = True
End Function

Function LeeFichero(ruta)
  On Error Resume Next
  LeeFichero = ""
  If fso.GetFile(ruta).Size = 0 Then Exit Function
  Dim f: Set f = fso.OpenTextFile(ruta, 1)
  LeeFichero = f.ReadAll
  f.Close
End Function

Sub Entrega(nombre)
  Dim final, tmp, st, accion
  If Not Pide("tipo=fichero&nombre=" & nombre) Then Exit Sub
  final = DESTINO & "\" & nombre
  tmp = DESTINO & "\_" & nombre & ".part"
  accion = "ENTREGADO"
  If fso.FileExists(final) Then
    If LeeFichero(final) = gTxt Then
      accion = "YA ESTABA"
    Else
      accion = "REEMPLAZADO (albaran corregido)"
    End If
  End If
  On Error Resume Next
  If accion <> "YA ESTABA" Then
    If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.Write gBytes: st.SaveToFile tmp, 2: st.Close
    If Err.Number <> 0 Then Apunta "ERROR escribiendo " & tmp & ": " & Err.Description: Err.Clear: Exit Sub
    If fso.FileExists(final) Then fso.DeleteFile final, True
    fso.MoveFile tmp, final
    If Err.Number <> 0 Or Not fso.FileExists(final) Then Apunta "ERROR renombrando " & nombre & ": " & Err.Description: Err.Clear: Exit Sub
  End If
  On Error GoTo 0
  If Pide("tipo=entregado&nombre=" & nombre) Then Apunta accion & "  " & nombre Else Apunta "AVISO " & nombre & " esta en ALBTIEND pero no se pudo marcar; se reintentara"
End Sub

' ---------------- principal ----------------
Dim lineas, i, n
If Not fso.FolderExists(DESTINO) Then Apunta "ERROR no encuentro la carpeta " & DESTINO: WScript.Quit 1
If Not Pide("tipo=pendientes") Then WScript.Quit 1
lineas = Split(Replace(gTxt, vbCr, ""), vbLf)
If lineas(0) <> "OK" Then Apunta "ERROR respuesta rara: " & Left(gTxt, 200): WScript.Quit 1
For i = 1 To UBound(lineas)
  n = Trim(lineas(i))
  If n <> "" Then Entrega n
Next
' Para saber de un vistazo que la tarea sigue viva
On Error Resume Next
Dim u: Set u = fso.CreateTextFile(carpeta & "\ultima-vez.txt", True): u.WriteLine Now: u.Close
