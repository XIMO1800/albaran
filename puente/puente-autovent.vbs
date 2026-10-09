' PUENTE AUTOVENT - sube a Google los resumenes de la PDA (.xps) que deja el programa en SERVIDORW10\AUTOVENT.
' Lo lanza una tarea programada cada 2 minutos. Cada fichero subido se mueve a AUTOVENT\ENVIADOS.
' La clave la coge de puente-albtiend.vbs (tiene que estar en la misma carpeta). Apunta lo que hace en envios.log.
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbysLH9zV2Tss0fjnkv9USLkm6HZx7y_sBxEPf9_XG_EzxTCydsWcHPoEPY6uGp51qk8LA/exec"
Const ORIGEN = "\\Servidor-i7\servidor\SERVIDORW10\AUTOVENT"

Dim fso, carpeta, logPath, clave
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
logPath = carpeta & "\envios.log"

Sub Apunta(t)
  On Error Resume Next
  Dim f: Set f = fso.OpenTextFile(logPath, 8, True)
  f.WriteLine Now & "  " & t
  f.Close
End Sub

Function LeeClave()
  Dim t, m, re
  LeeClave = ""
  On Error Resume Next
  t = fso.OpenTextFile(carpeta & "\puente-albtiend.vbs", 1).ReadAll
  Set re = New RegExp: re.Pattern = "Const\s+CLAVE\s*=\s*""([^""]+)"""
  Set m = re.Execute(t)
  If m.Count > 0 Then LeeClave = m(0).SubMatches(0)
End Function

Function Base64(ruta)
  Dim st, x, el
  Set st = CreateObject("ADODB.Stream")
  st.Type = 1: st.Open: st.LoadFromFile ruta
  Set x = CreateObject("MSXML2.DOMDocument.6.0")
  Set el = x.createElement("b")
  el.dataType = "bin.base64"
  el.nodeTypedValue = st.Read
  st.Close
  Base64 = Replace(Replace(el.text, vbLf, ""), vbCr, "")
End Function

Function Sube(ruta, nombre)
  Dim h
  Sube = False
  On Error Resume Next
  Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  h.setTimeouts 15000, 15000, 60000, 120000
  h.open "POST", URL & "?tipo=subirPda&clave=" & clave & "&nombre=" & Replace(nombre, " ", "%20"), False
  h.setRequestHeader "Content-Type", "text/plain"
  h.send Base64(ruta)
  If Err.Number <> 0 Then Apunta "ERROR subiendo " & nombre & ": " & Err.Description: Err.Clear: Exit Function
  If h.status <> 200 Then Apunta "ERROR Google respondio " & h.status & " con " & nombre: Exit Function
  If Left(h.responseText, 2) <> "OK" Then Apunta "ERROR " & nombre & ": " & Left(h.responseText, 200): Exit Function
  Sube = True
End Function

' ---------------- principal ----------------
Dim f, env, dest
clave = LeeClave()
If clave = "" Or clave = "PEGAR_AQUI_LA_CLAVE" Then Apunta "ERROR AUTOVENT: no encuentro la clave en puente-albtiend.vbs": WScript.Quit 1
If Not fso.FolderExists(ORIGEN) Then Apunta "ERROR no encuentro la carpeta " & ORIGEN: WScript.Quit 1
env = ORIGEN & "\ENVIADOS"
On Error Resume Next
If Not fso.FolderExists(env) Then fso.CreateFolder env
On Error GoTo 0
Dim lista, i, nom: lista = Array()
For Each f In fso.GetFolder(ORIGEN).Files
  If (LCase(fso.GetExtensionName(f.Name)) = "xps" Or LCase(fso.GetExtensionName(f.Name)) = "oxps") And DateDiff("s", f.DateLastModified, Now) > 20 Then
    ReDim Preserve lista(UBound(lista) + 1): lista(UBound(lista)) = f.Name
  End If
Next
For i = 0 To UBound(lista)
  nom = lista(i)
  Set f = fso.GetFile(ORIGEN & "\" & nom)
  If True Then
    If True Then
      If Sube(f.Path, f.Name) Then
        dest = env & "\" & f.Name
        On Error Resume Next
        If fso.FileExists(dest) Then fso.DeleteFile dest, True
        fso.MoveFile f.Path, dest
        If Err.Number <> 0 Then Apunta "AVISO " & f.Name & " subido pero no se pudo mover a ENVIADOS: " & Err.Description: Err.Clear Else Apunta "PDA SUBIDO  " & f.Name
        On Error GoTo 0
      End If
    End If
  End If
Next
