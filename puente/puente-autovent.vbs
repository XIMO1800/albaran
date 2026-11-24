' PUENTE AUTOVENT - sube a Google los resumenes de la PDA (.xps) que deja el programa en SERVIDORW10\AUTOVENT,
' y los datos de la PDA (cab01/lin01/cob01, cab02/lin02/cob02, CLIENTE.TXT, ARTICULO.TXT) para la revision de albaranes.
' Los cab/lin/cob se suben con la ruta y la fecha y hora del fichero en el nombre (R02_20261120_1741_cab.txt) para que no se machaquen.
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

Function Dos(n)
  Dos = Right("0" & n, 2)
End Function

' Nombre con el que se sube: los .xps tal cual; cab02.txt -> R02_AAAAMMDD_HHMM_cab.txt; CLIENTE/ARTICULO tal cual. "" = no se sube
Function NombreSubida(f)
  Dim ext, re, m, d
  NombreSubida = ""
  ext = LCase(fso.GetExtensionName(f.Name))
  If ext = "xps" Or ext = "oxps" Then NombreSubida = f.Name: Exit Function
  If ext <> "txt" Then Exit Function
  If UCase(f.Name) = "CLIENTE.TXT" Or UCase(f.Name) = "ARTICULO.TXT" Then NombreSubida = UCase(f.Name): Exit Function
  Set re = New RegExp: re.IgnoreCase = True: re.Pattern = "^(cab|lin|cob)(\d\d)\.txt$"
  Set m = re.Execute(f.Name)
  If m.Count = 0 Then Exit Function
  d = f.DateLastModified
  NombreSubida = "R" & m(0).SubMatches(1) & "_" & Year(d) & Dos(Month(d)) & Dos(Day(d)) & "_" & Dos(Hour(d)) & Dos(Minute(d)) & "_" & LCase(m(0).SubMatches(0)) & ".txt"
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
Dim lista, i, nom, nsub: lista = Array()
For Each f In fso.GetFolder(ORIGEN).Files
  If NombreSubida(f) <> "" And DateDiff("s", f.DateLastModified, Now) > 20 Then
    ReDim Preserve lista(UBound(lista) + 1): lista(UBound(lista)) = f.Name
  End If
Next
For i = 0 To UBound(lista)
  nom = lista(i)
  Set f = fso.GetFile(ORIGEN & "\" & nom)
      nsub = NombreSubida(f)
      If Sube(f.Path, nsub) Then
        dest = env & "\" & nsub
        On Error Resume Next
        If fso.FileExists(dest) Then fso.DeleteFile dest, True
        fso.MoveFile f.Path, dest
        If Err.Number <> 0 Then
          Apunta "AVISO " & nom & " subido pero no se pudo mover a ENVIADOS: " & Err.Description
          Err.Clear
        Else
          Apunta "PDA SUBIDO  " & nom & " como " & nsub
        End If
        On Error GoTo 0
      End If
Next
