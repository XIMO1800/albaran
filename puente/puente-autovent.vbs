' PUENTE AUTOVENT - sube a Google los resumenes de la PDA (.xps) que deja el programa en SERVIDORW10\AUTOVENT,
' y los datos de la PDA (cab01/lin01/cob01, cab02/lin02/cob02, CLIENTE.TXT, ARTICULO.TXT) para la revision de albaranes.
' Los cab/lin/cob se suben con la ruta y la fecha y hora del fichero en el nombre (R02_20261120_1741_cab.txt) para que no se machaquen.
' Lo lanza una tarea programada cada 2 minutos. Cada fichero subido se mueve a AUTOVENT\ENVIADOS.
' Ademas mira la carpeta compartida de FACTURACION (enlaces), donde deja los ficheros la descarga de la PDA, y sube una copia
' de los que sean nuevos (alli no se toca ni se mueve nada). Asi no hace falta copiarlos a mano.
' La clave la coge de puente-albtiend.vbs (tiene que estar en la misma carpeta). Apunta lo que hace en envios.log.
Option Explicit
Const URL = "https://script.google.com/macros/s/AKfycbysLH9zV2Tss0fjnkv9USLkm6HZx7y_sBxEPf9_XG_EzxTCydsWcHPoEPY6uGp51qk8LA/exec"
Const ORIGEN = "\\Servidor-i7\servidor\SERVIDORW10\AUTOVENT"
Const ENLACES = "\\FACTURACION\enlaces"

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
  d = FechaRuta(f, m(0).SubMatches(1))
  NombreSubida = "R" & m(0).SubMatches(1) & "_" & Year(d) & Dos(Month(d)) & Dos(Day(d)) & "_" & Dos(Hour(d)) & Dos(Minute(d)) & "_" & LCase(m(0).SubMatches(0)) & ".txt"
End Function

' Fecha con la que se nombran los ficheros de una descarga (R01_AAAAMMDD_HHMM_...): la del cab de esa descarga.
' Cuando la PDA se descarga, cab, lin y cob se escriben a la vez; esa fecha se guarda en pda-idRR.txt.
' Si despues alguien corrige a mano uno de ellos (Bernardo arregla unos kilos), conserva la fecha de su descarga:
' se sube con el mismo nombre y sustituye al que estaba mal.
Function FechaRuta(f, rr)
  Dim pc, pl, t, tid, st
  FechaRuta = f.DateLastModified
  pc = f.ParentFolder.Path & "\cab" & rr & ".txt"
  If Not fso.FileExists(pc) Then Exit Function
  t = fso.GetFile(pc).DateLastModified
  pl = f.ParentFolder.Path & "\lin" & rr & ".txt"
  st = carpeta & "\pda-id" & rr & ".txt"
  tid = Empty
  On Error Resume Next
  If fso.FileExists(st) Then tid = CDate(CDbl(Trim(fso.OpenTextFile(st, 1).ReadAll)))
  On Error GoTo 0
  If fso.FileExists(pl) Then
    If Abs(DateDiff("n", t, fso.GetFile(pl).DateLastModified)) <= 10 Then
      tid = t   ' descarga nueva (cab y lin escritos juntos)
      On Error Resume Next
      Dim w: Set w = fso.OpenTextFile(st, 2, True): w.Write CStr(CDbl(t)): w.Close
      On Error GoTo 0
    End If
  End If
  If IsEmpty(tid) Then tid = t
  FechaRuta = tid
End Function

' lin/cob que son de una descarga anterior (no se han vuelto a escribir en la ultima, p. ej. cob sin cobros) no se suben
Function EsDeOtraDescarga(f)
  Dim re, m
  EsDeOtraDescarga = False
  Set re = New RegExp: re.IgnoreCase = True: re.Pattern = "^(lin|cob)(\d\d)\.txt$"
  Set m = re.Execute(f.Name)
  If m.Count = 0 Then Exit Function
  If DateDiff("n", FechaRuta(f, m(0).SubMatches(1)), f.DateLastModified) < -10 Then EsDeOtraDescarga = True
End Function

' El mismo fichero ya esta en ENVIADOS (misma fecha y tamano): se copio a mano y ya se subio
Function MismoFichero(ruta, f)
  Dim g
  MismoFichero = False
  If Not fso.FileExists(ruta) Then Exit Function
  Set g = fso.GetFile(ruta)
  If g.Size = f.Size And Abs(DateDiff("s", g.DateLastModified, f.DateLastModified)) <= 2 Then MismoFichero = True
End Function

' Registro de lo ya subido desde enlaces: nombre|fecha|tamano
Function Firma(f)
  Firma = UCase(f.Name) & "|" & CStr(CDbl(f.DateLastModified)) & "|" & f.Size
End Function
Function YaSubido(firm)
  Dim t
  YaSubido = False
  On Error Resume Next
  If Not fso.FileExists(carpeta & "\pda-enlaces.txt") Then Exit Function
  t = fso.OpenTextFile(carpeta & "\pda-enlaces.txt", 1).ReadAll
  If InStr(1, vbCrLf & t, vbCrLf & firm & vbCrLf, 1) > 0 Then YaSubido = True
End Function
Sub ApuntaSubido(firm)
  On Error Resume Next
  Dim f: Set f = fso.OpenTextFile(carpeta & "\pda-enlaces.txt", 8, True)
  f.WriteLine firm
  f.Close
End Sub

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
Dim lista, i, nom, nsub: lista = Array()
env = ORIGEN & "\ENVIADOS"
' si el servidor no responde se apunta y se sigue con la carpeta de FACTURACION
If Not fso.FolderExists(ORIGEN) Then
  Apunta "ERROR no encuentro la carpeta " & ORIGEN
Else
On Error Resume Next
If Not fso.FolderExists(env) Then fso.CreateFolder env
On Error GoTo 0
For Each f In fso.GetFolder(ORIGEN).Files
  If NombreSubida(f) <> "" And DateDiff("s", f.DateLastModified, Now) > 20 Then
    ' el nombre de subida se calcula ahora, antes de mover nada (el de lin/cob depende del cab que esta al lado)
    ReDim Preserve lista(UBound(lista) + 1): lista(UBound(lista)) = f.Name & "|" & NombreSubida(f)
  End If
Next
For i = 0 To UBound(lista)
  nom = Split(lista(i), "|")(0): nsub = Split(lista(i), "|")(1)
  Set f = fso.GetFile(ORIGEN & "\" & nom)
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
End If

' ---------------- carpeta enlaces de FACTURACION (solo se lee) ----------------
If fso.FolderExists(ENLACES) Then
  Dim firm, tmp
  For Each f In fso.GetFolder(ENLACES).Files
    nsub = NombreSubida(f)
    If nsub <> "" And LCase(fso.GetExtensionName(f.Name)) = "txt" Then
      firm = Firma(f)
      If DateDiff("s", f.DateLastModified, Now) > 60 And Not YaSubido(firm) And Not EsDeOtraDescarga(f) Then
        If UCase(Left(nsub, 1)) = "R" And MismoFichero(env & "\" & nsub, f) Then
          ApuntaSubido firm   ' ya se subio copiandolo a mano
        Else
          tmp = carpeta & "\_" & nsub
          On Error Resume Next
          fso.CopyFile f.Path, tmp, True
          If Err.Number <> 0 Then
            Apunta "ERROR no se pudo copiar " & f.Path & ": " & Err.Description: Err.Clear
          Else
            On Error GoTo 0
            If Sube(tmp, nsub) Then
              ApuntaSubido firm
              Apunta "PDA SUBIDO  " & f.Path & " como " & nsub
              On Error Resume Next
              If UCase(Left(nsub, 1)) = "R" Then fso.CopyFile tmp, env & "\" & nsub, True
            End If
            On Error Resume Next
            fso.DeleteFile tmp, True
          End If
          On Error GoTo 0
        End If
      End If
    End If
  Next
End If
