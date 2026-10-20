# DEFENSE.md

Detection guidance for **vba-agent** — the VBA beacon module. For the
protocol it speaks, see the sibling agents' DEFENSE files; the wire is
shared.

## Network indicators

- **Beacon cadence.** One `POST /` per cycle to the configured origin.
  Idle answers are held open ~25–35 s (the hold IS the idle sleep), so a
  quiet agent emits ~2 requests/minute with an empty body and an empty
  200 back. Cadence, not volume, is the tell.
- **Identity header cluster.** Every POST carries the exact set
  `X-Api-Version: 1`, `X-Client-Id: 6`,
  `X-Client-Features: 0000000000000000`, plus `X-Device-Id`,
  `X-Session-Id`, `X-Device-Name`, `X-User-Id`, arch and OS headers.
  `X-Client-Id: 6` with an all-zero feature mask is the breed signature;
  the header NAMES alone are a strong anomaly signal for any egress.
- **Log ships.** Milestone lines (`vba agent beaconing to …`,
  `idle - empty answer`, `http <n>`, `exit`) POSTed with
  `X-Log-Only: 1`, at most 3 in flight, each answered immediately.
  These are the best content-level signatures available.
- **TLS/client fingerprint.** Requests come from `WINWORD.EXE` /
  `EXCEL.EXE` but present an MSXML `ServerXMLHTTP` (WinHTTP) client
  fingerprint and user agent — **not** the Office/ClickOnce update
  pattern the same process normally uses. Office process → non-Office
  HTTP stack is a high-value correlation.

## Host indicators

- **VBA project contents.** A module named `C2Agent` with
  `Public Function RunAgent()`; the one native declare is
  `Sleep` from `kernel32`. Look for the marker strings:
  `RunAgent`, `H_URL`, `MSXML2.ServerXMLHTTP`, `X-Log-Only`,
  `winmgmts:` / `StdRegProv`, `Win32_ComputerSystemProduct`,
  `command opcode`, `idle - empty answer`.
  (Loaded VBA source is AMSI-scanned on Windows 10+ at macro load —
  content inspection applies there; inside `vbaProject.bin` at rest the
  source is compressed MS-OVBA.)
- **Process behavior.** Office process doing machine inventory: WMI to
  `root\default:StdRegProv` (registry read of `MachineGuid`, explicit
  `Wow6432Node` path) and `root\cimv2` (`Win32_OperatingSystem`,
  `Win32_Processor`, `Win32_ComputerSystemProduct`), plus
  `Scriptlet.TypeLib` GUID generation — an unusual COM cluster for
  WINWORD/EXCEL.
- **Message-loop pumping.** The agent `DoEvents`-pumps through long HTTP
  holds, so the Office window stays responsive while a request is in
  flight for 25–35 s — no hang dialog. Responsive-but-idle Office with
  an open winhttp socket is the behavioral shape.
- **Environment coupling.** The module reads only `H_URL` plus the
  standard env arch/user variables; an `H_URL` process-environment
  variable pointing at a non-corporate origin on an Office process is a
  direct IOC.

## MITRE ATT&CK

| Tactic | Technique | Where |
|---|---|---|
| Initial Access | T1566.001 Spearphishing Attachment | the delivery vehicle this module is built into |
| Execution | T1204.002 Malicious File (user execution of macros) | host master / document open |
| Command and Control | T1071.001 Web Protocols | the beacon POST leg |
| Command and Control | T1132 Data Encoding (length-prefixed binary frames) | frame codec |
| Discovery | T1082 System Information Discovery | identity headers (arch, OS, hostname, user) |
| Discovery | T1012 Query Registry | `MachineGuid` derivation |

## Triage checklist

1. Suspect Office process with an outbound socket → check for the
   identity header cluster or `X-Log-Only` on any captured request.
2. Capture the `X-Device-Id` / `X-Session-Id` / `X-Device-Name` /
   `X-User-Id` values — they key the operator's roster row and correlate
   every other sighting of the machine.
3. Dump `vbaProject.bin` from any suspect document (oletools /
   `olevba`) and grep for the module strings above.
4. Hunt the WMI/registry COM cluster in Office process ETW/Sysmon logs
   (Sysmon events 1/3/12–13 correlation: process WINWORD + network
   connect + `MachineGuid` registry read).
5. Block egress for the beacon origin; the agent does NOT retry — one
   failed cycle ends the run (`"fail"`), so egress denial is clean
   containment for this breed.
