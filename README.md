# Ucs SQL Monitor

A lightweight Windows desktop monitor for SQL Server sessions, written in VB6. It shows the sessions of one or more servers in a single live list, highlights activity as it happens and, on SQL Server 2019 and later, keeps a history of the batches, RPCs, statements and errors each session runs.

## Features

- **Several servers in one list.** Every connected server feeds the same list, so sorting and filtering work across all of them. The Server column tells rows apart.
- **Two monitoring modes**, chosen per server in the Connect dialog:
  - **Extended Events (SQL 2019+):** live sessions and requests, plus a per-session history of completed batches, RPCs, errors (severity 11+) and cancellations, read from a shared Extended Events session. Optionally also statements inside stored procedures.
  - **sp_who2 (SQL 2000):** a polled `sp_who2` snapshot for old servers.
- **Server tree** with every saved connection profile. Connected servers are bold. Double-click a server to connect it with its saved settings; right-click for Connect, Disconnect and Properties.
- **Filtering** across Server, Program, DB, Host, Login, Status and Command, with `AND`, `OR`, `NOT`, brackets and `*` wildcards (see [Filter syntax](#filter-syntax)).
- **Sorting** by any column (click a header; Ctrl+click adds a secondary key), with sort arrows in the header.
- **Activity highlighting:** rows whose data changed since the last refresh are shown in green.
- **Session detail:** the lower pane shows the selected session's input buffer or event history.
- **Kill** the selected sessions, and a **Statistics** window that counts operations per host, login and database.

## Requirements

- Windows with the VB6 runtime (included in all current Windows versions).
- **Extended Events mode:** SQL Server 2019 or later and *ODBC Driver 17 or 18 for SQL Server* installed on the client. The newest installed driver is picked automatically.
- **sp_who2 mode:** any SQL Server from 2000 on, through the built-in SQLOLEDB provider.
- Permissions on the monitored server:
  - `VIEW SERVER STATE` for the session and request views;
  - `ALTER ANY EVENT SESSION` for Extended Events mode (to create, start and drop the monitoring session);
  - `ALTER ANY CONNECTION` (or `sysadmin`/`processadmin`) to kill sessions.

## Usage

Build it (see [Building](#building)), start `prj\UcsSqlMonitor.exe` and enter a server in the Connect dialog:

| Field | Meaning |
|---|---|
| SQL Server / SQL DB | Server name and optional initial database |
| User / Pass | SQL login; leave empty for Windows authentication |
| Type | Extended Events (SQL 2019+) or sp_who2 (SQL 2000) |
| Refresh rate | Polls per second for this server |
| Show system processes | Include system sessions |
| Encrypt connection | Encrypt the connection to the server |
| Trace statements in procedures | Also collect `sp_statement_completed` events (Extended Events mode) |

Each successful connection is saved as a profile, so the server appears in the tree and can be reconnected with a double-click. **Properties...** in the tree's context menu opens the dialog with that server's settings, and OK reconnects it with the changes.

| Shortcut | Action |
|---|---|
| Ctrl+F2 | Connect |
| Ctrl+F | Filter |
| F6 | Statistics |
| Ctrl+X | Kill selected sessions |

Settings and profiles are stored in `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Ucs SQL Monitor`.

### Filter syntax

| Filter | Shows rows where some column... |
|---|---|
| `dreem` | contains "dreem" |
| `dreem 2` | contains the phrase "dreem 2" (adjacent words are one phrase) |
| `dre*help` | starts with "dre" and ends with "help" (a term with `*` must match the whole value) |
| `sa AND master` | contains "sa", and some column contains "master" |
| `sa OR dreem` | contains either |
| `NOT sa` / `sa NOT master` | does not contain "sa" / contains "sa" but none contains "master" |
| `(sa OR dreem) AND NOT help` | brackets group; `AND` binds tighter than `OR` |
| `"and"` | contains the word "and" (quotes keep keywords as text) |

Matching is case-insensitive. `%` works like `*`.

## How Extended Events mode works

All monitors connected to a server share one event session named `UcsSqlMonitor` with an `event_file` target. Each monitor holds a shared application lock on it, and the last one to disconnect drops the session. The monitor's own queries are excluded from the captured events.

Every refresh reads only the events written since the previous one, together with the current sessions, requests and open transactions. Events reach the file within about one second (`MAX_DISPATCH_LATENCY`); the live request state is as fresh as the refresh rate. On connect, the history shows the last five minutes.

Statement events (`sp_statement_completed`) are added to the running session by the first monitor with **Trace statements in procedures** checked; monitors without the option skip them.

## Building

The project is built with Visual Basic 6.0 (SP6):

```bat
prj\build.bat
```

The script copies the sources to a temporary `compile` folder, adds line numbers with `VbCodeLines` (expected at `C:\work\BuildTools\VBCodeLines`), compiles with `VB6.EXE /m` into `prj\` and removes the temporary folder. The project can also be opened directly in the IDE from `prj\UcsSqlMonitor.vbp`; File → Make writes the executable to `prj\` as well. The executable is not tracked.

Two committed binaries only need rebuilding after their sources change:

- `res\UcsSqlMonitor.res` (manifest and tree icons): `powershell -File res\make-res.ps1`, which uses `rc.exe` from the Windows SDK.
- `typelib\UcsSqlMonitor.tlb` (helper interfaces for collections and in-place activation): `typelib\make.bat`, which uses `MKTYPLIB.EXE` from Visual C++ 6.0.

## Repository layout

| Path | Contents |
|---|---|
| `src` | Forms, classes, user controls and modules |
| `prj` | VB6 project, build script and the compiled executable (not tracked) |
| `res` | Manifest, application and tree icons, and the compiled resource file |
| `typelib` | Helper type library source and binary |

## Credits

Server tree icons are from the [Fugue Icons](https://p.yusukekamiyamane.com/) set by Yusuke Kamiyamane, licensed under [Creative Commons Attribution 3.0](https://creativecommons.org/licenses/by/3.0/).
