# vbastudio

**Debug Excel VBA the way you debug everything else.**

vbastudio is an external IDE for Excel VBA. It drives a real, open Excel
workbook over COM from VS Code — so you get real breakpoints, a real
Variables pane, and your VBA source as plain text files on disk, under git,
instead of trapped inside the VBE.

→ **[Full overview and features](website/index.html)**
→ **[Setup guide — build it and get running](website/install.html)**

## What it does

- **Real breakpoints and locals.** Set a breakpoint in the VS Code gutter,
  press F5, and execution actually stops there — inside your live workbook —
  with every local variable visible in the Variables pane.
- **Two-way sync.** Pull VBA source from Excel to a `src/` folder on disk,
  edit it like any other code, push changes back. A warning catches you
  before you ever debug a stale, un-pushed copy.
- **Fast ways to run anything.** A sidebar tree of every module and
  procedure, a status-bar quick-jump (type a name, hit enter), and a
  Command Palette flow — pick whichever's fastest.
- **No add-in inside Excel.** vbastudio drives the workbook you already
  have open; nothing to install inside Excel itself.

## Status

Built from source only right now — there's no published VS Code Marketplace
extension or downloadable installer yet. The [setup guide](website/install.html)
walks through building it yourself; it takes about 15 minutes and doesn't
assume prior VS Code or .NET experience.

## Project layout

| Path | What's there |
|---|---|
| `VbaStudio.Core/`, `VbaStudio.Interop/`, `VbaStudio.DapServer/` | The C# broker — COM automation, VBA parsing/instrumentation, the Debug Adapter Protocol server. |
| `VbaStudio.Tests/` | Unit tests for the broker. |
| `extension/` | The VS Code extension client. |
| `vba/` | VBA source injected into the debugged workbook at runtime (the probe agent). |
| `website/` | This project's site — [overview](website/index.html) and [setup guide](website/install.html). |
| `docs/` | Design specs and implementation plans written during development. |

## How it fits together

Excel keeps running your workbook. VS Code is where you read, edit, and
debug the code:

```
Excel (live workbook)  ⇄  src/*.bas, *.cls  (on disk, in git)  ⇄  VS Code (edit, breakpoints, debug)
```

Pull copies code out of Excel; push writes your edits back in; debugging
runs the real procedure in a temporary shadow copy of the workbook, so your
original file is never touched mid-run.

## Security

vbastudio automates a real, local Excel instance over COM — it runs with whatever
permissions the Excel process already has, the same trust boundary as any VBA macro
you'd run yourself. While a debug session is active, the broker listens on
`http://localhost:8731/probe/` (loopback only, never exposed to the network) to receive
the blocking calls that implement breakpoints; that endpoint is unauthenticated, so any
other process running as you on the same machine could reach it during a session. There
is no telemetry and no network activity beyond that loopback connection.

Found a security issue? Open a GitHub issue (or a private security advisory, if the repo
has one enabled) rather than a public pull request with exploit details.

## Trademarks

vbastudio is an independent project and is not affiliated with, endorsed by, or
sponsored by Microsoft Corporation. Excel, VBA, Visual Studio, and Visual Studio Code
are trademarks of Microsoft Corporation.

## License

[MIT](LICENSE)
