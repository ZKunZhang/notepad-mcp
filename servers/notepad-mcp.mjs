#!/usr/bin/env node
import { spawn } from "node:child_process";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

function isWindows() {
  return process.platform === "win32";
}

function delay(ms) {
  return new Promise((res) => setTimeout(res, ms));
}

function runPowerShell(script, { timeoutMs = 15000 } = {}) {
  return new Promise((resolve, reject) => {
    const exe = "powershell.exe"; // Prefer inbox PowerShell for broad compatibility
    const ps = spawn(exe, [
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy",
      "Bypass",
      "-Command",
      "-",
    ]);
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => {
      ps.kill();
      reject(new Error("PowerShell timed out"));
    }, timeoutMs);
    ps.stdout.on("data", (d) => (stdout += d.toString()));
    ps.stderr.on("data", (d) => (stderr += d.toString()));
    ps.on("error", reject);
    ps.on("exit", (code) => {
      clearTimeout(timer);
      if (code === 0) resolve({ stdout: stdout.trim() });
      else reject(new Error(stderr.trim() || `PowerShell exited ${code}`));
    });
    ps.stdin.end(script);
  });
}

async function openNotepad(filePath) {
  const args = [];
  if (filePath && String(filePath).trim()) args.push(String(filePath));
  const child = spawn("notepad.exe", args, {
    detached: true,
    stdio: "ignore",
  });
  const pid = child.pid;
  child.unref();
  // Let window initialize
  await waitForMainWindow(pid, 50, 100);
  return { pid };
}

async function waitForMainWindow(pid, retries = 50, sleepMs = 100) {
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
$pid = ${Number(pid)}
for ($i=0; $i -lt ${Number(retries)}; $i++) {
  try {
    $p = Get-Process -Id $pid
    if ($p -and $p.MainWindowHandle -ne 0) { Write-Output $p.MainWindowHandle; break }
  } catch {}
  Start-Sleep -Milliseconds ${Number(sleepMs)}
}`;
  await runPowerShell(script).catch(() => {});
}

async function focusNotepad({ pid } = {}) {
  const target = pid ? `Get-Process -Id ${Number(pid)}` : `Get-Process notepad | Sort-Object StartTime -Descending | Select-Object -First 1`;
  const script = `
$ErrorActionPreference = 'Stop'
$p = ${target}
for ($i=0; $i -lt 50; $i++) {
  if ($p -and $p.MainWindowHandle -ne 0) { break }
  Start-Sleep -Milliseconds 100
}
if (-not $p -or $p.MainWindowHandle -eq 0) { throw 'Notepad window not ready' }
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
[Win]::SetForegroundWindow([IntPtr]$p.MainWindowHandle) | Out-Null
";
  await runPowerShell(script);
  await delay(150);
}

function psHereStringEscape(text = "") {
  // Use a double-quoted here-string; escape closing token if it appears.
  return String(text).replace(/"@/g, '"@" + "" + "@');
}

async function pasteTextIntoNotepad({ text, pid } = {}) {
  await focusNotepad({ pid });
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$val = @"
${psHereStringEscape(text)}
"@
Set-Clipboard -Value $val
[System.Windows.Forms.SendKeys]::SendWait('^v')
`;
  await runPowerShell(script);
}

async function sendKeysToNotepad({ keys, pid } = {}) {
  await focusNotepad({ pid });
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('${String(keys).replace(/'/g, "''")}')
`;
  await runPowerShell(script);
}

async function saveNotepad({ filePath, pid } = {}) {
  await focusNotepad({ pid });
  if (!filePath) {
    // Simple save
    await sendKeysToNotepad({ keys: '^s' });
    return;
  }
  const script = `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('^s')
Start-Sleep -Milliseconds 200
$path = @"
${psHereStringEscape(filePath)}
"@
Set-Clipboard -Value $path
[System.Windows.Forms.SendKeys]::SendWait('^a')
[System.Windows.Forms.SendKeys]::SendWait('^v')
Start-Sleep -Milliseconds 100
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
`;
  await runPowerShell(script);
}

async function closeNotepad({ pid, dontSave = true } = {}) {
  await focusNotepad({ pid });
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.SendKeys]::SendWait('%{F4}')
Start-Sleep -Milliseconds 200
# If prompted to save: Alt+N (don't save) or Alt+Y (save) – fallback to letter
try {
  if (${dontSave ? 1 : 0}) { [System.Windows.Forms.SendKeys]::SendWait('n') } else { [System.Windows.Forms.SendKeys]::SendWait('y') }
} catch {}
`;
  await runPowerShell(script).catch(() => {});
}

async function listNotepadWindows() {
  const script = `
$ErrorActionPreference = 'SilentlyContinue'
$procs = Get-Process notepad | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object Id, MainWindowTitle
$procs | ConvertTo-Json -Compress
`;
  const { stdout } = await runPowerShell(script);
  try {
    const parsed = stdout ? JSON.parse(stdout) : [];
    // Normalize to array
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}

async function main() {
  if (!isWindows()) {
    console.error("This MCP server only runs on Windows.");
    process.exit(1);
  }

  const server = new Server(
    { name: "mcp-notepad", version: "0.1.0" },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler("tools/list", async () => ({
    tools: [
      {
        name: "open_notepad",
        description: "Open Windows Notepad optionally with a file path.",
        input_schema: {
          type: "object",
          properties: { filePath: { type: "string", description: "Optional file path to open" } },
        },
      },
      {
        name: "list_notepad_windows",
        description: "List open Notepad windows with pid and title.",
        input_schema: { type: "object", properties: {} },
      },
      {
        name: "paste_text",
        description: "Paste provided text into the focused Notepad (or by pid). Uses clipboard + Ctrl+V for reliability.",
        input_schema: {
          type: "object",
          required: ["text"],
          properties: {
            text: { type: "string", description: "Text to insert (clipboard-based paste)" },
            pid: { type: "number", description: "Optional Notepad process id to focus before pasting" },
          },
        },
      },
      {
        name: "send_keys",
        description: "Send keystrokes to Notepad (e.g., ^s for Ctrl+S, %{F4} for Alt+F4).",
        input_schema: {
          type: "object",
          required: ["keys"],
          properties: {
            keys: { type: "string", description: "SendKeys format string" },
            pid: { type: "number", description: "Optional Notepad process id" },
          },
        },
      },
      {
        name: "save_file",
        description: "Save the current Notepad document. If filePath is provided, perform Save As to that path.",
        input_schema: {
          type: "object",
          properties: {
            filePath: { type: "string", description: "Optional absolute path to Save As" },
            pid: { type: "number", description: "Optional Notepad process id" },
          },
        },
      },
      {
        name: "close_notepad",
        description: "Close Notepad, optionally selecting Don't Save.",
        input_schema: {
          type: "object",
          properties: {
            pid: { type: "number", description: "Optional Notepad process id" },
            dontSave: { type: "boolean", description: "If true, choose Don't Save on prompt", default: true },
          },
        },
      },
    ],
  }));

  server.setRequestHandler("tools/call", async (req) => {
    const name = req.params.name;
    const args = (req.params.arguments ?? {});
    try {
      switch (name) {
        case "open_notepad": {
          const { filePath } = args;
          const { pid } = await openNotepad(filePath);
          return { content: [{ type: "text", text: `Opened Notepad (pid=${pid})` }] };
        }
        case "list_notepad_windows": {
          const list = await listNotepadWindows();
          return { content: [{ type: "text", text: JSON.stringify(list) }] };
        }
        case "paste_text": {
          const { text, pid } = args;
          if (typeof text !== "string" || !text.length) throw new Error("text is required");
          await pasteTextIntoNotepad({ text, pid });
          return { content: [{ type: "text", text: "Pasted text into Notepad" }] };
        }
        case "send_keys": {
          const { keys, pid } = args;
          if (typeof keys !== "string" || !keys.length) throw new Error("keys is required");
          await sendKeysToNotepad({ keys, pid });
          return { content: [{ type: "text", text: `Sent keys: ${keys}` }] };
        }
        case "save_file": {
          const { filePath, pid } = args;
          await saveNotepad({ filePath, pid });
          return { content: [{ type: "text", text: filePath ? `Saved As: ${filePath}` : "Saved" }] };
        }
        case "close_notepad": {
          const { pid, dontSave } = args;
          await closeNotepad({ pid, dontSave: dontSave !== false });
          return { content: [{ type: "text", text: "Closed Notepad" }] };
        }
        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Error: ${err?.message || String(err)}` }],
      };
    }
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
