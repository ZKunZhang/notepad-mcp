**MCP Notepad (Windows)**

- Purpose: MCP server that controls Windows Notepad via PowerShell automation.
- Requirements: Windows, Node.js 18+, PowerShell (inbox), Notepad.

**Install**
- Run: `npm install`

**Run (stdio)**
- `npm start`

**Tools**
- `open_notepad` { filePath? }
- `list_notepad_windows` {}
- `paste_text` { text, pid? }
- `send_keys` { keys, pid? }
- `save_file` { filePath?, pid? }
- `close_notepad` { pid?, dontSave? }

Notes
- Uses clipboard + Ctrl+V for robust text insertion.
- Focuses a Notepad window (by pid or latest) before actions.
- Save As uses Ctrl+S then pastes the target path and presses Enter.

