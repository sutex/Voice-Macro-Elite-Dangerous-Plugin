# VoiceMacro Elite Dangerous Plugin (VM.ElitePlugin)

**VM.ElitePlugin** is a modular plugin for [VoiceMacro](https://www.voicemacro.net/) that interfaces with *Elite Dangerous* telemetry files, publishing structured variables for in-game voice automation and AI workflows.

---

## 🚀 Overview
The plugin monitors Elite Dangerous’ `Status.json` and rotating Journal files, parses events, and updates VoiceMacro variables (`vm_*` and persistent `*_p` mirrors`) in real time.  
It provides a lightweight JSON bridge (`Machina.signal.json`), Discord-style presence publishing, and full keybind introspection via `.binds` parsing.

No injection or memory hooks — everything is driven by file watchers and JSON parsing within Frontier’s EULA guidelines.

---

## ✳️ Main Capabilities
- **Status.json Watcher:** debounced monitoring; updates environment, HUD, Firegroup, GUI focus, etc.  
- **Journal Tailer:** safe rotation handling; populates `vm_*` variables from Journal events.  
- **Heartbeat + Persistence:** keeps `vm_plugin_alive`, mirrors bare variables to `*_p` for macro visibility.  
- **Presence Publishing:** builds one- or two-line presence text (Discord-style) written to `Machina.signal.json`.  
- **MachinaBridge:** atomic JSON snapshots for external tools and “AI” integrations.  
- **BindReader:** parses `.binds` XML, exports `ed_*` variables, generates `ED_Binds_Report.txt`, rotates backups.  
- **VmRemoteClient:** lightweight HTTP client to fire VoiceMacro macros via Remote Control (auto-retry localhost).  
- **GUI Router:** triggers macros on focus changes (CommsPanel / GalaxyMap).  
- **Command Interface:** `ReceiveParams` supports commands like `checkbinds`, `rp_on/off`, `aion/aioff`, `auditstart`, `getvar`, etc.  
- **Logger:** capped log file (`VMPlugin.log`) with rotation safety.  
- **Safe API Shim:** `VMApiShim` adapts to multiple VoiceMacro API versions gracefully.

---

## 🗂️ Files Produced
| File | Description |
|------|--------------|
| `VMPlugin.log` | Main plugin log (size-capped) |
| `StatusDump.txt / .json / .csv` | Audit snapshots |
| `VMVarTypes.txt` | List of macro-usable variables |
| `ED_Binds_Report.txt` | Human-readable keybind list |
| `Machina.signal.json` | Presence/AI bridge snapshot |
| `backups\*.binds` | Auto-rotated bind backups | 


---

## ⚙️ Compatibility
- **.NET Framework 4.6.1**
- **VoiceMacro 1.4+**
- **Elite Dangerous (Live / Odyssey)**
- No admin rights required (read-only game telemetry).

---

## 📄 License
Free for personal use.  
Developed for hobbyist under Frontier’s EULA and VoiceMacro plugin framework.

---

## 🧾 Credits
Created and maintained by **Sutex Osofar**  
Uses [Newtonsoft.Json 13.0.3](https://www.newtonsoft.com/json)

---

### 🔖 Quick Start
1. Drop `VM.ElitePlugin.dll` into your VoiceMacro `Plugins` folder.  
2. Start VoiceMacro → it will load automatically.  
3. Check `%LocalAppData%\VM.ElitePlugin\VMPlugin.log` for status.  
4. Use `checkbinds` or `edbinds` macros to verify binds.  
5. Open Elite Dangerous — variables appear in VoiceMacro (`vm_*`, `ed_*`).

---

All files live in:
