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
Created by Sutex Osofar using ChatGPT and cross-checking with CoPilot.

The idea from the beginning, since I'm not a coder and had to learn, was to prove the viability of AI to make a Voice Macro plugin for Elite Dangerous; only time and players testing the plugin will prove if that goal has been achieved. As such, I won't be maintaining this plugin since I have reached the end of the line of using AI (herding feral cats) to get this far. It will now be up to individual users (those who can actually code) to improve/polish the plugin from this point forward if indeed this plugin proves to be useful :)
Uses [Newtonsoft.Json 13.0.3](https://www.newtonsoft.com/json)

---

### 🔖 Quick Start
1. Drop `VM.ElitePlugin.dll` into your VoiceMacro `Plugins` folder.  
2. Start VoiceMacro → it will load automatically.  
3. Check `%LocalAppData%\VM.ElitePlugin\VMPlugin.log` for status.  
4. Use `checkbinds` or `edbinds` macros to verify binds.  
5. Open Elite Dangerous — variables appear in VoiceMacro (`vm_*`, `ed_*`).

---

VoiceMacro variables (vm_*)
State & Navigation

vm_docked (bool) — 1 when docked; else 0

vm_docked_p (bool) — persistent mirror

vm_in_supercruise (bool) — 1 in supercruise; else 0

vm_in_supercruise_p (bool) — persistent mirror

vm_on_foot (bool) — 1 on foot; else 0

vm_on_foot_p (bool) — persistent mirror

vm_environment (text) — Derived: Docked / Supercruise / OnFoot / Space

vm_environment_p (text) — persistent mirror

vm_hud_mode (text) — “Combat” or “Analysis” (from Status/Journal)

vm_hud_mode_p (text) — persistent mirror

vm_firegroup (int) — active firegroup index

vm_firegroup_p (int) — persistent mirror

vm_startjump (bool pulse) — hyperspace StartJump (0 then 1 only on JumpType=Hyperspace)

vm_startjump_p (bool) — persistent mirror

Location

vm_star_system (text) — current star system

vm_star_system_p (text) — persistent mirror

vm_body (text) — current body name

vm_body_p (text) — persistent mirror

vm_station_name (text) — station name if docked

vm_station_name_p (text) — persistent mirror

GUI Focus Router

vm_gui_focus (int) — numeric GUI focus code from Status.json

vm_gui_focus_p (int) — persistent mirror

vm_gui_focus_name (text) — friendly focus name (e.g., CommsPanel, GalaxyMap)

vm_gui_focus_name_p (text) — persistent mirror

Targeting

vm_target_locked (bool) — 1 when target lock active; else 0

vm_target_name_s (text) — target name (string helper used by presence)

Music

vm_music_track (text) — current music cue from Journal (short-lived)

vm_music_track_p (text) — persistent mirror

Wing

vm_in_wing (bool) — 1 when in wing; else 0

vm_in_wing_p (bool) — persistent mirror

vm_wing_members (int) — best-effort wing member count

vm_wing_members_p (int) — persistent mirror

vm_wing_leader (text) — wing leader name (when known)

vm_wing_leader_p (text) — persistent mirror

Multicrew

vm_in_multicrew (bool) — 1 when in multicrew; else 0

vm_in_multicrew_p (bool) — persistent mirror

vm_multicrew_role (text) — e.g., Gunner

vm_multicrew_role_p (text) — persistent mirror

vm_multicrew_captain (text) — captain name

vm_multicrew_captain_p (text) — persistent mirror

Paths & Timestamps

vm_status_path (text) — full path to Status.json

vm_status_path_p (text) — persistent mirror

vm_journal_dir (text) — “Saved Games” Journal folder

vm_journal_dir_p (text) — persistent mirror

vm_active_journal (text) — current journal filename being tailed

vm_active_journal_p (text) — persistent mirror

vm_status_last_ts (text) — last Status.json “timestamp” processed (display/local)

vm_status_last_ts_p (text) — persistent mirror

vm_journal_last_ts (text) — last Journal event timestamp processed (display/local)

vm_journal_last_ts_p (text) — persistent mirror

Plugin Health

vm_plugin_loaded (bool) — 1 after plugin init completes

vm_plugin_loaded_p (bool) — persistent mirror

vm_plugin_loaded_s (text) — string helper used in code (init seed)

vm_plugin_alive (iso8601) — heartbeat time (UTC ISO)

vm_plugin_alive_p (iso8601) — persistent mirror

vm_plugin_alive_s (iso8601) — string helper mirror for display

Presence / Machina

vm_presence_oneshot_s (text) — one-shot presence line (consumed → reset to “-”)

vm_rp_state_s (text) — presence state “on”/“off”

vm_ai_last_action_s (text) — last Machina/AI action string

Remote Macro Client (VoiceMacro Remote)

vm_remote_baseurl_s (text) — base URL for ExecuteMacro (e.g., http://127.0.0.1:8080)

vm_remote_debounce_ms (int) — debounce window for firing macros (ms)

Audit & Debug (helpers that surface into VM)

vm_audit_state_s (text) — audit timer state running/stopped

vm_audit_last_path_s (text) — path of last StatusDump written

vm_dbg_last_cmd_s (text) — last command string received by plugin

vm_log_checkbinds_s (text) — human-readable checkbinds / report path

vm_self_test_result_s (text) — selftest summary (HUD/FG/System/OnFoot)

vm_last_types_path_s (text) — last VMVarTypes.txt written path

vm_getvar_value_s (text) — result from getvar <name>

Mirrors: Every “bare” vm_* above that represents state is mirrored to a persistent *_p counterpart by the plugin’s central VMSet path, so your macros can read stable values even across session reloads.

Bind variables (ed_*) — what’s guaranteed + how the rest are created

These are written by BindReader and its report path:

Fixed/core ed_*:

ed_bindspath (text) — full path of the active .binds file

ed_status (text) — ok / parse_error / not_found

ed_error (text) — parse/backup error message (short)

ed_report (text) — full human-readable keybinds report text

ed_backup_last (text) — last backup file path written

ed_backup_error (text) — last backup error message

ed_utc (text) — timestamp used in the report

(helper scratch seen in code) ed_p, ed_s (internal helper/text scratch)

Dynamic per-action exports:

ed_<ActionName> (text) — one per action in the binds XML, value is the chosen keyboard chord (secondary preferred), or "" (quoted empty) if unbound.
Examples (names depend on your .binds): ed_ThrottleUp, ed_UIToggle, ed_TargetNextSystem, etc.

BindReader also writes ED_Binds_Report.txt and sets vm_log_checkbinds_s with guidance/report paths.

How to generate VMVarTypes.txt (from your plugin)

The code exposes a command that writes a curated variable guide to disk:

Command names: vmvartypes or types

What happens: plugin runs WriteVmVarTypes(), writes %LocalAppData%\VM.ElitePlugin\VMVarTypes.txt, and sets:

vm_last_types_path_s → full path to the file

vm_log_checkbinds_s → progress/error text if something goes wrong

Two easy ways to trigger it

From a VoiceMacro macro:

Action: Other… → Call External Plug-In → select VM.ElitePlugin

Parameters: vmvartypes

Run the macro; then check vm_last_types_path_s or open the file path.

Via Remote Execute (if you use your HTTP client):

Call your macro that passes vmvartypes into the plugin’s ReceiveParams (your VmRemoteClient already handles retries/localhost fallback).

The generated VMVarTypes.txt contains sections:

Flight/State (vm_docked_p, vm_in_supercruise_p, …)

Location (vm_star_system_p, vm_station_name_p, …)

Focus/Target/Wing/Multicrew, etc.

Plugin/Health, Remote, Presence/AI tools
…and a one-line explanation for each macro-usable *_p var (exact strings are baked in your Plugin.vb).

Quick reference — Commands you’ll actually use (ReceiveParams)

vmvartypes / types — write VMVarTypes.txt (set vm_last_types_path_s)

checkbinds, checkbindings, checkbind — validate binds; update vm_log_checkbinds_s & report

edbinds — parse binds & export all ed_*

rp_on, rp_off, lfw_on, lfw_off — presence toggles/oneshot

aion, aioff, aitoggle — Machina/AI bridge toggles

auditstart, auditstop, auditstatus, auditcsv — audit controls/snapshots

setremote <url>, setremotedebounce <ms>, getremote, helpreomte, remotecheck [macro] — remote client controls

getvar <name> — publish value to vm_getvar_value_s

All files live in:
