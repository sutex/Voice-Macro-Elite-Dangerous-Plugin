Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports vmAPI ' vmInterface only

<ComVisible(True)>
Public Class Plugin
    Implements vmAPI.vmInterface

    ' -----------------------------
    ' Logging & base dir
    ' -----------------------------
    Private _log As Logger
    Private _baseDir As String = ""
    Private Sub WriteLog(msg As String)
        If _log IsNot Nothing Then _log.Write(msg)
    End Sub

    ' -----------------------------
    ' VM audit (bare vars only)
    ' -----------------------------
    Private Class AuditEntry
        Public Name As String
        Public Value As String = ""
        Public TypeName As String = "string"
        Public UpdatedUtc As DateTime = DateTime.MinValue
        Public SeenWrite As Boolean = False
        Public SeenRead As Boolean = False
        Public Sub UpdateType()
            Dim v = If(Value, "")
            If v = "" Then
                TypeName = "string" : Exit Sub
            End If
            Dim s = v.Trim().ToLowerInvariant()
            Dim i32 As Integer, i64 As Long, d As Double
            If Integer.TryParse(v, i32) Then TypeName = "int" : Exit Sub
            If Long.TryParse(v, i64) Then TypeName = "int" : Exit Sub
            If Double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, d) Then TypeName = "float" : Exit Sub
            If s = "0" OrElse s = "1" OrElse s = "true" OrElse s = "false" Then TypeName = "bool" : Exit Sub
            If v.Length >= 20 AndAlso (s.Contains("t") OrElse s.Contains(":")) AndAlso (s.EndsWith("z") OrElse s.Contains("+")) Then
                TypeName = "iso8601" : Exit Sub
            End If
            TypeName = "string"
        End Sub
    End Class

    Private ReadOnly _audit As New Dictionary(Of String, AuditEntry)(StringComparer.OrdinalIgnoreCase)
    Private _auditTimer As System.Threading.Timer
    Private _auditPeriodMs As Integer = 30000 ' 30s
    Private _auditRunning As Boolean = False

    Private Function GetAudit(name As String) As AuditEntry
        If String.IsNullOrEmpty(name) Then Return Nothing
        Dim e As AuditEntry = Nothing
        If Not _audit.TryGetValue(name, e) Then
            e = New AuditEntry() With {.Name = name}
            _audit(name) = e
        End If
        Return e
    End Function

    Private Sub TrackWrite(name As String, value As String)
        Dim e = GetAudit(name)
        If e Is Nothing Then Exit Sub
        e.SeenWrite = True
        e.Value = If(value, "")
        e.UpdateType()
        e.UpdatedUtc = DateTime.UtcNow
    End Sub

    Private Sub TrackRead(name As String, value As String)
        Dim e = GetAudit(name)
        If e Is Nothing Then Exit Sub
        e.SeenRead = True
        If value IsNot Nothing AndAlso value <> "" Then
            e.Value = value
            e.UpdateType()
            e.UpdatedUtc = DateTime.UtcNow
        End If
    End Sub

    Private Function NormalizeVmRead(v As String) As String
        If v Is Nothing Then Return ""
        Dim t = v.Trim()
        If t.Equals("ERR!", StringComparison.OrdinalIgnoreCase) _
           OrElse t.Equals("NULL", StringComparison.OrdinalIgnoreCase) _
           OrElse t.Equals("N/A", StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If
        Return v
    End Function

    ' -----------------------------
    ' VM shim (+ mirror-on-write for _p)
    ' -----------------------------
    Private Shared Function HasVmScopeSuffix(name As String) As Boolean
        If String.IsNullOrEmpty(name) Then Return False
        Dim n = name.ToLowerInvariant()
        Return n.EndsWith("_p") OrElse n.EndsWith("_g") OrElse n.EndsWith("_s")
    End Function

    Private Shared Function IsVmBareCandidate(name As String) As Boolean
        If String.IsNullOrEmpty(name) Then Return False
        If Not name.StartsWith("vm_", StringComparison.OrdinalIgnoreCase) Then Return False
        If HasVmScopeSuffix(name) Then Return False
        Return True
    End Function

    Private Sub VMSet(name As String, value As String)
        Try
            Dim n As String = If(name, "").Trim()
            Dim v As String = If(value, "")

            If n.Length = 0 Then Exit Sub

            TrackWrite(n, v)
            VMApiShim.SetVariable(n, v, AddressOf WriteLog)

            If IsVmBareCandidate(n) Then
                Dim mirror As String = n & "_p"
                TrackWrite(mirror, v)
                VMApiShim.SetVariable(mirror, v, AddressOf WriteLog)
            End If

        Catch ex As Exception
            WriteLog("[VMSet] " & ex.Message)
        End Try
    End Sub

    Private Function VMGet(name As String) As String
        Try
            Dim n As String = If(name, "").Trim()
            Dim v = NormalizeVmRead(VMApiShim.GetVariable(n, AddressOf WriteLog))
            TrackRead(n, v)
            Return v
        Catch ex As Exception
            WriteLog("[VMGet] " & ex.Message)
            Return ""
        End Try
    End Function

    ' -----------------------------
    ' Paths, watchers, state
    ' -----------------------------
    Private _journalDir As String = ""
    Private _statusPath As String = ""
    Private _statusWatcher As FileSystemWatcher
    Private _journalWatcher As FileSystemWatcher

    Private _activeJournal As String = Nothing
    Private _lastJournalWriteUtc As DateTime = DateTime.MinValue
    Private _tailPosition As Long = 0
    Private ReadOnly _journalLock As New Object

    Private _statusDebounceDue As Integer = 0
    Private _journalDebounceDue As Integer = 0

    Private _handlers As EventHandlers
    Private _presence As DiscordPresence

    ' -----------------------------
    ' Machina + Heartbeat
    ' -----------------------------
    Private _machina As MachinaBridge
    Private _machinaEnabled As Boolean = False
    Private _hbTimer As System.Threading.Timer

    ' -----------------------------
    ' VM Remote (ExecuteMacro) — Restricted to Comms/Galaxy ONLY
    ' -----------------------------
    Private _vmProfileName As String = ""
    Private _remote As VmRemoteClient = Nothing
    Private _remoteBaseUrl As String = "http://localhost:8080/"
    Private _remoteDebounceMs As Integer = 350
    Private _remoteConfigPath As String = ""
    Private ReadOnly _remoteLock As New Object()

    ' GUI Focus router (Comms/Galaxy only)
    Private _lastGuiFocusName As String = ""

    ' -----------------------------
    ' Heartbeat
    ' -----------------------------
    Private Sub StartHeartbeat()
        Try
            If _hbTimer IsNot Nothing Then Return
            _hbTimer = New System.Threading.Timer(
                Sub(state As Object)
                    Try
                        Dim ts As String = DateTime.UtcNow.ToString("o")
                        VMSet("vm_plugin_alive", ts)
                        VMSet("vm_plugin_alive_s", ts)
                        CheckPresenceOneShot()
                        RefreshPresenceSnapshot()
                    Catch
                    End Try
                End Sub,
                Nothing, 2000, 5000)
            WriteLog("Heartbeat started")
        Catch ex As Exception
            WriteLog("Heartbeat error: " & ex.Message)
        End Try
    End Sub

    Private Sub StopHeartbeat()
        Try
            If _hbTimer IsNot Nothing Then
                _hbTimer.Dispose()
                _hbTimer = Nothing
                WriteLog("Heartbeat stopped")
            End If
        Catch ex As Exception
            WriteLog("StopHeartbeat error: " & ex.Message)
        End Try
    End Sub

    ' -----------------------------
    ' vmInterface — properties
    ' -----------------------------
    Public ReadOnly Property DisplayName As String Implements vmAPI.vmInterface.DisplayName
        Get
            Return "EliteDangerousWatcher"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements vmAPI.vmInterface.Description
        Get
            Return "Watches Elite Dangerous Status.json and Journals, and publishes variables to VoiceMacro."
        End Get
    End Property

    Public ReadOnly Property ID As String Implements vmAPI.vmInterface.ID
        Get
            Return "EDPlugin.Watcher.001"
        End Get
    End Property

    ' -----------------------------
    ' vmInterface — lifecycle
    ' -----------------------------
    Public Sub New()
        Try
            _baseDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VM.ElitePlugin")
            If Not Directory.Exists(_baseDir) Then Directory.CreateDirectory(_baseDir)
            _log = New Logger(System.IO.Path.Combine(_baseDir, "VMPlugin.log"))
        Catch
            _baseDir = ""
            _log = New Logger("VMPlugin.log")
        End Try
    End Sub

    Public Sub Init() Implements vmAPI.vmInterface.Init
        WriteLog("Init()")

        VMSet("vm_log_checkbinds_s", "")
        ' Avoid ERR! surfacing on fresh profile for string helpers
        VMSet("vm_rp_state_s", "off")         ' small, non-empty default
        VMSet("vm_presence_oneshot_s", "-")   ' was "" → empty causes ERR! in VM
        VMSet("vm_ai_last_action_s", "Idle")   ' was "" → empty causes ERR!
        VMSet("vm_target_locked", "0")         ' seed to avoid ERR! before first event


        Try
            _journalDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                                                 "Saved Games", "Frontier Developments", "Elite Dangerous")
            _statusPath = System.IO.Path.Combine(_journalDir, "Status.json")
            WriteLog("Journal dir: " & _journalDir)
            WriteLog("Status path: " & _statusPath)
            VMSet("vm_journal_dir", _journalDir)
            VMSet("vm_status_path", _statusPath)
        Catch ex As Exception
            WriteLog("Init path error: " & ex.Message)
        End Try

        _handlers = New EventHandlers(AddressOf VMSet, AddressOf VMGet, AddressOf WriteLog)

        ' ---------- Remote config load + client init ----------
        Try
            _remoteConfigPath = System.IO.Path.Combine(_baseDir, "VMRemote.json")
        Catch
            _remoteConfigPath = "VMRemote.json"
        End Try

        LoadRemoteConfig()

        ' VM Remote client
        _remote = New VmRemoteClient(AddressOf WriteLog)
        Try
            _remote.SetBaseUrl(_remoteBaseUrl)
            _remote.SetDebounce(_remoteDebounceMs)
        Catch
        End Try


        ' Publish VM-visible mirrors
        VMSet("vm_remote_baseurl_s", _remoteBaseUrl)
        VMSet("vm_remote_debounce_ms", _remoteDebounceMs.ToString())

        ' ---------- Normal cold loads ----------
        ColdLoadStatus()
        ColdLoadLatestJournal()

        StartStatusWatcher()
        StartJournalWatcher()

        _machina = New MachinaBridge(AddressOf WriteLog)
        _presence = New DiscordPresence(AddressOf PublishPresence)
        _presence.SetVmGetter(AddressOf VMGet)
        StartHeartbeat()

        VMSet("vm_plugin_loaded", "1")
        VMSet("vm_plugin_loaded_s", "1")

        VMSet("vm_audit_state_s", "stopped")
        PrimeCoreMirrors()
    End Sub

    ' >>> BEGIN ReprimeEnvironment (inserted)
    ' Re-publish environment/mirrors on profile switch so *_p never shows ERR!
    Private Sub ReprimeEnvironment()
        Try
            VMSet("vm_status_path", _statusPath)
            VMSet("vm_journal_dir", _journalDir)
            VMSet("vm_active_journal", If(_activeJournal Is Nothing, "", System.IO.Path.GetFileName(_activeJournal)))
            VMSet("vm_remote_baseurl_s", _remoteBaseUrl)
            VMSet("vm_remote_debounce_ms", _remoteDebounceMs.ToString())
        Catch ex As Exception
            WriteLog("ReprimeEnvironment error: " & ex.Message)
        End Try
    End Sub
    ' <<< END ReprimeEnvironment


    Public Sub Dispose() Implements vmAPI.vmInterface.Dispose
        WriteLog("Dispose()")
        Try
            If _statusWatcher IsNot Nothing Then
                RemoveHandler _statusWatcher.Changed, AddressOf OnStatusChanged
                RemoveHandler _statusWatcher.Created, AddressOf OnStatusChanged
                RemoveHandler _statusWatcher.Renamed, AddressOf OnStatusRenamed
                _statusWatcher.EnableRaisingEvents = False
                _statusWatcher.Dispose()
                _statusWatcher = Nothing
            End If
        Catch ex As Exception
            WriteLog("Dispose status watcher: " & ex.Message)
        End Try
        Try
            If _journalWatcher IsNot Nothing Then
                RemoveHandler _journalWatcher.Changed, AddressOf OnJournalChanged
                RemoveHandler _journalWatcher.Created, AddressOf OnJournalChanged
                RemoveHandler _journalWatcher.Renamed, AddressOf OnJournalRenamed
                _journalWatcher.EnableRaisingEvents = False
                _journalWatcher.Dispose()
                _journalWatcher = Nothing
            End If
        Catch ex As Exception
            WriteLog("Dispose journal watcher: " & ex.Message)
        End Try

        StopHeartbeat()
        Try
            If _presence IsNot Nothing Then _presence.StopPresence()
        Catch
        End Try

        StopAuditTimer()
    End Sub

    Public Sub ProfileSwitched(ProfileGUID As String, ProfileName As String) Implements vmAPI.vmInterface.ProfileSwitched
        WriteLog($"ProfileSwitched: {ProfileGUID}, {ProfileName}")
        Try
            ' remember active VM profile name for path-style ExecuteMacro
            _vmProfileName = If(ProfileName, "")

            ' 1) Make sure all macro-visible *_p mirrors exist in the new VM profile
            PrimeCoreMirrors()

            ' 2) Re-publish environment values so *_p don’t show ERR! after switch
            ReprimeEnvironment()

            ' 3) Refresh features that depend on these mirrors
            RefreshPresenceSnapshot()
            RouterCheckGuiFocus()
        Catch ex As Exception
            WriteLog("ProfileSwitched handler error: " & ex.Message)
        End Try
    End Sub



    Public Sub ReceiveParams(Param1 As String, Param2 As String, Param3 As String, Synchron As Boolean) Implements vmAPI.vmInterface.ReceiveParams
        Try
            Dim cmdRaw As String = If(Param1, "")
            Dim sbc As New StringBuilder()
            For Each ch As Char In cmdRaw
                If Not Char.IsControl(ch) Then sbc.Append(ch)
            Next
            Dim cmd As String = sbc.ToString().Trim()
            If cmd.Length >= 2 Then
                If (cmd.StartsWith("""") AndAlso cmd.EndsWith("""")) OrElse (cmd.StartsWith("'") AndAlso cmd.EndsWith("'")) Then
                    cmd = cmd.Substring(1, cmd.Length - 2)
                End If
            End If

            Dim verb As String = ""
            Dim arg1 As String = ""
            If cmd.Length > 0 Then
                Dim parts = cmd.Split(New Char() {" "c}, 2, StringSplitOptions.RemoveEmptyEntries)
                verb = parts(0).ToLowerInvariant()
                If parts.Length > 1 Then arg1 = parts(1).Trim()
            End If

            VMSet("vm_dbg_last_cmd_s", If(cmd, ""))
            WriteLog($"ReceiveParams: raw='{cmdRaw}', verb='{verb}', arg1='{arg1}', Synchron={Synchron}")

            Select Case verb

                Case "selfteststatus"
                    Dim hud As String = VMGet("vm_hud_mode") : If String.IsNullOrEmpty(hud) Then hud = "?"
                    Dim fg As String = VMGet("vm_firegroup") : If String.IsNullOrEmpty(fg) Then fg = "?"
                    Dim sys As String = VMGet("vm_star_system")
                    Dim onf As String = VMGet("vm_on_foot")
                    VMSet("vm_self_test_result_s", $"HUD={hud}, FireGroup={fg}, Sys={sys}, OnFoot={onf}")

                Case "checkbinds", "checkbindings", "checkbind"
                    VMSet("vm_log_checkbinds_s", "CheckBinds: running...")
                    WriteLog("ReceiveParams: invoking CheckKeybindsAndLog()")
                    CheckKeybindsAndLog()

                ' ----- Machina toggles -----
                Case "aion"
                    _machinaEnabled = True
                    _machina.SetEnabled(True)
                    VMSet("vm_ai_last_action_s", "AI route: ON")

                Case "aioff"
                    _machinaEnabled = False
                    _machina.SetEnabled(False)
                    VMSet("vm_ai_last_action_s", "AI route: OFF")

                Case "aitoggle"
                    _machinaEnabled = Not _machinaEnabled
                    _machina.SetEnabled(_machinaEnabled)
                    VMSet("vm_ai_last_action_s", If(_machinaEnabled, "AI route: ON", "AI route: OFF"))

                Case "machinasnapshot"
                    Dim snap As New System.Collections.Generic.Dictionary(Of String, String)()
                    snap("environment") = VMGet("vm_environment")
                    snap("system") = VMGet("vm_star_system")
                    snap("body") = VMGet("vm_body")
                    snap("station") = VMGet("vm_station_name")
                    snap("on_foot") = VMGet("vm_on_foot")
                    snap("docked") = VMGet("vm_docked")
                    snap("in_supercruise") = VMGet("vm_in_supercruise")
                    snap("hud_mode") = VMGet("vm_hud_mode")
                    snap("firegroup") = VMGet("vm_firegroup")
                    _machina.PublishSnapshot(snap)
                    VMSet("vm_ai_last_action_s", "AI snapshot published")

                                ' ----- Presence -----
                Case "rp_off"
                    Try
                        ' Local snapshots to avoid races if other threads change fields
                        Dim presenceLocal = _presence
                        Dim machinaLocal = _machina

                        ' Stop presence safely
                        If presenceLocal IsNot Nothing Then
                            Try
                                presenceLocal.StopPresence()
                            Catch ex As Exception
                                WriteLog("[Presence] StopPresence error: " & ex.Message & " | " & ex.StackTrace)
                            End Try
                        End If

                        ' Update VM state
                        Try
                            VMSet("vm_rp_state_s", "off")
                        Catch ex As Exception
                            WriteLog("[Presence] VMSet(rp_state) error: " & ex.Message & " | " & ex.StackTrace)
                        End Try

                        ' Clear any lingering presence card in external consumers
                        If machinaLocal IsNot Nothing Then
                            Try
                                Dim payload As New System.Collections.Generic.Dictionary(Of String, String) From {
                                    {"kind", "presence"},
                                    {"mode", "off"},
                                    {"line1", ""},
                                    {"line2", ""},
                                    {"timestamp", DateTime.UtcNow.ToString("o")}
                                }
                                machinaLocal.PublishSnapshot(payload)
                            Catch ex As Exception
                                WriteLog("[Presence] PublishSnapshot(off) error: " & ex.Message & " | " & ex.StackTrace)
                            End Try
                        End If

                    Catch ex As Exception
                        ' Catch-all to prevent any unexpected exception from escaping this case handler
                        WriteLog("[Presence] rp_off unexpected error: " & ex.Message & " | " & ex.StackTrace)
                    End Try

                Case "lfw_on"
                    Dim sys As String = VMGet("vm_star_system")
                    If String.IsNullOrEmpty(sys) Then sys = "Unknown"
                    Dim msg As String = "LFW — " & sys
                    VMSet("vm_presence_oneshot_s", msg)
                    If _presence IsNot Nothing Then _presence.StartPresence()
                    VMSet("vm_rp_state_s", "on")
                    VMSet("vm_rp_last_action_s", "lfw_on")

                Case "lfw_off"
                    VMSet("vm_presence_oneshot_s", "")
                    VMSet("vm_rp_last_action_s", "lfw_off")


                ' ----- Types / Audit -----
                Case "vmvartypes", "types"
                    WriteVmVarTypes()

                Case "auditstatus"
                    WriteAuditSnapshot(writeCsv:=False)

                Case "auditcsv"
                    WriteAuditSnapshot(writeCsv:=True)

                Case "auditstart"
                    StartAuditTimer()

                Case "auditstop"
                    StopAuditTimer()

                Case "auditscope"
                    If String.IsNullOrEmpty(arg1) Then
                        VMSet("vm_log_checkbinds_s", "auditscope usage: auditscope <varName>")
                    Else
                        AuditscopeProbe(arg1)
                    End If

                Case "edbinds"
                    Try
                        If BindReader.ScanAndPublish(AddressOf WriteLog) Then
                            WriteLog("[BindReader] scan ok")
                        Else
                            WriteLog("[BindReader] scan completed with errors; see ed_status/ed_error")
                        End If
                    Catch ex As Exception
                        WriteLog("[BindReader] scan failed: " & ex.Message)
                    End Try

                    Return

                Case "edreport"
                    Try
                        Dim rep As String = VMApiShim.GetVariable("ed_report", AddressOf WriteLog)
                        If String.IsNullOrEmpty(rep) Then rep = "(ed_report is empty)"
                        VMSet("vm_log_checkbinds_s", rep)
                    Catch ex As Exception
                        WriteLog("[BindReader] edreport failed: " & ex.Message)
                    End Try
                    Return

                Case "edpath"
                    Try
                        Dim p As String = VMApiShim.GetVariable("ed_bindspath", AddressOf WriteLog)
                        Dim st As String = VMApiShim.GetVariable("ed_status", AddressOf WriteLog)
                        Dim errMsg As String = VMApiShim.GetVariable("ed_error", AddressOf WriteLog)
                        Dim line As New System.Text.StringBuilder()
                        line.Append("ed_bindspath=").Append(p).Append(" | ed_status=").Append(st)
                        If Not String.IsNullOrEmpty(errMsg) Then
                            line.Append(" | ed_error=").Append(errMsg)
                        End If
                        VMSet("vm_log_checkbinds_s", line.ToString())
                    Catch ex As Exception
                        WriteLog("[BindReader] edpath failed: " & ex.Message)
                    End Try
                    Return

                ' ----- Remote config (URL/debounce only) -----
                Case "setremote"
                    HandleSetRemote(arg1)

                Case "setremotedebounce"
                    HandleSetRemoteDebounce(arg1)

                Case "getremote"
                    Dim info = $"Remote: URL={_remoteBaseUrl}, debounce={_remoteDebounceMs} ms"
                    VMSet("vm_log_checkbinds_s", info)
                    WriteLog("[Cmd] " & info)

                Case "helpremote"
                    WriteRemoteHelp()

                Case "remotecheck"
                    HandleRemoteCheck(arg1)

                Case "getvar" : HandleGetVar(arg1) : Return

                Case Else
                    VMSet("vm_log_checkbinds_s", $"Unknown command: {If(cmd, "")}")
                    WriteLog($"ReceiveParams: unknown command '{cmd}'")

            End Select

        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "ReceiveParams error: " & ex.Message)
            WriteLog("ReceiveParams error: " & ex.ToString())
        End Try
    End Sub

    ' -----------------------------
    ' Remote config helpers
    ' -----------------------------
    Private Sub LoadRemoteConfig()
        Try
            If String.IsNullOrEmpty(_remoteConfigPath) OrElse Not File.Exists(_remoteConfigPath) Then Exit Sub
            Dim txt = File.ReadAllText(_remoteConfigPath, Encoding.UTF8)
            Dim jo = JObject.Parse(txt)

            Dim url As String = CStr(jo("baseUrl"))
            Dim db As Integer = 0
            Integer.TryParse(CStr(jo("debounceMs")), db)

            Dim u As Uri = Nothing
            If Not String.IsNullOrWhiteSpace(url) AndAlso Uri.TryCreate(url, UriKind.Absolute, u) _
               AndAlso (u.Scheme = Uri.UriSchemeHttp OrElse u.Scheme = Uri.UriSchemeHttps) Then
                If Not url.EndsWith("/") Then url &= "/"
                _remoteBaseUrl = url
            End If

            If db >= 100 AndAlso db <= 5000 Then
                _remoteDebounceMs = db
            End If

            WriteLog($"[RemoteCfg] Loaded {_remoteBaseUrl} @ {_remoteDebounceMs} ms")
        Catch ex As Exception
            WriteLog($"[RemoteCfg] Load error: {ex.Message}")
        End Try
    End Sub

    Private Sub SaveRemoteConfig()
        Try
            Dim jo As New JObject(
                New JProperty("baseUrl", _remoteBaseUrl),
                New JProperty("debounceMs", _remoteDebounceMs)
            )
            Dim tmp = _remoteConfigPath & ".tmp"
            File.WriteAllText(tmp, jo.ToString(Formatting.Indented), Encoding.UTF8)
            If File.Exists(_remoteConfigPath) Then
                File.Replace(tmp, _remoteConfigPath, _remoteConfigPath & ".bak")
            Else
                File.Move(tmp, _remoteConfigPath)
            End If
            WriteLog($"[RemoteCfg] Saved {_remoteConfigPath}")
        Catch ex As Exception
            WriteLog($"[RemoteCfg] Save error: {ex.Message}")
        End Try
    End Sub

    Private Sub HandleSetRemote(arg As String)
        Dim url As String = If(arg, "").Trim()
        Dim ok As Boolean = False
        If url <> "" Then
            Dim u As Uri = Nothing
            If Uri.TryCreate(url, UriKind.Absolute, u) AndAlso
               (u.Scheme = Uri.UriSchemeHttp OrElse u.Scheme = Uri.UriSchemeHttps) Then
                If Not url.EndsWith("/") Then url &= "/"
                SyncLock _remoteLock
                    _remoteBaseUrl = url
                    Try
                        If _remote IsNot Nothing Then _remote.SetBaseUrl(_remoteBaseUrl)
                    Catch
                    End Try
                End SyncLock
                VMSet("vm_remote_baseurl_s", _remoteBaseUrl)
                SaveRemoteConfig()
                ok = True
            End If
        End If
        Dim msg = If(ok, $"Remote base url set to: {_remoteBaseUrl}",
                         "Error: setremote expects a valid http(s) URL (example: http://localhost:8080/).")
        VMSet("vm_log_checkbinds_s", msg)
        WriteLog("[Cmd] " & msg)
    End Sub

    Private Sub HandleSetRemoteDebounce(arg As String)
        Dim s As String = If(arg, "").Trim()
        Dim v As Integer
        Dim ok As Boolean = Integer.TryParse(s, v) AndAlso v >= 100 AndAlso v <= 5000
        Dim msg As String
        If ok Then
            SyncLock _remoteLock
                _remoteDebounceMs = v
            End SyncLock
            Try : _remote.SetDebounce(_remoteDebounceMs) : Catch : End Try
            VMSet("vm_remote_debounce_ms", _remoteDebounceMs.ToString())
            SaveRemoteConfig()
            msg = $"Remote debounce set to: {_remoteDebounceMs} ms"
        Else
            msg = "Error: setremotedebounce expects 100–5000 (ms)."
        End If
        VMSet("vm_log_checkbinds_s", msg)
        WriteLog("[Cmd] " & msg)
    End Sub

    Private Sub WriteRemoteHelp()
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("VM Remote — quick help")
            sb.AppendLine("Commands (Call External Plug-In → Param1):")
            sb.AppendLine("  setremote <url>          : Set base URL. Example: setremote http://localhost:8080/")
            sb.AppendLine("  setremotedebounce <ms>   : Set debounce 100–5000 ms. Example: setremotedebounce 350")
            sb.AppendLine("  getremote                : Show current URL + debounce.")
            sb.AppendLine("  remotecheck [macro]      : Send a test macro (default: ED_CommsOpen)")
            sb.AppendLine("  helpremote               : Show this help.")
            sb.AppendLine()
            sb.AppendLine("Notes:")
            sb.AppendLine(" • VoiceMacro → Settings → Remote Control must be ON (default port 8080).")
            sb.AppendLine(" • Macros shipped & expected in VoiceMacro: ED_CommsOpen, ED_CommsClose, ED_GalaxyOpen, ED_GalaxyClose.")
            VMSet("vm_log_checkbinds_s", sb.ToString())
            WriteLog("[Help] helpremote shown")
        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "helpremote error: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleRemoteCheck(arg As String)
        Try
            Dim macro As String = If(String.IsNullOrWhiteSpace(arg), "ED_CommsOpen", arg.Trim())
            If _remote Is Nothing Then
                VMSet("vm_log_checkbinds_s", "remotecheck: VM Remote client not initialized.")
                Return
            End If
            _remote.TrySendWithProfile(_vmProfileName, macro)
            Dim msg As String = $"remotecheck: sent '{macro}' @ {_remoteBaseUrl}"
            VMSet("vm_log_checkbinds_s", msg)
            WriteLog("[Cmd] " & msg)
        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "remotecheck error: " & ex.Message)
            WriteLog("remotecheck error: " & ex.ToString())
        End Try
    End Sub
    Private Sub HandleGetVar(arg As String)
        Try
            Dim name As String = If(arg, "").Trim()
            If name.Length = 0 Then
                VMSet("vm_log_checkbinds_s", "getvar usage: getvar <varName>")
                VMSet("vm_getvar_value_s", "")
                Return
            End If

            ' strip surrounding quotes if present
            If (name.StartsWith("""") AndAlso name.EndsWith("""")) OrElse (name.StartsWith("'") AndAlso name.EndsWith("'")) Then
                name = name.Substring(1, name.Length - 2).Trim()
            End If

            ' read both bare and _p (macro-visible) values
            Dim bareRaw As String = VMApiShim.GetVariable(name, AddressOf WriteLog)
            Dim pRaw As String = VMApiShim.GetVariable(name & "_p", AddressOf WriteLog)

            Dim bare As String = NormalizeVmRead(bareRaw)
            Dim persisted As String = NormalizeVmRead(pRaw)

            ' prefer _p when present; otherwise fall back to bare
            Dim picked As String = If(persisted <> "", persisted, bare)

            ' publish the value for VoiceMacro to read
            VMSet("vm_getvar_value_s", picked)

            ' optional: human-readable trace to your debug string var
            VMSet("vm_log_checkbinds_s", $"getvar {name}: bare='{bare}' _p='{persisted}' -> '{picked}'")
            WriteLog("[getvar] " & $"getvar {name}: bare='{bare}' _p='{persisted}' -> '{picked}'")

        Catch ex As Exception
            VMSet("vm_getvar_value_s", "")
            VMSet("vm_log_checkbinds_s", "getvar error: " & ex.Message)
            WriteLog("[getvar] error: " & ex.ToString())
        End Try
    End Sub

    ' -----------------------------
    ' Cold loads
    ' -----------------------------
    Private Sub ColdLoadStatus()
        Try
            If Not File.Exists(_statusPath) Then Exit Sub
            Dim text As String = SafeReadAllText(_statusPath)
            If Not String.IsNullOrWhiteSpace(text) Then
                _handlers.ProcessStatusJson(text)
                CheckPresenceOneShot()
                RefreshPresenceSnapshot()
                RouterCheckGuiFocus() ' run router once on cold state
                WriteLog("ColdLoad: Status.json processed")
            End If
        Catch ex As Exception
            WriteLog("ColdLoadStatus error: " & ex.Message)
        End Try
    End Sub

    Private Sub ColdLoadLatestJournal()
        Try
            If String.IsNullOrWhiteSpace(_journalDir) OrElse Not Directory.Exists(_journalDir) Then Exit Sub
            Dim latest As String = GetLatestJournalFile()
            If latest Is Nothing Then Exit Sub

            _activeJournal = latest
            VMSet("vm_active_journal", System.IO.Path.GetFileName(latest))
            _tailPosition = 0

            Dim processed As Integer = 0
            Using fs As New FileStream(latest, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using sr As New StreamReader(fs, Encoding.UTF8)
                    While True
                        Dim ln = sr.ReadLine()
                        If ln Is Nothing Then Exit While
                        If ln.Length > 0 Then
                            _handlers.ProcessJournalLine(ln)
                            processed += 1
                        End If
                    End While
                End Using
            End Using

            CheckPresenceOneShot()
            RefreshPresenceSnapshot()
            WriteLog("ColdLoad: processed " & processed & " lines from " & System.IO.Path.GetFileName(latest))

            Dim fi = New FileInfo(latest)
            _tailPosition = fi.Length
            _lastJournalWriteUtc = fi.LastWriteTimeUtc
        Catch ex As Exception
            WriteLog("ColdLoadLatestJournal error: " & ex.Message)
        End Try
    End Sub

    ' -----------------------------
    ' Watchers
    ' -----------------------------
    Private Sub StartStatusWatcher()
        Try
            Dim dir = System.IO.Path.GetDirectoryName(_statusPath)
            If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Exit Sub
            _statusWatcher = New FileSystemWatcher(dir, "Status.json") With {
                .NotifyFilter = NotifyFilters.LastWrite Or NotifyFilters.Size Or NotifyFilters.FileName,
                .IncludeSubdirectories = False,
                .EnableRaisingEvents = True
            }
            AddHandler _statusWatcher.Changed, AddressOf OnStatusChanged
            AddHandler _statusWatcher.Created, AddressOf OnStatusChanged
            AddHandler _statusWatcher.Renamed, AddressOf OnStatusRenamed
            WriteLog("Status watcher started")
        Catch ex As Exception
            WriteLog("StartStatusWatcher error: " & ex.Message)
        End Try
    End Sub

    Private Sub StartJournalWatcher()
        Try
            If String.IsNullOrWhiteSpace(_journalDir) OrElse Not Directory.Exists(_journalDir) Then Exit Sub
            _journalWatcher = New FileSystemWatcher(_journalDir, "Journal*.log") With {
                .NotifyFilter = NotifyFilters.LastWrite Or NotifyFilters.FileName Or NotifyFilters.CreationTime,
                .IncludeSubdirectories = False,
                .EnableRaisingEvents = True
            }
            AddHandler _journalWatcher.Changed, AddressOf OnJournalChanged
            AddHandler _journalWatcher.Created, AddressOf OnJournalChanged
            AddHandler _journalWatcher.Renamed, AddressOf OnJournalRenamed
            WriteLog("Journal watcher started")
        Catch ex As Exception
            WriteLog("StartJournalWatcher error: " & ex.Message)
        End Try
    End Sub

    Private Sub OnStatusChanged(sender As Object, e As FileSystemEventArgs)
        Dim due = Interlocked.Increment(_statusDebounceDue)
        ThreadPool.QueueUserWorkItem(
            Sub()
                Thread.Sleep(50)
                If Interlocked.Decrement(_statusDebounceDue) = 0 Then
                    Try
                        Dim txt = SafeReadAllText(e.FullPath)
                        If Not String.IsNullOrWhiteSpace(txt) Then
                            _handlers.ProcessStatusJson(txt)
                            CheckPresenceOneShot()
                            RefreshPresenceSnapshot()
                            RouterCheckGuiFocus()
                        End If
                    Catch ex As Exception
                        WriteLog("OnStatusChanged error: " & ex.Message)
                    End Try
                End If
            End Sub)
    End Sub

    Private Sub OnStatusRenamed(sender As Object, e As RenamedEventArgs)
        OnStatusChanged(sender, New FileSystemEventArgs(WatcherChangeTypes.Changed,
                                                        System.IO.Path.GetDirectoryName(e.FullPath),
                                                        System.IO.Path.GetFileName(e.FullPath)))
    End Sub

    Private Sub OnJournalChanged(sender As Object, e As FileSystemEventArgs)
        Dim due = Interlocked.Increment(_journalDebounceDue)
        ThreadPool.QueueUserWorkItem(
            Sub()
                Thread.Sleep(60)
                If Interlocked.Decrement(_journalDebounceDue) = 0 Then
                    Try
                        ProcessJournalTail()
                        CheckPresenceOneShot()
                        RefreshPresenceSnapshot()
                        ' NOTE: No ExecuteMacro fires here (restricted to GUI focus router only)
                    Catch ex As Exception
                        WriteLog("OnJournalChanged error: " & ex.Message)
                    End Try
                End If
            End Sub)
    End Sub

    Private Sub OnJournalRenamed(sender As Object, e As RenamedEventArgs)
        OnJournalChanged(sender, New FileSystemEventArgs(WatcherChangeTypes.Changed,
                                                         System.IO.Path.GetDirectoryName(e.FullPath),
                                                         System.IO.Path.GetFileName(e.FullPath)))
    End Sub

    ' -----------------------------
    ' GUI Focus Router: Comms/Galaxy ONLY
    ' -----------------------------
    Private Sub RouterCheckGuiFocus()
        Try
            Dim curName As String = VMGet("vm_gui_focus_name")
            If String.IsNullOrEmpty(curName) Then curName = ""

            Dim prev As String = _lastGuiFocusName
            If String.Equals(curName, prev, StringComparison.Ordinal) Then Exit Sub

            _lastGuiFocusName = curName

            ' Handle CommsPanel transitions
            Dim curComms As Boolean = String.Equals(curName, "CommsPanel", StringComparison.OrdinalIgnoreCase)
            Dim prevComms As Boolean = String.Equals(prev, "CommsPanel", StringComparison.OrdinalIgnoreCase)
            If curComms AndAlso Not prevComms Then
                WriteLog("Router: ED_CommsOpen")
                FireVmMacro("ED_CommsOpen")
            ElseIf Not curComms AndAlso prevComms Then
                WriteLog("Router: ED_CommsClose")
                FireVmMacro("ED_CommsClose")
            End If

            ' Handle GalaxyMap transitions
            Dim curGalaxy As Boolean = String.Equals(curName, "GalaxyMap", StringComparison.OrdinalIgnoreCase)
            Dim prevGalaxy As Boolean = String.Equals(prev, "GalaxyMap", StringComparison.OrdinalIgnoreCase)
            If curGalaxy AndAlso Not prevGalaxy Then
                WriteLog("Router: ED_GalaxyOpen")
                FireVmMacro("ED_GalaxyOpen")
            ElseIf Not curGalaxy AndAlso prevGalaxy Then
                WriteLog("Router: ED_GalaxyClose")
                FireVmMacro("ED_GalaxyClose")
            End If

        Catch ex As Exception
            WriteLog("RouterCheckGuiFocus error: " & ex.Message)
        End Try
    End Sub


    'FireVmMacro
    Private Sub FireVmMacro(macroName As String)
        Try
            If _remote Is Nothing Then Exit Sub
            If String.IsNullOrWhiteSpace(macroName) Then Exit Sub

            ' Use profile-name path: /ExecuteMacro=<ProfileName>/<Macro>
            _remote.TrySendWithProfile(_vmProfileName, macroName)

        Catch ex As System.Net.WebException
            Dim http = TryCast(ex.Response, System.Net.HttpWebResponse)
            If http IsNot Nothing AndAlso CInt(http.StatusCode) = 400 Then
                WriteLog("[Remote][Hint] 400 Invalid Hostname. Try: setremote http://127.0.0.1:8080/")
            End If
            WriteLog("FireVmMacro WebException: " & ex.Message)
        Catch ex As Exception
            WriteLog("FireVmMacro error: " & ex.Message)
        End Try
    End Sub


    ' -----------------------------
    ' Presence one-shot
    ' -----------------------------
    Private Sub CheckPresenceOneShot()
        Try
            Dim one As String = VMGet("vm_presence_oneshot_s")
            ' Ignore the sentinel "-" (used to avoid VM ERR! on empty strings)
            If Not String.IsNullOrWhiteSpace(one) AndAlso Not one.Equals("-", StringComparison.Ordinal) Then
                If _presence IsNot Nothing Then
                    _presence.ShowOneShot(one)
                Else
                    WriteLog("[Presence OneShot] " & one)
                End If
                ' Reset back to the sentinel, not empty
                VMSet("vm_presence_oneshot_s", "-")
            End If
        Catch
        End Try
    End Sub


    ' -----------------------------
    ' Presence snapshot
    ' -----------------------------
    Private Sub RefreshPresenceSnapshot()
        Try
            If _presence Is Nothing Then Exit Sub

            Dim inWing As Boolean = (VMGet("vm_in_wing_p") = "1")
            _presence.SetWingState(inWing)

            Dim hud As String = VMGet("vm_hud_mode")
            Dim sys As String = VMGet("vm_star_system")
            Dim body As String = VMGet("vm_body")
            Dim docked As Boolean = (VMGet("vm_docked") = "1")
            Dim inSC As Boolean = (VMGet("vm_in_supercruise") = "1")
            Dim station As String = VMGet("vm_station_name")

            _presence.UpdateFromSnapshot(hud, sys, body, docked, inSC, station)

        Catch
        End Try
    End Sub

    ' -----------------------------
    ' Safe file IO
    ' -----------------------------
    Private Function SafeReadAllText(fullPath As String) As String
        For i As Integer = 1 To 5
            Try
                Using fs As New FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Using sr As New StreamReader(fs, Encoding.UTF8)
                        Return sr.ReadToEnd()
                    End Using
                End Using
            Catch ex As IOException
                Thread.Sleep(25)
            Catch ex As UnauthorizedAccessException
                Thread.Sleep(25)
            End Try
        Next
        Return Nothing
    End Function

    Private Function GetLatestJournalFile() As String
        Dim files = Directory.GetFiles(_journalDir, "Journal*.log", SearchOption.TopDirectoryOnly)
        If files Is Nothing OrElse files.Length = 0 Then Return Nothing
        Dim best As String = Nothing
        Dim bestUtc As DateTime = DateTime.MinValue
        For Each f In files
            Dim wt = File.GetLastWriteTimeUtc(f)
            If wt > bestUtc Then bestUtc = wt : best = f
        Next
        Return best
    End Function

    ' -----------------------------
    ' Journal tail (rotation-aware)
    ' -----------------------------
    Private Sub ProcessJournalTail()
        If String.IsNullOrWhiteSpace(_journalDir) OrElse Not Directory.Exists(_journalDir) Then Exit Sub
        SyncLock _journalLock
            Dim latest = GetLatestJournalFile()
            If latest Is Nothing Then Exit Sub

            Dim latestInfo = New FileInfo(latest)
            If _activeJournal Is Nothing OrElse
           Not String.Equals(_activeJournal, latest, StringComparison.OrdinalIgnoreCase) OrElse
           latestInfo.LastWriteTimeUtc > _lastJournalWriteUtc Then

                If _activeJournal Is Nothing OrElse Not String.Equals(_activeJournal, latest, StringComparison.OrdinalIgnoreCase) Then
                    WriteLog("Journal rotation -> " & System.IO.Path.GetFileName(latest))
                    _tailPosition = 0
                    VMSet("vm_active_journal", System.IO.Path.GetFileName(latest))
                End If
                _activeJournal = latest
                _lastJournalWriteUtc = latestInfo.LastWriteTimeUtc
            End If

            Try
                Using fs As New FileStream(_activeJournal, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                    ' Handle truncate/shrink (rare, but possible during rotation)
                    If _tailPosition < 0 OrElse _tailPosition > fs.Length Then
                        WriteLog($"Journal tail reset: file shrank (oldPos={_tailPosition}, newLen={fs.Length})")
                        _tailPosition = 0
                    End If

                    fs.Position = _tailPosition
                    Dim readCount As Integer = 0

                    Using sr As New StreamReader(fs, Encoding.UTF8)
                        While True
                            Dim line = sr.ReadLine()
                            If line Is Nothing Then Exit While
                            If line.Length > 0 Then
                                _handlers.ProcessJournalLine(line)
                                readCount += 1
                            End If
                        End While
                        _tailPosition = fs.Position
                    End Using

                    If readCount > 0 Then
                        WriteLog($"Journal read: +{readCount} lines (pos={_tailPosition})")
                    End If
                End Using

            Catch ex As IOException
                WriteLog($"Journal IOEX: {ex.GetType().Name} '{ex.Message}' path={_activeJournal} pos={_tailPosition}")
            Catch ex As UnauthorizedAccessException
                WriteLog($"Journal IOEX: {ex.GetType().Name} '{ex.Message}' path={_activeJournal} pos={_tailPosition}")
            End Try
        End SyncLock
    End Sub


    ' -----------------------------
    ' Keybind checking (file-log only)
    ' -----------------------------
    Private Sub CheckKeybindsAndLog()
        Try
            VMSet("vm_log_checkbinds_s", "CheckBinds: starting...")

            Dim bindsPath = FindActiveBindsPath()
            If String.IsNullOrEmpty(bindsPath) OrElse Not File.Exists(bindsPath) Then
                WriteLog("Keybind check: binds file not found. Start the game and open Options once, then re-run the check.")
                WriteLog("%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                VMSet("vm_log_checkbinds_s",
                  "Keybind check: binds file not found. Start the game and open Options once, then re-run the check." & Environment.NewLine &
                  "%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                Return
            End If

            Dim xdoc As System.Xml.Linq.XDocument = Nothing
            Dim root As System.Xml.Linq.XElement = Nothing

            Try
                xdoc = System.Xml.Linq.XDocument.Load(bindsPath)
                If xdoc Is Nothing OrElse xdoc.Root Is Nothing Then
                    WriteLog("Keybind check: failed to parse binds XML (empty root).")
                    WriteLog("%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                    VMSet("vm_log_checkbinds_s",
                      "Keybind check: failed to parse binds XML." & Environment.NewLine &
                      "%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                    Return
                End If
                root = xdoc.Root
            Catch ex As Exception
                WriteLog("Keybind check: failed to read/parse binds XML: " & ex.Message)
                WriteLog("%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                VMSet("vm_log_checkbinds_s",
                  "Keybind check: failed to read/parse binds XML." & Environment.NewLine &
                  "%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
                Return
            End Try

            Dim requiredTags As String() = {
                "CycleFireGroupNext",
                "CycleFireGroupPrevious",
                "PlayerHUDModeToggle",
                "Hyperspace",
                "Supercruise",
                "LandingGearToggle",
                "DeployHardpointToggle",
                "ToggleCargoScoop"
            }

            Dim missing As New System.Collections.Generic.List(Of String)
            For Each tag In requiredTags
                Dim act As System.Xml.Linq.XElement = root.Element(tag)
                If act Is Nothing OrElse Not IsBindingPresent(act) Then
                    missing.Add(tag)
                End If
            Next

            WriteLog("Keybind check: " & System.IO.Path.GetFileName(bindsPath))
            If missing.Count > 0 Then
                WriteLog("Keybinds missing: " & String.Join(", ", missing.ToArray()))
            Else
                WriteLog("Keybinds: all required present.")
            End If
            WriteLog("%LocalAppData%\VM.ElitePlugin\VMPlugin.log")

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("Keybind check: " & System.IO.Path.GetFileName(bindsPath))
            If missing.Count > 0 Then
                sb.AppendLine("Keybinds missing:")
                For Each m As String In missing
                    sb.AppendLine(" - " & m)
                Next
            Else
                sb.AppendLine("Keybinds: all required present.")
            End If
            sb.AppendLine("%LocalAppData%\VM.ElitePlugin\VMPlugin.log")
            VMSet("vm_log_checkbinds_s", sb.ToString())

        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "CheckBinds failed: " & ex.Message)
            WriteLog("Keybind check error: " & ex.ToString())
        End Try
    End Sub

    Private Function FindActiveBindsPath() As String
        Try
            Dim optionsDir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Frontier Developments", "Elite Dangerous", "Options", "Bindings")
            If Not System.IO.Directory.Exists(optionsDir) Then Return Nothing

            Dim startPath = System.IO.Path.Combine(optionsDir, "StartPreset.start")
            If System.IO.File.Exists(startPath) Then
                Dim preset As String = Nothing
                For Each ln In System.IO.File.ReadLines(startPath, System.Text.Encoding.UTF8)
                    Dim t = ln.Trim()
                    If t.Length > 0 Then preset = t : Exit For
                Next
                If Not String.IsNullOrEmpty(preset) Then
                    Dim candidate = System.IO.Path.Combine(optionsDir, preset & ".binds")
                    If System.IO.File.Exists(candidate) Then Return candidate
                End If
            End If

            Dim files = System.IO.Directory.GetFiles(optionsDir, "*.binds", System.IO.SearchOption.TopDirectoryOnly)
            If files Is Nothing OrElse files.Length = 0 Then Return Nothing
            Dim mostRecent = files.OrderByDescending(Function(f) System.IO.File.GetLastWriteTimeUtc(f)).First()
            Return mostRecent
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function IsBindingPresent(act As System.Xml.Linq.XElement) As Boolean
        Try
            Dim b = act.Element("Binding")
            If b IsNot Nothing Then
                Dim dev = CStr(b.Attribute("Device"))
                Dim key = CStr(b.Attribute("Key"))
                If Not String.IsNullOrEmpty(dev) AndAlso Not dev.Equals("{NoDevice}", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.IsNullOrEmpty(key) Then
                    Return True
                End If
            End If

            Dim p = act.Element("Primary")
            If p IsNot Nothing Then
                Dim dev = CStr(p.Attribute("Device"))
                Dim key = CStr(p.Attribute("Key"))
                If Not String.IsNullOrEmpty(dev) AndAlso Not dev.Equals("{NoDevice}", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.IsNullOrEmpty(key) Then
                    Return True
                End If
            End If

            Dim s = act.Element("Secondary")
            If s IsNot Nothing Then
                Dim dev = CStr(s.Attribute("Device"))
                Dim key = CStr(s.Attribute("Key"))
                If Not String.IsNullOrEmpty(dev) AndAlso Not dev.Equals("{NoDevice}", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not String.IsNullOrEmpty(key) Then
                    Return True
                End If
            End If

            Return False
        Catch
            Return False
        End Try
    End Function

    ' ============================================================
    '                     AUDIT / STATUS (Dump)
    ' ============================================================
    Private Sub StartAuditTimer()
        Try
            If _auditTimer Is Nothing Then
                _auditTimer = New System.Threading.Timer(
                    Sub(state As Object)
                        Try
                            WriteAuditSnapshot(writeCsv:=False)
                        Catch
                        End Try
                    End Sub,
                    Nothing, _auditPeriodMs, _auditPeriodMs)
            Else
                _auditTimer.Change(_auditPeriodMs, _auditPeriodMs)
            End If
            _auditRunning = True
            VMSet("vm_audit_state_s", "running")
            WriteLog("Audit timer started")
        Catch ex As Exception
            WriteLog("Audit timer error: " & ex.Message)
        End Try
    End Sub

    Private Sub StopAuditTimer()
        Try
            If _auditTimer IsNot Nothing Then
                _auditTimer.Change(Timeout.Infinite, Timeout.Infinite)
                _auditTimer.Dispose()
                _auditTimer = Nothing
            End If
            _auditRunning = False
            VMSet("vm_audit_state_s", "stopped")
            WriteLog("Audit timer stopped")
        Catch ex As Exception
            WriteLog("Audit stop error: " & ex.Message)
        End Try
    End Sub

    Private Sub AuditscopeProbe(varName As String)
        Try
            If String.IsNullOrEmpty(varName) Then
                VMSet("vm_log_checkbinds_s", "auditscope usage: auditscope <varName>")
                Return
            End If

            Dim bareRaw = VMApiShim.GetVariable(varName, AddressOf WriteLog)
            Dim persRaw = VMApiShim.GetVariable(varName & "_p", AddressOf WriteLog)

            Dim bare = NormalizeVmRead(bareRaw)
            Dim pers = NormalizeVmRead(persRaw)

            Dim macroVisible As String = If(String.IsNullOrEmpty(pers), "No", "Yes")
            Dim report = $"auditscope {varName}: bare='{bare}', _p='{pers}', MacroVisible={macroVisible}"
            WriteLog(report)
            VMSet("vm_log_checkbinds_s", report)
        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "auditscope error: " & ex.Message)
        End Try
    End Sub

    Private Sub WriteAuditSnapshot(Optional writeCsv As Boolean = False)
        Try
            If String.IsNullOrEmpty(_baseDir) Then
                VMSet("vm_log_checkbinds_s", "Audit: base dir not set")
                Return
            End If

            Dim txtPath As String = System.IO.Path.Combine(_baseDir, "StatusDump.txt")
            Dim jsonPath As String = System.IO.Path.Combine(_baseDir, "StatusDump.json")
            Dim csvPath As String = System.IO.Path.Combine(_baseDir, "StatusDump.csv")

            Dim entries = _audit.Values.OrderBy(Function(e) e.Name, StringComparer.OrdinalIgnoreCase).ToList()

            Dim lines As New List(Of String)()
            lines.Add("Status Dump - " & DateTime.UtcNow.ToString("u"))
            lines.Add("Totals: " & entries.Count & " vars")
            lines.Add("Name | Type | Updated (UTC) | Source | Value")
            lines.Add(New String("-"c, 100))

            For Each e In entries
                Dim src As String = If(e.SeenWrite AndAlso e.SeenRead, "Write+Read", If(e.SeenWrite, "Write", If(e.SeenRead, "Read", "")))
                Dim val = SanitizeOneLine(If(e.Value, ""))
                lines.Add($"{e.Name} | {e.TypeName} | {If(e.UpdatedUtc = DateTime.MinValue, "", e.UpdatedUtc.ToString("u"))} | {src} | {Truncate(val, 256)}")
            Next
            File.WriteAllLines(txtPath, lines.ToArray(), Encoding.UTF8)

            Dim ja As New JArray()
            For Each e In entries
                Dim o As New JObject()
                o("name") = e.Name
                o("type") = e.TypeName
                o("updated_utc") = If(e.UpdatedUtc = DateTime.MinValue, "", e.UpdatedUtc.ToString("o"))
                o("seen_write") = e.SeenWrite
                o("seen_read") = e.SeenRead
                o("value") = e.Value
                ja.Add(o)
            Next
            Dim root As New JObject()
            root("snapshot_utc") = DateTime.UtcNow.ToString("o")
            root("count") = entries.Count
            root("vars") = ja
            File.WriteAllText(jsonPath, root.ToString(Formatting.Indented), Encoding.UTF8)

            If writeCsv Then
                Dim sbCsv As New StringBuilder()
                sbCsv.AppendLine("name,type,updated_utc,source,value")
                For Each e In entries
                    Dim src As String = If(e.SeenWrite AndAlso e.SeenRead, "Write+Read", If(e.SeenWrite, "Write", If(e.SeenRead, "Read", "")))
                    Dim row = String.Join(",", New String() {
                        CsvEsc(e.Name),
                        CsvEsc(e.TypeName),
                        CsvEsc(If(e.UpdatedUtc = DateTime.MinValue, "", e.UpdatedUtc.ToString("o"))),
                        CsvEsc(src),
                        CsvEsc(e.Value)
                    })
                    sbCsv.AppendLine(row)
                Next
                File.WriteAllText(csvPath, sbCsv.ToString(), Encoding.UTF8)
            End If

            VMSet("vm_audit_last_path_s", txtPath)
            VMSet("vm_log_checkbinds_s", "Audit written: " & txtPath)
            WriteLog("Audit snapshot written: " & txtPath)
        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "Audit error: " & ex.Message)
            WriteLog("Audit error: " & ex.ToString())
        End Try
    End Sub

    ' -----------------------------
    ' Prime bare → *_p mirrors once (with defaults & backfills)
    ' -----------------------------
    Private Sub PrimeCoreMirrors()
        Try
            ' 1) Backfill any *existing* values so *_p mirrors exist in this VM profile
            Dim core() As String = {
                "vm_hud_mode",
                "vm_firegroup",
                "vm_docked",
                "vm_in_supercruise",
                "vm_on_foot",
                "vm_environment",
                "vm_star_system",
                "vm_body",
                "vm_station_name",
                "vm_gui_focus",
                "vm_gui_focus_name",
                "vm_music_track",
                "vm_in_wing",
                "vm_wing_members",
                "vm_wing_leader",
                "vm_in_multicrew",
                "vm_multicrew_role",
                "vm_multicrew_captain",
                "vm_status_last_ts",
                "vm_journal_last_ts",
                "vm_active_journal",
                "vm_status_path",
                "vm_journal_dir",
                "vm_target_locked",   ' <-- add: ensure defined as 0/1
                "vm_target_name_s"    ' <-- add: ensure defined as a non-empty string
            }
            For Each n In core
                Dim v As String = VMGet(n)
                If Not String.IsNullOrEmpty(v) Then
                    VMSet(n, v)
                End If
            Next

            ' 2) Booleans/ints that must never be undefined (avoid ERR!)
            Dim defaultZeros() As String = {
                "vm_in_wing",
                "vm_wing_members",
                "vm_in_multicrew",
                "vm_docked",
                "vm_in_supercruise",
                "vm_on_foot",
                "vm_target_locked",
                "vm_startjump"' <-- add
            }
            For Each n In defaultZeros
                If String.IsNullOrEmpty(VMGet(n)) Then
                    VMSet(n, "0")
                End If
            Next

            ' 3) String defaults for things that otherwise show ERR! when unset
            '    (DO NOT use empty string: VM treats empty as delete → ERR!)
            If String.IsNullOrEmpty(VMGet("vm_target_name_s")) Then VMSet("vm_target_name_s", "NoTarget")
            If String.IsNullOrEmpty(VMGet("vm_wing_leader")) Then VMSet("vm_wing_leader", "Unknown")
            If String.IsNullOrEmpty(VMGet("vm_multicrew_role")) Then VMSet("vm_multicrew_role", "Unknown")
            If String.IsNullOrEmpty(VMGet("vm_multicrew_captain")) Then VMSet("vm_multicrew_captain", "Unknown")
            If String.IsNullOrEmpty(VMGet("vm_music_track")) Then VMSet("vm_music_track", "NoTrack")
            If String.IsNullOrEmpty(VMGet("vm_presence_oneshot_s")) Then VMSet("vm_presence_oneshot_s", "-")
            If String.IsNullOrEmpty(VMGet("vm_ai_last_action_s")) Then VMSet("vm_ai_last_action_s", "Idle")


            ' 4) Backfill friendly GUI focus name if only numeric exists
            Dim gf As String = VMGet("vm_gui_focus")
            Dim gfname As String = VMGet("vm_gui_focus_name")
            If Not String.IsNullOrEmpty(gf) AndAlso String.IsNullOrEmpty(gfname) Then
                Dim code As Integer
                If Integer.TryParse(gf, code) Then
                    Dim friendly As String
                    Select Case code
                        Case 0 : friendly = "NoFocus"
                        Case 1 : friendly = "InternalPanel"
                        Case 2 : friendly = "ExternalPanel"
                        Case 3 : friendly = "CommsPanel"
                        Case 4 : friendly = "RolePanel"
                        Case 5 : friendly = "StationServices"
                        Case 6 : friendly = "GalaxyMap"
                        Case 7 : friendly = "SystemMap"
                        Case 8 : friendly = "Orrery"
                        Case 9 : friendly = "FSS"
                        Case 10 : friendly = "SAAScanner"
                        Case 11 : friendly = "Codex"
                        Case Else : friendly = "Unknown"
                    End Select
                    VMSet("vm_gui_focus_name", friendly)
                End If
            End If
        Catch
        End Try

    End Sub

    ' ============================================================
    '        VMVarTypes generator (macro-usable *_p vars)
    ' ============================================================
    Private Sub WriteVmVarTypes()
        Try
            If String.IsNullOrEmpty(_baseDir) Then
                VMSet("vm_log_checkbinds_s", "VMVarTypes: base dir not set")
                Return
            End If

            Dim path As String = System.IO.Path.Combine(_baseDir, "VMVarTypes.txt")
            Dim sb As New StringBuilder()

            sb.AppendLine("VoiceMacro-Usable Variables")
            sb.AppendLine("Format: EVENT | USED IN VM MACROS | TYPE | WHAT IT DOES")
            sb.AppendLine()

            sb.AppendLine("=== Flight / State ===")
            sb.AppendLine("Docked | vm_docked_p | bool | 1 if docked; else 0")
            sb.AppendLine("Supercruise | vm_in_supercruise_p | bool | 1 if in supercruise; else 0")
            sb.AppendLine("OnFoot | vm_on_foot_p | bool | 1 if on foot; else 0")
            sb.AppendLine("Environment (derived) | vm_environment_p | text | Docked / Supercruise / OnFoot / Space")
            sb.AppendLine()

            sb.AppendLine("=== HUD / Firegroup ===")
            sb.AppendLine("HudMode | vm_hud_mode_p | text | Current HUD mode: Combat or Analysis")
            sb.AppendLine("FireGroup | vm_firegroup_p | int | Current firegroup index (0..N)")
            sb.AppendLine()

            sb.AppendLine("=== Docking / Station / Location ===")
            sb.AppendLine("Docked | vm_station_name_p | text | Station (when docked)")
            sb.AppendLine("Location/FSD/SC Exit | vm_star_system_p | text | Current star system")
            sb.AppendLine("Location/FSD/SC Exit | vm_body_p | text | Current body (best effort)")
            sb.AppendLine()

            sb.AppendLine("=== Supercruise / FSD ===")
            sb.AppendLine("SupercruiseEntry | vm_in_supercruise_p | bool | Set to 1 on entry")
            sb.AppendLine("SupercruiseExit | vm_in_supercruise_p | bool | Set to 0 on exit")
            sb.AppendLine("StartJump | vm_startjump_p | bool | 1 on hyperspace start, 0 on arrival (FSDJump)")
            sb.AppendLine()

            sb.AppendLine("=== GUI ===")
            sb.AppendLine("GuiFocus | vm_gui_focus_p | int | Current GUI focus code (numeric)")
            sb.AppendLine("GuiFocus (friendly name) | vm_gui_focus_name_p | text | Friendly name (e.g., LeftPanel)")
            sb.AppendLine()

            sb.AppendLine("=== Music ===")
            sb.AppendLine("Music | vm_music_track_p | text | Music tag (e.g., Combat, Exploration)")
            sb.AppendLine()

            sb.AppendLine("=== Wing ===")
            sb.AppendLine("WingJoin / WingLeave | vm_in_wing_p | bool | 1 in wing; 0 otherwise")
            sb.AppendLine("WingJoin / WingAdd | vm_wing_members_p | int | Wing member count (best effort)")
            sb.AppendLine("WingJoin | vm_wing_leader_p | text | Wing leader CMDR (if provided)")
            sb.AppendLine()

            sb.AppendLine("=== Multicrew ===")
            sb.AppendLine("JoinACrew / QuitACrew | vm_in_multicrew_p | bool | 1 when in multicrew; 0 otherwise")
            sb.AppendLine("CrewRoleChange | vm_multicrew_role_p | text | Multicrew role (e.g., Gunner)")
            sb.AppendLine("JoinACrew | vm_multicrew_captain_p | text | Captain CMDR (if provided)")
            sb.AppendLine()

            sb.AppendLine("=== Paths / Timestamps ===")
            sb.AppendLine("Status timestamp | vm_status_last_ts_p | text | Last Status.json time (local display)")
            sb.AppendLine("Any Journal line | vm_journal_last_ts_p | text | Last Journal time (local display)")
            sb.AppendLine("Init / Path setup | vm_status_path_p | text | Full path to Status.json")
            sb.AppendLine("Init / Path setup | vm_journal_dir_p | text | Journals folder path")
            sb.AppendLine("ColdLoad / Rotation | vm_active_journal_p | text | Current journal filename")
            sb.AppendLine()

            sb.AppendLine("=== Plugin / Health ===")
            sb.AppendLine("Init complete | vm_plugin_loaded_p | bool | 1 after plugin Init completes")
            sb.AppendLine("Heartbeat | vm_plugin_alive_p | iso8601 | Heartbeat timestamp (UTC ISO)")
            sb.AppendLine()

            sb.AppendLine("=== Remote / ExecuteMacro ===")
            sb.AppendLine("Remote base URL | vm_remote_baseurl_s | text | VM Remote base URL (ExecuteMacro endpoint)")
            sb.AppendLine("Remote debounce (ms) | vm_remote_debounce_ms | int | Debounce window for macro fires (ms)")
            sb.AppendLine()

            sb.AppendLine("=== Presence / AI / Tools (string helpers) ===")
            sb.AppendLine("selfteststatus | vm_self_test_result_s | text | Self-test summary (HUD/FG/System/OnFoot)")
            sb.AppendLine("vmvartypes | vm_last_types_path_s | text | Last VMVarTypes path")
            sb.AppendLine("auditstatus / auditcsv | vm_audit_last_path_s | text | Last StatusDump path")
            sb.AppendLine("auditstart / auditstop | vm_audit_state_s | text | Audit timer state")
            sb.AppendLine("checkbinds | vm_log_checkbinds_s | text | CheckBinds summary / instructions")
            sb.AppendLine("any plugin command | vm_dbg_last_cmd_s | text | Last command received by plugin")
            sb.AppendLine("rp_on / rp_off | vm_rp_state_s | text | RP toggle: on/off")
            sb.AppendLine("lfw_on / lfw_off | vm_presence_oneshot_s | text | One-shot presence line (consumed)")
            sb.AppendLine("aion / aioff / aitoggle | vm_ai_last_action_s | text | Last AI/Machina action")
            sb.AppendLine()

            System.IO.File.WriteAllText(path, sb.ToString(), System.Text.Encoding.UTF8)
            VMSet("vm_last_types_path_s", path)
            WriteLog("VMVarTypes written: " & path)
        Catch ex As Exception
            VMSet("vm_log_checkbinds_s", "VMVarTypes error: " & ex.Message)
            WriteLog("VMVarTypes error: " & ex.ToString())
        End Try
    End Sub

    ' --- Presence publisher: logs AND emits JSON snapshot for external helper ---
    Private Sub PublishPresence(text As String)
        Try
            ' Preserve current behavior: write to log
            WriteLog("[Presence] " & text)

            ' Split the DiscordPresence-composed text into up to two lines
            Dim line1 As String = ""
            Dim line2 As String = ""
            If Not String.IsNullOrEmpty(text) Then
                Dim parts = text.Replace(vbCrLf, vbLf).Split(New Char() {ChrW(10)}, 2, StringSplitOptions.None)
                line1 = parts(0).Trim()
                If parts.Length > 1 Then line2 = parts(1).Trim()
            End If

            ' Determine mode using existing VM vars (no new variables introduced)
            Dim inWing As Boolean = (VMGet("vm_in_wing_p") = "1")
            Dim oneshot As String = VMGet("vm_presence_oneshot_s")
            Dim hasOneShot As Boolean = Not String.IsNullOrWhiteSpace(oneshot) AndAlso
                            Not String.Equals(oneshot, "-", StringComparison.Ordinal)
            Dim mode As String

            If inWing Then
                mode = "wing"           ' detailed two-line card from DiscordPresence
            ElseIf hasOneShot Then
                mode = "lfw"            ' one-line LFW broadcast
                line1 = oneshot
                line2 = ""
            Else
                mode = "off"
                line1 = ""
                line2 = ""
            End If


            ' Emit JSON snapshot via MachinaBridge (atomic write)
            Dim payload As New System.Collections.Generic.Dictionary(Of String, String) From {
            {"kind", "presence"},
            {"mode", mode},
            {"line1", line1},
            {"line2", line2},
            {"timestamp", DateTime.UtcNow.ToString("o")}
        }

            If _machina IsNot Nothing Then
                _machina.PublishSnapshot(payload)
            End If

        Catch ex As Exception
            WriteLog("[Presence][ERR] " & ex.Message)
        End Try
    End Sub


    Private Shared Function GetVarGroup(name As String) As String
        Dim n = name.ToLowerInvariant()
        If n.Contains("star_system") OrElse n.Contains("station_name") OrElse n = "vm_body" OrElse n.Contains("body_") Then Return "Location"
        If n.Contains("docked") OrElse n.Contains("in_supercruise") OrElse n.Contains("landed") Then Return "Flight/State"
        If n.Contains("hud") OrElse n.Contains("firegroup") Then Return "HUD/Firegroup"
        If n.Contains("target_") Then Return "Targeting"
        If n.Contains("wing") OrElse n.Contains("presence") OrElse n.Contains("rp_") Then Return "Wing/Presence"
        If n.Contains("music") Then Return "Music"
        If n.Contains("plugin_") OrElse n.Contains("audit_") Then Return "Plugin/Health"
        If n.Contains("journal_") OrElse n.Contains("status_path") Then Return "Paths"
        Return "Misc"
    End Function

    Private Shared Function TypeRange(typeName As String) As String
        Select Case typeName
            Case "bool" : Return "0 or 1"
            Case "int" : Return "integer"
            Case "float" : Return "decimal"
            Case Else : Return "text"
        End Select
    End Function

    Private Function SanitizeOneLine(s As String) As String
        If s Is Nothing Then Return ""
        Dim t = s.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ")
        Return t
    End Function
    Private Function Truncate(s As String, maxLen As Integer) As String
        If s Is Nothing Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen - 1) & "…"
    End Function
    Private Function CsvEsc(s As String) As String
        If s Is Nothing Then s = ""
        Dim needs = s.Contains(",") OrElse s.Contains("""") OrElse s.Contains(vbCr) OrElse s.Contains(vbLf)
        Dim t = s.Replace("""", """""")
        If needs Then Return """" & t & """"
        Return t
    End Function

End Class
