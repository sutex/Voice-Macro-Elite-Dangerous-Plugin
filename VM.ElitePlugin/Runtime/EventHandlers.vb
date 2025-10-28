Option Explicit On
Option Strict On

Imports System
Imports System.Globalization
Imports Newtonsoft.Json.Linq

' EventHandlers: parses Status.json and Journal lines and writes bare vm_* via _vmSet.
' NOTE: Plugin.VMSet mirrors bare -> *_p automatically and records audit.
' POLICY: Never write empty strings for macro-visible fields (empty deletes in VM).
' Covered journals (subset): Location, FSDJump, SupercruiseEntry/Exit, Docked/Undocked,
' ApproachSettlement, Music, WingJoin/WingLeave/WingAdd, JoinACrew/QuitACrew/CrewRoleChange,
' Embark/Disembark, ShipTargeted.
Public Class EventHandlers
    Private ReadOnly _vmSet As Action(Of String, String)
    Private ReadOnly _vmGet As Func(Of String, String)
    Private ReadOnly _log As Action(Of String)

    Public Sub New(vmSet As Action(Of String, String),
                   vmGet As Func(Of String, String),
                   log As Action(Of String))
        _vmSet = vmSet
        _vmGet = vmGet
        _log = log
    End Sub

    ' ============================================================
    ' Status.json
    ' ============================================================
    Public Sub ProcessStatusJson(text As String)
        Try
            If String.IsNullOrWhiteSpace(text) Then Exit Sub

            Dim jo As JObject = Nothing
            Try
                jo = JObject.Parse(text)
            Catch ex As Exception
                _log("ProcessStatusJson: parse error: " & ex.Message)
                Exit Sub
            End Try
            If jo Is Nothing Then Exit Sub

            ' Timestamp (optional cosmetic)
            Dim tsTok = jo.SelectToken("timestamp")
            If tsTok IsNot Nothing Then
                Dim tsLocal = ToLocalTimestamp(tsTok.ToString())
                If tsLocal <> "" Then _vmSet("vm_status_last_ts", tsLocal)
            End If

            ' -----------------------
            ' HUD mode — correct key: HudInAnalysisMode (Boolean)
            ' -----------------------
            Dim hudTok = jo.SelectToken("HudInAnalysisMode")
            If hudTok IsNot Nothing Then
                Dim hudAnalysis As Boolean
                If Boolean.TryParse(hudTok.ToString(), hudAnalysis) Then
                    Dim hudModeTxt As String = If(hudAnalysis, "Analysis", "Combat")
                    _vmSet("vm_hud_mode", hudModeTxt)
                End If
            End If

            ' Fallback: if HudInAnalysisMode wasn't set, derive from Flags bit 27 (Hud in Analysis)
            If String.IsNullOrEmpty(_vmGet("vm_hud_mode")) Then
                Dim flagsVal2 As Long = 0
                Dim flagsTok2 = jo.SelectToken("Flags")
                If flagsTok2 IsNot Nothing Then Long.TryParse(flagsTok2.ToString(), flagsVal2)
                Dim hudAnalysis2 As Boolean = BitOn(flagsVal2, 27) ' uses your existing BitOn helper
                _vmSet("vm_hud_mode", If(hudAnalysis2, "Analysis", "Combat"))
            End If

            ' -----------------------
            ' Firegroup (correct casing: Firegroup)
            ' -----------------------
            Dim fgTok = jo.SelectToken("Firegroup")
            If fgTok IsNot Nothing Then
                Dim fgIdx As Integer
                If Integer.TryParse(fgTok.ToString(), fgIdx) Then
                    _vmSet("vm_firegroup", fgIdx.ToString(CultureInfo.InvariantCulture))
                End If
            End If

            ' -----------------------
            ' Flags (bitmask) — robust fallbacks
            ' -----------------------
            Dim flagsVal As Long = 0
            Dim flagsTok = jo.SelectToken("Flags")
            If flagsTok IsNot Nothing Then
                Long.TryParse(flagsTok.ToString(), flagsVal)
            End If

            ' Prefer explicit booleans when present, but OR with flags for robustness
            Dim docked As Boolean = GetBool(jo, "Docked") OrElse BitOn(flagsVal, 0)           ' bit 0
            Dim supercruise As Boolean = GetBool(jo, "Supercruise") OrElse BitOn(flagsVal, 4) ' bit 4

            ' OnFoot: prefer boolean; if missing, OR across Odyssey on-foot bits (29..33)
            Dim onFootProp As JToken = jo.SelectToken("OnFoot")
            Dim onFoot As Boolean
            If onFootProp IsNot Nothing Then
                Dim b As Boolean
                If Boolean.TryParse(onFootProp.ToString(), b) Then onFoot = b
            Else
                onFoot = (BitOn(flagsVal, 29) OrElse BitOn(flagsVal, 30) OrElse BitOn(flagsVal, 31) OrElse BitOn(flagsVal, 32) OrElse BitOn(flagsVal, 33))
            End If

            _vmSet("vm_docked", If(docked, "1", "0"))
            _vmSet("vm_in_supercruise", If(supercruise, "1", "0"))
            _vmSet("vm_on_foot", If(onFoot, "1", "0"))

            ' -----------------------
            ' GuiFocus (optional, numeric) + friendly name
            ' -----------------------
            Dim gfTok = jo.SelectToken("GuiFocus")
            If gfTok IsNot Nothing Then
                Dim gfVal As Integer
                If Integer.TryParse(gfTok.ToString(), gfVal) Then
                    _vmSet("vm_gui_focus", gfVal.ToString(CultureInfo.InvariantCulture))
                    _vmSet("vm_gui_focus_name", GuiFocusName(gfVal))
                End If
            End If

            ' -----------------------
            ' Environment hint (simple and stable)
            ' -----------------------
            WriteEnvironmentDerived(onFoot, supercruise, docked)

        Catch ex As Exception
            _log("ProcessStatusJson: error: " & ex.Message)
        End Try
    End Sub

    ' ============================================================
    ' Journal line
    ' ============================================================
    Public Sub ProcessJournalLine(line As String)
        Try
            If String.IsNullOrWhiteSpace(line) Then Exit Sub

            Dim jo As JObject = Nothing
            Try
                jo = JObject.Parse(line)
            Catch
                ' Some journal files include trailing junk/empty lines — ignore
                Exit Sub
            End Try
            If jo Is Nothing Then Exit Sub

            ' Remember last journal timestamp
            Dim tsTok = jo.SelectToken("timestamp")
            If tsTok IsNot Nothing Then
                Dim tsLocal = ToLocalTimestamp(tsTok.ToString())
                If tsLocal <> "" Then _vmSet("vm_journal_last_ts", tsLocal)
            End If

            Dim evtTok = jo.SelectToken("event")
            If evtTok Is Nothing Then Exit Sub
            Dim evt As String = evtTok.ToString()

            Select Case evt

                ' -------------------
                ' Location / session anchoring
                ' -------------------
                Case "Location"
                    WriteLocationLike(jo)
                _vmSet("vm_startjump", "0")  ' ensure cleared on session anchors


                    ' -------------------
                    ' Hyperspace / FSD
                    ' -------------------
                Case "FSDJump"
                    WriteLocationLike(jo)     ' StarSystem & Body if any
                    _vmSet("vm_in_supercruise", "0") ' after jump we are in normal space at arrival
                    _vmSet("vm_target_locked", "0")
                    _vmSet("vm_target_name_s", "NoTarget")
                    _vmSet("vm_startjump", "0")   ' clear level on arrival
                    RecomputeEnvironment()

                Case "StartJump"
                    Dim jt As String = jo.SelectToken("JumpType")?.ToString()
                    If String.Equals(jt, "Hyperspace", StringComparison.OrdinalIgnoreCase) Then
                        _log("Journal: StartJump (Hyperspace)")
                        ' Immediate pulse: clear any stale 1, then set 1 so VM always sees a change
                        _vmSet("vm_startjump", "0")
                        _vmSet("vm_startjump", "1")
                    Else
                        _log("Journal: StartJump (" & If(String.IsNullOrEmpty(jt), "?", jt) & ") — ignored")
                    End If

                Case "SupercruiseExit"
                    WriteLocationLike(jo)     ' Body and StarSystem present; not docked
                    _vmSet("vm_in_supercruise", "0")
                    _vmSet("vm_startjump", "0")   ' ensure cleared on SC exit
                    RecomputeEnvironment()


                Case "SupercruiseEntry"
                    _vmSet("vm_in_supercruise", "1")
                    _vmSet("vm_target_locked", "0")
                    _vmSet("vm_target_name_s", "NoTarget")
                    RecomputeEnvironment()


                ' -------------------
                ' On-foot edges
                ' -------------------
                Case "Embark"
                    ' Back in a seat/vehicle/ship -> no longer on foot
                    _vmSet("vm_on_foot", "0")
                    RecomputeEnvironment()

                Case "Disembark"
                    ' Player leaves a seat/vehicle/ship -> on foot
                    _vmSet("vm_on_foot", "1")
                    RecomputeEnvironment()

                ' -------------------
                ' Docking / station
                ' -------------------
                Case "Docked"
                    WriteLocationLike(jo)     ' StarSystem, StationName, Body
                    _vmSet("vm_docked", "1")
                    _vmSet("vm_target_locked", "0")
                    _vmSet("vm_target_name_s", "NoTarget")
                    RecomputeEnvironment()


                Case "Undocked"
                    _vmSet("vm_docked", "0")
                    RecomputeEnvironment()

                Case "ApproachSettlement"
                    ' Often includes BodyName and StarSystem
                    WriteLocationLike(jo)

                ' -------------------
                ' Music (short-lived hint)
                ' -------------------
                Case "Music"
                    Dim mt = jo.SelectToken("MusicTrack")?.ToString()
                    If Not String.IsNullOrEmpty(mt) Then
                        _vmSet("vm_music_track", mt)
                    End If
                    ' do NOT write empty if missing


                 ' -------------------
                 ' Targeting (for Discord line 2: "Engaging …")
                 ' -------------------

                Case "ShipTargeted"
                    ' Journal provides TargetLocked plus optional pilot/ship strings.
                    Dim lockedTok = jo.SelectToken("TargetLocked")
                    Dim locked As Boolean = False
                    If lockedTok IsNot Nothing Then Boolean.TryParse(lockedTok.ToString(), locked)

                    If locked Then
                        _vmSet("vm_target_locked", "1")

                        ' Prefer Localised → fall back to non-localised → ship
                        Dim name As String = jo.SelectToken("PilotName_Localised")?.ToString()
                        If String.IsNullOrEmpty(name) Then name = jo.SelectToken("PilotName")?.ToString()
                        If String.IsNullOrEmpty(name) Then name = jo.SelectToken("Ship_Localised")?.ToString()
                        If String.IsNullOrEmpty(name) Then name = jo.SelectToken("Ship")?.ToString()

                        ' Sanitize tokens like $npc_name_decorate:#name=GutBuster;
                        If Not String.IsNullOrEmpty(name) AndAlso name.StartsWith("$") AndAlso name.EndsWith(";") Then
                            Dim hash = name.IndexOf("#name=", StringComparison.OrdinalIgnoreCase)
                            If hash >= 0 Then
                                Dim val = name.Substring(hash + 6)
                                If val.EndsWith(";") Then val = val.Substring(0, val.Length - 1)
                                name = val
                            Else
                                name = name.Trim("$"c, ";"c).Replace("_", " ")
                            End If
                        End If

                        ' Normalize placeholders/empty to a safe non-empty value (prevents ERR!)
                        If String.IsNullOrWhiteSpace(name) _
                           OrElse name.Equals("ERR!", StringComparison.OrdinalIgnoreCase) _
                           OrElse name.Equals("N/A", StringComparison.OrdinalIgnoreCase) _
                           OrElse name.Equals("NULL", StringComparison.OrdinalIgnoreCase) Then
                            name = "Unknown"
                        End If

                        ' Length guard (keep presence/overlays tidy)
                        If name.Length > 60 Then name = name.Substring(0, 59) & "…"

                        ' Change-only update (reduces churn during scan stages)
                        Dim cur As String = _vmGet("vm_target_name_s")
                        If Not String.Equals(cur, name, StringComparison.Ordinal) Then
                            _vmSet("vm_target_name_s", name)
                        End If
                    Else
                        ' target cleared / lost
                        _vmSet("vm_target_locked", "0")
                        ' IMPORTANT: never write empty here; use a safe placeholder
                        _vmSet("vm_target_name_s", "NoTarget")
                    End If



                ' -------------------
                ' Wing basics
                ' -------------------
                Case "WingJoin"
                    _vmSet("vm_in_wing", "1")
                    _vmSet("vm_presence_oneshot_s", "")
                    Dim m = jo.SelectToken("Others")
                    If m IsNot Nothing AndAlso m.Type = JTokenType.Array Then
                        Dim cnt As Integer = CType(m, JArray).Count + 1 ' +1 includes you
                        If cnt > 4 Then cnt = 4
                        _vmSet("vm_wing_members", cnt.ToString(CultureInfo.InvariantCulture))
                    Else
                        _vmSet("vm_wing_members", "1") ' at least yourself
                    End If

                Case "WingLeave"
                    _vmSet("vm_in_wing", "0")
                    _vmSet("vm_wing_members", "0")
                    ' DO NOT clear with empty string (VM treats empty as delete)
                    ' _vmSet("vm_wing_leader", "")

                Case "WingAdd"
                    ' Someone added — adjust members best-effort (cap at 4)
                    Dim cur = ParseIntSafe(_vmGet("vm_wing_members"))
                    If cur <= 0 Then cur = 1
                    Dim nxt As Integer = cur + 1
                    If nxt > 4 Then nxt = 4
                    _vmSet("vm_wing_members", nxt.ToString(CultureInfo.InvariantCulture))

                ' -------------------
                ' Multicrew basics
                ' -------------------
                Case "JoinACrew"
                    _vmSet("vm_in_multicrew", "1")
                    Dim cap = jo.SelectToken("Captain")?.ToString()
                    If Not String.IsNullOrEmpty(cap) Then _vmSet("vm_multicrew_captain", cap)

                Case "QuitACrew"
                    _vmSet("vm_in_multicrew", "0")
                    ' DO NOT clear with empty string (VM treats empty as delete)
                    ' _vmSet("vm_multicrew_role", "")
                    ' _vmSet("vm_multicrew_captain", "")

                Case "CrewRoleChange"
                    Dim role = jo.SelectToken("Role")?.ToString()
                    If Not String.IsNullOrEmpty(role) Then _vmSet("vm_multicrew_role", role)

                Case Else
                    ' no-op
            End Select

        Catch ex As Exception
            _log("ProcessJournalLine: error: " & ex.Message)
        End Try
    End Sub

    ' ------------------------------------------------------------
    ' Helpers
    ' ------------------------------------------------------------
    Private Sub WriteLocationLike(jo As JObject)
        Try
            Dim sys = jo.SelectToken("StarSystem")?.ToString()
            If Not String.IsNullOrEmpty(sys) Then _vmSet("vm_star_system", sys)

            ' Journal sometimes uses "Body" or "BodyName"
            Dim body = jo.SelectToken("Body")?.ToString()
            If String.IsNullOrEmpty(body) Then body = jo.SelectToken("BodyName")?.ToString()
            If Not String.IsNullOrEmpty(body) Then _vmSet("vm_body", body)

            Dim stName = jo.SelectToken("StationName")?.ToString()
            If Not String.IsNullOrEmpty(stName) Then _vmSet("vm_station_name", stName)

        Catch ex As Exception
            _log("WriteLocationLike: " & ex.Message)
        End Try
    End Sub

    Private Sub RecomputeEnvironment()
        Try
            Dim docked As Boolean = (_vmGet("vm_docked") = "1")
            Dim sc As Boolean = (_vmGet("vm_in_supercruise") = "1")
            Dim onFoot As Boolean = (_vmGet("vm_on_foot") = "1")
            WriteEnvironmentDerived(onFoot, sc, docked)
        Catch
        End Try
    End Sub

    Private Sub WriteEnvironmentDerived(onFoot As Boolean, supercruise As Boolean, docked As Boolean)
        Dim env As String
        If onFoot Then
            env = "OnFoot"
        ElseIf supercruise Then
            env = "Supercruise"
        ElseIf docked Then
            env = "Docked"
        Else
            env = "Space"
        End If
        _vmSet("vm_environment", env)
    End Sub

    Private Shared Function GuiFocusName(code As Integer) As String
        ' Conservative mapping: values as observed in Frontier docs/community lists.
        Select Case code
            Case 0 : Return "NoFocus"
            Case 1 : Return "InternalPanel"   ' right
            Case 2 : Return "ExternalPanel"   ' left
            Case 3 : Return "CommsPanel"
            Case 4 : Return "RolePanel"
            Case 5 : Return "StationServices"
            Case 6 : Return "GalaxyMap"
            Case 7 : Return "SystemMap"
            Case 8 : Return "Orrery"
            Case 9 : Return "FSS"
            Case 10 : Return "SAAScanner"
            Case 11 : Return "Codex"
            Case Else : Return "Unknown"
        End Select
    End Function

    Private Shared Function GetBool(jo As JObject, prop As String) As Boolean
        Dim t = jo.SelectToken(prop)
        If t Is Nothing Then Return False
        Dim b As Boolean
        If Boolean.TryParse(t.ToString(), b) Then Return b
        Return False
    End Function

    Private Shared Function BitOn(mask As Long, bit As Integer) As Boolean
        If bit < 0 OrElse bit > 62 Then Return False
        Return (mask And (1L << bit)) <> 0
    End Function

    Private Shared Function ToLocalTimestamp(iso As String) As String
        Try
            Dim dt As DateTime
            If DateTime.TryParse(iso, Nothing, DateTimeStyles.AdjustToUniversal Or DateTimeStyles.AssumeUniversal, dt) Then
                Return dt.ToLocalTime().ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture)
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function ParseIntSafe(s As String) As Integer
        Dim v As Integer
        If Integer.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, v) Then Return v
        Return 0
    End Function

End Class

