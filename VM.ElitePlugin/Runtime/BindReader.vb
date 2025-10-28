Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq

Public Module BindReader

    ' --- watcher state (module scope) ---
    Private ReadOnly _sync As New Object()
    Private _watcher As IO.FileSystemWatcher = Nothing
    Private _debounce As System.Timers.Timer = Nothing
    Private _log As Action(Of String) = Nothing
    Private Const BINDS_DEBOUNCE_MS As Double = 750.0
    Private _bootstrapped As Boolean = False

    Public Function ScanAndPublish(logger As Action(Of String)) As Boolean
        Try
            Dim bindsPath As String = FindActiveBindsFile()
            If String.IsNullOrEmpty(bindsPath) OrElse Not File.Exists(bindsPath) Then
                SafeSet("ed_status", "not_found", logger)
                SafeSet("ed_bindspath", "", logger)
                SafeSet("ed_error", "binds file not found", logger)
                SafeSet("ed_report", "No binds file found.", logger)
                Return False
            End If

            Dim snap As Snapshot = ScanFile(bindsPath, logger)

            SafeSet("ed_bindspath", bindsPath, logger)
            SafeSet("ed_status", If(snap.ParseError Is Nothing, "ok", "parse_error"), logger)
            SafeSet("ed_error", If(snap.ParseError, ""), logger)

            For Each kvp In snap.Actions
                Dim exported As String = SelectExportChord(kvp.Value)
                SafeSet("ed_" & kvp.Key, exported, logger)
            Next

            Dim reportText As String = BuildReport(snap)
            SafeSet("ed_report", reportText, logger)

            ' ---- write readable report file + surface path in vm_log_checkbinds_s ----
            Dim reportPath As String = WriteBindsReport(snap, bindsPath, logger)
            If reportPath.Length > 0 Then
                SafeSet("vm_log_checkbinds_s", "Binds written: " & reportPath, logger)
            End If
            ' -----------------------------------------------------------------------------

            Try
                Dim backupPath As String = WriteBackup(bindsPath, snap.PresetName, logger)
                If backupPath.Length > 0 Then
                    SafeSet("ed_backup_last", backupPath, logger)
                    SafeSet("ed_backup_error", "", logger)
                    RotateBackups(snap.PresetName, 2, logger)
                End If
            Catch ex As Exception
                Log(logger, "[BindReader] backup_error: " & Trunc(ex.Message, 240))
                SafeSet("ed_backup_error", Trunc(ex.Message, 240), logger)
            End Try

            ' start watcher after first successful publish
            EnsureWatcherAfterFirstScan(logger)

            Return snap.ParseError Is Nothing

        Catch ex As Exception
            Log(logger, "[BindReader] fatal: " & Trunc(ex.Message, 240))
            Try
                SafeSet("ed_status", "parse_error", logger)
                SafeSet("ed_error", Trunc(ex.Message, 240), logger)
            Catch
            End Try
            Return False
        End Try
    End Function

    Friend Class Snapshot
        Public Property PresetName As String = ""
        Public Property ParseError As String = Nothing
        Public Property Actions As New Dictionary(Of String, ActionEntry)(StringComparer.OrdinalIgnoreCase)
    End Class

    Friend Class ActionEntry
        Public Property Name As String = ""
        Public Property Group As String = "misc"
        Public Property PrimaryIsKeyboard As Boolean = False
        Public Property PrimaryChord As String = ""
        Public Property PrimaryVJoyIndex As Integer = -1
        Public Property PrimaryKeyToken As String = ""
        Public Property SecondaryIsKeyboard As Boolean = False
        Public Property SecondaryChord As String = ""
        Public Property SecondaryVJoyIndex As Integer = -1
        Public Property SecondaryKeyToken As String = ""

        ' derived at print time
        Public ReadOnly Property SubGroup As String
            Get
                Return MapSubGroup(Name, Group)
            End Get
        End Property
    End Class

    Private Function ScanFile(path As String, logger As Action(Of String)) As Snapshot
        Dim snap As New Snapshot()
        Try
            Dim doc As XDocument = XDocument.Load(path)
            Dim root As XElement = doc.Root
            If root Is Nothing Then Throw New InvalidDataException("missing Root")

            Dim presetAttr As XAttribute = root.Attribute("PresetName")
            snap.PresetName = If(presetAttr IsNot Nothing, presetAttr.Value, "")

            For Each node As XElement In root.Elements()
                Dim actionName As String = node.Name.LocalName
                If Not IsActionNode(actionName) Then Continue For

                Dim entry As New ActionEntry() With {
                    .Name = actionName,
                    .Group = MapGroup(actionName)
                }

                Dim pri As XElement = GetChild(node, "Primary")
                If pri IsNot Nothing Then FillSlot(pri, True, entry)

                Dim sec As XElement = GetChild(node, "Secondary")
                If sec IsNot Nothing Then FillSlot(sec, False, entry)

                Dim bindEl As XElement = GetChild(node, "Binding")
                If bindEl IsNot Nothing Then
                    Dim dev As String = ReadAttr(bindEl, "Device")
                    Dim idx As Integer = ReadAttrInt(bindEl, "DeviceIndex", -1)
                    Dim keyTok As String = ReadAttr(bindEl, "Key")
                    If String.Equals(dev, "vJoy", StringComparison.OrdinalIgnoreCase) Then
                        If entry.PrimaryVJoyIndex < 0 AndAlso entry.PrimaryKeyToken = "" Then
                            entry.PrimaryVJoyIndex = idx
                            entry.PrimaryKeyToken = keyTok
                        End If
                    End If
                End If

                snap.Actions(entry.Name) = entry
            Next

        Catch ex As Exception
            Log(logger, "[BindReader] parse_error: " & Trunc(ex.Message, 240))
            snap.ParseError = Trunc(ex.Message, 240)
        End Try

        Return snap
    End Function

    Private Sub FillSlot(slot As XElement, isPrimary As Boolean, entry As ActionEntry)
        Dim dev As String = ReadAttr(slot, "Device")
        Dim key As String = ReadAttr(slot, "Key")

        Dim isKbd As Boolean = String.Equals(dev, "Keyboard", StringComparison.OrdinalIgnoreCase) AndAlso key.Length > 0
        If isKbd Then
            Dim chord As String = BuildKeyboardChord(slot, key)
            If isPrimary Then
                entry.PrimaryIsKeyboard = True
                entry.PrimaryChord = chord
            Else
                entry.SecondaryIsKeyboard = True
                entry.SecondaryChord = chord
            End If
            Return
        End If

        If String.Equals(dev, "vJoy", StringComparison.OrdinalIgnoreCase) Then
            Dim vjoyIdx As Integer = ReadAttrInt(slot, "DeviceIndex", -1)
            If isPrimary Then
                If entry.PrimaryVJoyIndex < 0 AndAlso vjoyIdx >= 0 Then entry.PrimaryVJoyIndex = vjoyIdx
                If entry.PrimaryKeyToken = "" AndAlso key.Length > 0 Then entry.PrimaryKeyToken = key
            Else
                If entry.SecondaryVJoyIndex < 0 AndAlso vjoyIdx >= 0 Then entry.SecondaryVJoyIndex = vjoyIdx
                If entry.SecondaryKeyToken = "" AndAlso key.Length > 0 Then entry.SecondaryKeyToken = key
            End If
        End If
    End Sub

    Private Function BuildKeyboardChord(slot As XElement, keyToken As String) As String
        Dim baseKey As String = NormalizeKey(keyToken)
        Dim mods As New List(Of String)
        For Each m As XElement In slot.Elements()
            If String.Equals(m.Name.LocalName, "Modifier", StringComparison.OrdinalIgnoreCase) Then
                Dim mDev As String = ReadAttr(m, "Device")
                Dim mKey As String = ReadAttr(m, "Key")
                If String.Equals(mDev, "Keyboard", StringComparison.OrdinalIgnoreCase) AndAlso mKey.Length > 0 Then
                    mods.Add(NormalizeKey(mKey))
                End If
            End If
        Next
        Dim ordered = mods.OrderBy(Function(s) ModPriority(s)).ThenBy(Function(s) s, StringComparer.Ordinal)
        Dim sb As New StringBuilder()
        For Each modKey In ordered
            If sb.Length > 0 Then sb.Append("+"c)
            sb.Append(modKey)
        Next
        If sb.Length > 0 Then sb.Append("+"c)
        sb.Append(baseKey)
        Return sb.ToString()
    End Function

    Private Function NormalizeKey(keyToken As String) As String
        If keyToken Is Nothing Then Return ""
        Dim t As String = keyToken
        If t.StartsWith("Key_", StringComparison.OrdinalIgnoreCase) Then
            t = t.Substring(4)
        End If

        ' Consistent Control naming: Control -> Ctrl (preserve Left/Right first)
        If t.IndexOf("Control", StringComparison.OrdinalIgnoreCase) >= 0 Then
            t = ReplaceIgnoreCase(t, "LeftControl", "LeftCtrl")
            t = ReplaceIgnoreCase(t, "RightControl", "RightCtrl")
            t = ReplaceIgnoreCase(t, "Control", "Ctrl")
        End If

        Return t
    End Function

    Private Function ReplaceIgnoreCase(input As String, search As String, replacement As String) As String
        If String.IsNullOrEmpty(input) OrElse String.IsNullOrEmpty(search) Then Return input
        Dim sb As New StringBuilder(input.Length)
        Dim i As Integer = 0
        While True
            Dim idx As Integer = input.IndexOf(search, i, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then
                sb.Append(input, i, input.Length - i)
                Exit While
            End If
            sb.Append(input, i, idx - i)
            sb.Append(replacement)
            i = idx + search.Length
        End While
        Return sb.ToString()
    End Function

    Private Function ModPriority(keyName As String) As Integer
        Dim k As String = keyName.ToLowerInvariant()
        If k.Contains("control") OrElse k.Contains("ctrl") Then Return 0
        If k.Contains("alt") Then Return 1
        If k.Contains("shift") Then Return 2
        Return 3
    End Function

    Private Function SelectExportChord(a As ActionEntry) As String
        If a.SecondaryIsKeyboard AndAlso a.SecondaryChord.Length > 0 Then Return a.SecondaryChord
        If a.PrimaryIsKeyboard AndAlso a.PrimaryChord.Length > 0 Then Return a.PrimaryChord
        Return ""
    End Function

    ' ------------------
    ' FORMATTED REPORT
    ' ------------------
    Private Function BuildReport(snap As Snapshot) As String
        Dim sb As New StringBuilder(8192)

        ' Header + legend for empty quotes
        sb.AppendLine("Elite Dangerous — Keybinds Snapshot")
        sb.AppendLine("Note: """" means no keybinding found.")
        sb.AppendLine()

        ' Precompute column width for aligned "ed_<name> =   value"
        Dim maxVarLen As Integer = 0
        For Each a In snap.Actions.Values
            Dim varNameLen = ("ed_" & a.Name).Length
            If varNameLen > maxVarLen Then maxVarLen = varNameLen
        Next
        Dim pad As Integer = Math.Max(28, maxVarLen + 2) ' a little extra breathing room
        Dim afterEq As String = " =   " ' equals then THREE spaces before value

        ' Build grouped listing order (top level)
        Dim groups = snap.Actions.Values _
            .GroupBy(Function(x) x.Group, StringComparer.OrdinalIgnoreCase) _
            .OrderBy(Function(g) g.Key, StringComparer.OrdinalIgnoreCase)

        ' 1) VARIABLES (grouped, easy to copy)
        sb.AppendLine("VARIABLES")
        sb.AppendLine("---------")

        For Each g In groups
            sb.AppendLine()
            sb.AppendLine("[" & g.Key & "]")
            sb.AppendLine(New String("-"c, Math.Max(6, g.Key.Length + 2)))

            Dim block As IEnumerable(Of ActionEntry) = g

            ' If misc: add Elite-style sub-groups
            If String.Equals(g.Key, "misc", StringComparison.OrdinalIgnoreCase) Then
                Dim subGroups = g.GroupBy(Function(x) x.SubGroup, StringComparer.OrdinalIgnoreCase) _
                                 .OrderBy(Function(sg) SubGroupSortOrder(sg.Key)) _
                                 .ThenBy(Function(sg) sg.Key, StringComparer.OrdinalIgnoreCase)

                For Each sg In subGroups
                    ' subgroup header and a clean separator
                    sb.AppendLine()
                    sb.AppendLine("  [" & sg.Key & "]")
                    sb.AppendLine("  " & New String("-"c, Math.Max(6, sg.Key.Length + 2)))

                    For Each a In sg.OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase)
                        Dim exportVal As String = SelectExportChord(a)
                        Dim varName As String = "ed_" & a.Name
                        Dim valueOut As String = If(exportVal.Length = 0, """" & """"c, exportVal)
                        sb.Append("  ")
                        sb.Append(varName.PadRight(pad))
                        sb.Append(afterEq)
                        sb.AppendLine(valueOut)
                    Next
                Next

            Else
                ' non-misc: flat within the group
                For Each a In block.OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase)
                    Dim exportVal As String = SelectExportChord(a)
                    Dim varName As String = "ed_" & a.Name
                    Dim valueOut As String = If(exportVal.Length = 0, """" & """"c, exportVal)
                    sb.Append(varName.PadRight(pad)).Append(afterEq).AppendLine(valueOut)
                Next
            End If
        Next

        sb.AppendLine()
        sb.AppendLine(New String("="c, 64))
        sb.AppendLine()

        ' 2) DETAIL (grouped; misc also broken into sub-groups)
        sb.AppendLine("DETAIL")
        sb.AppendLine("------")
        For Each g In groups
            sb.AppendLine()
            sb.AppendLine("[Group: " & g.Key & "]")
            sb.AppendLine()

            Dim detailSections As IEnumerable(Of IGrouping(Of String, ActionEntry))
            If String.Equals(g.Key, "misc", StringComparison.OrdinalIgnoreCase) Then
                detailSections = g.GroupBy(Function(x) x.SubGroup, StringComparer.OrdinalIgnoreCase) _
                                  .OrderBy(Function(sg) SubGroupSortOrder(sg.Key)) _
                                  .ThenBy(Function(sg) sg.Key, StringComparer.OrdinalIgnoreCase)
            Else
                ' Single pseudo-section to reuse rendering below
                detailSections = {g.GroupBy(Function(x) "", StringComparer.OrdinalIgnoreCase).First()}
            End If

            For Each sect In detailSections
                If sect.Key IsNot Nothing AndAlso sect.Key <> "" Then
                    sb.AppendLine("[" & sect.Key & "]")
                    sb.AppendLine(New String("-"c, Math.Max(6, sect.Key.Length)))
                End If

                For Each a In sect.OrderBy(Function(x) x.Name, StringComparer.OrdinalIgnoreCase)
                    Dim exportVal As String = SelectExportChord(a)
                    Dim used As String = "None"
                    If a.SecondaryIsKeyboard AndAlso exportVal = a.SecondaryChord Then used = "Secondary"
                    If a.PrimaryIsKeyboard AndAlso exportVal = a.PrimaryChord Then used = "Primary"

                    sb.AppendLine(a.Name & " :")

                    If a.PrimaryIsKeyboard Then
                        sb.AppendLine("  Primary   : " & a.PrimaryChord)
                    ElseIf a.PrimaryVJoyIndex >= 0 OrElse a.PrimaryKeyToken.Length > 0 Then
                        sb.AppendLine(String.Format(CultureInfo.InvariantCulture, "  Primary   : VJoy_Index={0}, Key={1}", a.PrimaryVJoyIndex, a.PrimaryKeyToken))
                    Else
                        sb.AppendLine("  Primary   : (none)")
                    End If

                    If a.SecondaryIsKeyboard Then
                        sb.AppendLine("  Secondary : " & a.SecondaryChord)
                    ElseIf a.SecondaryVJoyIndex >= 0 OrElse a.SecondaryKeyToken.Length > 0 Then
                        sb.AppendLine(String.Format(CultureInfo.InvariantCulture, "  Secondary : VJoy_Index={0}, Key={1}", a.SecondaryVJoyIndex, a.SecondaryKeyToken))
                    Else
                        sb.AppendLine("  Secondary : (none)")
                    End If

                    Dim varName As String = "ed_" & a.Name
                    Dim valueOut As String = If(exportVal.Length = 0, """" & """"c, exportVal)
                    sb.Append("  ").Append(varName.PadRight(pad)).Append(afterEq).Append(valueOut)
                    sb.Append("   (Used: ").Append(used).AppendLine(")")
                    sb.AppendLine()
                Next
            Next
        Next

        If snap.ParseError IsNot Nothing Then
            Dim warn As String = "[parse_error] " & snap.ParseError
            Return warn & Environment.NewLine & Environment.NewLine & sb.ToString() & Environment.NewLine & warn & Environment.NewLine
        End If

        Return sb.ToString()
    End Function

    ' ------------------
    ' GROUPING
    ' ------------------
    Private Function MapGroup(actionName As String) As String
        Dim n As String = actionName.ToLowerInvariant()
        If n.EndsWith("_landing", StringComparison.Ordinal) Then Return "landing"
        If n.EndsWith("_buggy", StringComparison.Ordinal) Then Return "buggy"
        If n.EndsWith("_humanoid", StringComparison.Ordinal) Then Return "humanoid"
        If n.StartsWith("explorationfss", StringComparison.Ordinal) OrElse n.StartsWith("explorationsaa", StringComparison.Ordinal) Then Return "fss"
        If n.StartsWith("cam", StringComparison.Ordinal) OrElse n.Contains("camera") Then Return "camera"
        Return "misc"
    End Function

    ' Map sub-groups for "misc" based on Elite categories provided by you.
    ' Names are NOT changed — just classified.
    Private Function MapSubGroup(actionName As String, topGroup As String) As String
        If Not String.Equals(topGroup, "misc", StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If

        Dim n As String = actionName.ToLowerInvariant()

        ' 1) Flight Controls
        If n.Contains("pitch") OrElse n.Contains("yaw") OrElse n.Contains("roll") _
           OrElse n.Contains("thrust") OrElse n.Contains("steer") _
           OrElse n.Contains("lateral") OrElse n.Contains("vertical") _
           OrElse n = "throttleaxis" OrElse n.Contains("setspeed") _
           OrElse n = "forwardkey" OrElse n = "backwardkey" Then
            Return "Flight Controls"
        End If

        ' 2) Power Management
        If n.Contains("increaseenginespower") OrElse n.Contains("increasesystemspower") _
           OrElse n.Contains("increaseweaponspower") OrElse n.Contains("resetpowerdistribution") Then
            Return "Power Management"
        End If

        ' 3) Systems (fa/supercruise/hyperspace etc)
        If n.Contains("toggleflightassist") OrElse n.Contains("supercruise") _
           OrElse n.Contains("hyperspace") OrElse n.Contains("hypersupercombination") _
           OrElse n.Contains("toggleadvance") OrElse n.Contains("uifocusmode") Then
            Return "Systems"
        End If

        ' 4) Targeting
        If n.Contains("target") OrElse n.Contains("wingman") _
           OrElse n.Contains("selecthighestthreat") Then
            Return "Targeting"
        End If

        ' 5) Weapons
        If n.Contains("primaryfire") OrElse n.Contains("secondaryfire") _
           OrElse n.Contains("firegroup") OrElse n.Contains("deployhardpoint") _
           OrElse n.Contains("deployheatsink") OrElse n.Contains("usechaff") _
           OrElse n.Contains("chargeecm") OrElse n.Contains("shipspotlighttoggle") Then
            Return "Weapons"
        End If

        ' 6) Scanner & Detection
        If n.Contains("fssmouse") OrElse n.Contains("fstop") _
           OrElse n.Contains("radar") OrElse n.Contains("discovery") Then
            Return "Scanner & Detection"
        End If

        ' 7) Utility
        If n.Contains("landinggear") OrElse n.Contains("cargoscoop") _
           OrElse n.Contains("nightvisiontoggle") OrElse n.Contains("useshieldcell") _
           OrElse n.Contains("headlightsbuggybutton") OrElse n.Contains("buggy") _
           OrElse n.Contains("microphonemute") OrElse n.Contains("rotatesettlement") _
           OrElse n.Contains("trigg") Then
            Return "Utility"
        End If

        ' 8) Navigation
        If n.Contains("galaxymap") OrElse n.Contains("systemmap") _
           OrElse n.Contains("orbitlines") OrElse n.Contains("targetnextroutesystem") _
           OrElse n.Contains("recalldismissship") Then
            Return "Navigation"
        End If

        ' 9) UI & Camera
        If n.StartsWith("ui_") OrElse n = "uifocus" OrElse n.Contains("uifocus") _
           OrElse n.Contains("freecam") OrElse n.Contains("headlook") _
           OrElse n = "playerhudmodetoggle" OrElse n = "pause" _
           OrElse n = "hmdreset" OrElse n.Contains("storecam") Then
            Return "UI & Camera"
        End If

        ' 10) Multi-Crew
        If n.StartsWith("multicrew") OrElse n.Contains("rolepanelfocusoptions") Then
            Return "Multi-Crew"
        End If

        ' Fallback so nothing stays ungrouped
        Return "Utility"
    End Function

    ' Stablizes order of sub-groups when printing
    Private Function SubGroupSortOrder(name As String) As Integer
        Select Case name
            Case "Flight Controls" : Return 1
            Case "Power Management" : Return 2
            Case "Systems" : Return 3
            Case "Targeting" : Return 4
            Case "Weapons" : Return 5
            Case "Scanner & Detection" : Return 6
            Case "Utility" : Return 7
            Case "Navigation" : Return 8
            Case "UI & Camera" : Return 9
            Case "Multi-Crew" : Return 10
            Case Else : Return 99
        End Select
    End Function

    ' ------------------
    ' FILTERS / HELPERS
    ' ------------------
    Private Function IsActionNode(localName As String) As Boolean
        If String.IsNullOrEmpty(localName) Then Return False
        Select Case localName
            Case "KeyboardLayout", "MouseXMode", "MouseXDecay", "MouseYMode", "MouseYDecay",
                 "MouseSensitivity", "MouseDecayRate", "MouseDeadzone", "MouseLinearity",
                 "MouseGUI", "BlockMouseDecay", "EnableMenuGroups", "EnableMenuGroupsOnFoot",
                 "EnableAimAssistOnFoot", "EnableCameraLockOn", "DriveAssistDefault",
                 "ThrottleIncrement", "BuggyThrottleIncrement", "ThrottleRange", "ThrottleRangeFreeCam",
                 "YawToRollMode", "YawToRollSensitivity", "YawToRollMode_FAOff", "YawToRollMode_Landing",
                 "MuteButtonMode", "CqcMuteButtonMode", "WeaponColourToggle", "EngineColourToggle",
                 "UseAlternateFlightValuesToggle", "ToggleButtonUpInput",
                 "HeadlookDefault", "HeadlookIncrement", "HeadlookMode", "HeadlookResetOnToggle",
                 "HeadlookSensitivity", "HeadlookSmoothing", "MouseHeadlook", "MouseHeadlookInvert",
                 "MouseHeadlookSensitivity", "HeadlookMotionSensitivity", "yawRotateHeadlook",
                 "FSSTuningSensitivity", "SAAThirdPersonMouseSensitivity",
                 "FreeCamMouseSensitivity", "FreeCamMouseYDecay", "FreeCamMouseXDecay",
                 "PlacementCamMouseSensitivity", "PlacementCamMouseXDecay", "PlacementCamMouseYDecay",
                 "EnableRumbleTrigger", "EnableMenuGroupsSRV", "ShowPGScoreSummaryInput"
                Return False
        End Select
        Return True
    End Function

    Private Function GetChild(parent As XElement, localName As String) As XElement
        For Each c As XElement In parent.Elements()
            If String.Equals(c.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase) Then
                Return c
            End If
        Next
        Return Nothing
    End Function

    ' ------------------
    ' BACKUPS
    ' ------------------
    Private Function WriteBackup(bindsPath As String, presetName As String, logger As Action(Of String)) As String
        Try
            Dim logDir As String = GetPluginLogDir()
            If String.IsNullOrEmpty(logDir) Then Return ""
            If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)

            Dim stamp As String = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) & "Z"
            Dim safePreset As String = If(String.IsNullOrEmpty(presetName), "UnknownPreset", SanitizeFileName(presetName))
            Dim fileName As String = safePreset & "_" & stamp & ".binds"
            Dim dest As String = Path.Combine(logDir, fileName)

            File.Copy(bindsPath, dest, overwrite:=False)
            Log(logger, "[BindReader] backup -> " & dest)
            Return dest
        Catch ex As Exception
            Log(logger, "[BindReader] backup failed: " & Trunc(ex.Message, 240))
            Return ""
        End Try
    End Function

    Private Sub RotateBackups(presetName As String, maxKeep As Integer, logger As Action(Of String))
        Try
            Dim logDir As String = GetPluginLogDir()
            If String.IsNullOrEmpty(logDir) OrElse Not Directory.Exists(logDir) Then Return
            Dim safePreset As String = If(String.IsNullOrEmpty(presetName), "UnknownPreset", SanitizeFileName(presetName))

            Dim files As List(Of FileInfo) =
                New DirectoryInfo(logDir).
                    GetFiles(safePreset & "_*.binds", SearchOption.TopDirectoryOnly).
                    OrderByDescending(Function(f) f.LastWriteTimeUtc).
                    ToList()

            If files.Count <= maxKeep Then Return

            For i As Integer = maxKeep To files.Count - 1
                Try
                    files(i).Delete()
                    Log(logger, "[BindReader] backup rotate: deleted " & files(i).FullName)
                Catch ex As Exception
                    Log(logger, "[BindReader] backup rotate failed: " & Trunc(ex.Message, 240))
                End Try
            Next
        Catch ex As Exception
            Log(logger, "[BindReader] rotate error: " & Trunc(ex.Message, 240))
        End Try
    End Sub

    ' ------------------
    ' PATHS / IO
    ' ------------------
    Private Function FindActiveBindsFile() As String
        Try
            Dim dir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                             "Frontier Developments\Elite Dangerous\Options\Bindings")
            If Not Directory.Exists(dir) Then Return ""
            Dim latest As FileInfo = Nothing
            For Each f As String In Directory.GetFiles(dir, "*.binds", SearchOption.TopDirectoryOnly)
                Dim fi As New FileInfo(f)
                If latest Is Nothing OrElse fi.LastWriteTimeUtc > latest.LastWriteTimeUtc Then
                    latest = fi
                End If
            Next
            Return If(latest IsNot Nothing, latest.FullName, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function GetPluginLogDir() As String
        Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VM.ElitePlugin")
    End Function

    Private Function ReadAttr(el As XElement, name As String) As String
        Dim a As XAttribute = el.Attribute(name)
        If a Is Nothing Then Return ""
        Return a.Value
    End Function

    Private Function ReadAttrInt(el As XElement, name As String, defValue As Integer) As Integer
        Dim s As String = ReadAttr(el, name)
        If s.Length = 0 Then Return defValue
        Dim v As Integer
        If Integer.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, v) Then Return v
        Return defValue
    End Function

    Private Function SanitizeFileName(s As String) As String
        Dim bad() As Char = Path.GetInvalidFileNameChars()
        Dim sb As New StringBuilder(s.Length)
        For Each ch As Char In s
            If bad.Contains(ch) Then
                sb.Append("_"c)
            Else
                sb.Append(ch)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function Trunc(s As String, maxLen As Integer) As String
        If s Is Nothing Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen)
    End Function

    Private Sub SafeSet(varName As String, value As String, logger As Action(Of String))
        Try
            Dim cur As String = ""
            Try
                cur = VMApiShim.GetVariable(varName, logger)
            Catch
            End Try
            If StringComparer.Ordinal.Compare(cur, value) <> 0 Then
                VMApiShim.SetVariable(varName, value, logger)
            End If
        Catch ex As Exception
            Log(logger, "[BindReader] SetVariable failed: " & varName & " :: " & Trunc(ex.Message, 240))
        End Try
    End Sub

    Private Sub Log(logger As Action(Of String), msg As String)
        If logger IsNot Nothing Then
            Try
                logger(msg)
            Catch
            End Try
        End If
    End Sub

    ' --- watcher wiring ---
    Private Sub EnsureWatcherAfterFirstScan(ByVal logger As Action(Of String))
        If logger Is Nothing Then Exit Sub
        SyncLock _sync
            _log = logger
            If _bootstrapped Then Exit Sub

            Dim bindsDir As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Frontier Developments\Elite Dangerous\Options\Bindings")

            If Not Directory.Exists(bindsDir) Then
                Try : logger("[BindReader] bindings folder not found: " & bindsDir) : Catch : End Try
                Exit Sub
            End If

            _debounce = New System.Timers.Timer(BINDS_DEBOUNCE_MS)
            _debounce.AutoReset = False
            AddHandler _debounce.Elapsed,
                Sub(_s, _e)
                    Try
                        Dim lg = _log
                        If lg IsNot Nothing Then
                            Try : ScanAndPublish(lg) : Catch ex As Exception : lg("[BindReader] rescan failed: " & ex.Message) : End Try
                        End If
                    Catch
                    End Try
                End Sub

            _watcher = New IO.FileSystemWatcher(bindsDir, "*.binds")
            _watcher.NotifyFilter = IO.NotifyFilters.LastWrite Or IO.NotifyFilters.CreationTime Or IO.NotifyFilters.FileName Or IO.NotifyFilters.Size
            AddHandler _watcher.Changed, AddressOf OnBindsChanged
            AddHandler _watcher.Created, AddressOf OnBindsChanged
            AddHandler _watcher.Renamed, AddressOf OnBindsRenamed
            AddHandler _watcher.Deleted, AddressOf OnBindsChanged
            _watcher.EnableRaisingEvents = True

            Try : logger("[BindReader] watcher started: " & bindsDir) : Catch : End Try
            _bootstrapped = True

            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnProcessExit
        End SyncLock
    End Sub

    Private Sub OnBindsChanged(sender As Object, e As IO.FileSystemEventArgs)
        Try
            Dim t = _debounce
            If t Is Nothing Then Return
            t.Stop()
            t.Start()
        Catch ex As Exception
            Dim lg = _log
            If lg IsNot Nothing Then Try : lg("[BindReader] debounce error: " & ex.Message) : Catch : End Try
        End Try
    End Sub

    Private Sub OnBindsRenamed(sender As Object, e As IO.RenamedEventArgs)
        OnBindsChanged(sender, e)
    End Sub

    Private Sub OnProcessExit(sender As Object, e As EventArgs)
        Try
            SyncLock _sync
                If _watcher IsNot Nothing Then
                    Try
                        _watcher.EnableRaisingEvents = False
                        RemoveHandler _watcher.Changed, AddressOf OnBindsChanged
                        RemoveHandler _watcher.Created, AddressOf OnBindsChanged
                        RemoveHandler _watcher.Renamed, AddressOf OnBindsRenamed
                        RemoveHandler _watcher.Deleted, AddressOf OnBindsChanged
                        _watcher.Dispose()
                    Catch
                    End Try
                    _watcher = Nothing
                End If
                If _debounce IsNot Nothing Then
                    Try
                        _debounce.Stop()
                        _debounce.Dispose()
                    Catch
                    End Try
                    _debounce = Nothing
                End If
            End SyncLock
        Catch
        End Try
    End Sub

    ' ---- write a readable binds report into the plugin log directory ----
    Private Function WriteBindsReport(snap As Snapshot, bindsPath As String, logger As Action(Of String)) As String
        Try
            Dim logDir As String = GetPluginLogDir()
            If String.IsNullOrEmpty(logDir) Then Return ""
            If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)

            Dim outPath As String = Path.Combine(logDir, "ED_Binds_Report.txt")

            Dim sb As New StringBuilder()
            sb.AppendLine("Elite Dangerous — Keybinds Snapshot")
            sb.AppendLine("Preset : " & If(String.IsNullOrEmpty(snap.PresetName), "(unknown)", snap.PresetName))
            sb.AppendLine("Source : " & bindsPath)
            sb.AppendLine("Time   : " & DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) & "Z")
            sb.AppendLine()
            sb.AppendLine(BuildReport(snap))

            File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8)
            Log(logger, "[BindReader] report -> " & outPath)
            Return outPath
        Catch ex As Exception
            Log(logger, "[BindReader] report failed: " & Trunc(ex.Message, 240))
            Return ""
        End Try
    End Function
    ' -------------------------------------------------------------------------

End Module
