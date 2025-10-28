Option Explicit On
Option Strict On

Imports System
Imports System.Diagnostics

' -----------------------------------------------------------------------------
' DiscordPresence  (simple + wing-gated)
' -----------------------------------------------------------------------------
' • Presence is OFF when not in wing unless a pre-wing one-shot is set.
' • One-shot (e.g., "LFW — <System>") shows only while NOT in wing.
' • In wing, presence shows two lines:
'     Line 1: Where (Docked / Supercruise / In Space) with station/body/system.
'     Line 2: "Engaging {name}" when vm_target_locked=1 (reads vm_target_name_s).
' • No on-foot, no surface wording, no music.
' • Debounce (~350 ms) and publish only on change.
' -----------------------------------------------------------------------------
Public Class DiscordPresence
    ' ---- Publishing & logging ----
    Private ReadOnly _publish As Action(Of String)
    Private ReadOnly _log As Action(Of String)

    ' ---- Runtime state ----
    Private _enabled As Boolean = True
    Private _inWing As Boolean = False
    Private _oneShot As String = ""

    Private _lastLine1 As String = Nothing
    Private _lastLine2 As String = Nothing

    Private _debounceMs As Integer = 350
    Private ReadOnly _sw As Stopwatch = Stopwatch.StartNew()

    ' Optional VM getter for target lock & name, and optional oneshot read
    Private _vmGet As Func(Of String, String) = Nothing

    ' Safe no-op delegate so we don't inline multi-line lambdas inside If(...)
    Private Shared ReadOnly NoOp As Action(Of String) = Sub(__ As String)
                                                        End Sub

    ' ---------------- Constructors ----------------

    Public Sub New(publish As Action(Of String))
        _publish = If(publish, NoOp)
        _log = NoOp
    End Sub

    Public Sub New(publish As Action(Of String), logger As Action(Of String))
        _publish = If(publish, NoOp)
        _log = If(logger, NoOp)
    End Sub

    ' ---------------- Public API ----------------

    Public Sub StartPresence()
        _enabled = True
        ' No immediate publish; caller drives UpdateFromSnapshot.
    End Sub

    Public Sub StopPresence()
        _enabled = False
        PushIfChanged("", "", True)
    End Sub

    ' Pre-wing one-shot (only used when not in wing)
    Public Sub ShowOneShot(text As String)
        _oneShot = If(text, "")
        ' Published on next UpdateFromSnapshot
    End Sub

    ' Wing state (gates presence); joining wing clears local one-shot
    Public Sub SetWingState(inWing As Boolean)
        _inWing = inWing
        If inWing Then _oneShot = ""
    End Sub

    ' Optional VM getter (to read vm_target_* and vm_presence_oneshot_s)
    Public Sub SetVmGetter(vmGet As Func(Of String, String))
        _vmGet = vmGet
    End Sub

    ' Optional: debounce tuning
    Public Sub SetDebounce(milliseconds As Integer)
        If milliseconds >= 0 Then _debounceMs = milliseconds
    End Sub

    ' ---------------- Updates ----------------

    ' Main overload (6 params) — NO onFoot, NO music
    Public Sub UpdateFromSnapshot(hud As String,
                                  sys As String,
                                  body As String,
                                  docked As Boolean,
                                  inSC As Boolean,
                                  station As String)
        If Not _enabled Then
            PushIfChanged("", "", True)
            Exit Sub
        End If

        ' Wing gating + pre-wing one-shot
        Dim haveOneShot As Boolean = (Not _inWing) AndAlso HasOneShot()
        If Not _inWing Then
            If haveOneShot Then
                PushIfChanged(CurrentOneShot(), "")
            Else
                PushIfChanged("", "")
            End If
            Exit Sub
        End If

        ' In wing -> normal two-line presence
        Dim line1 As String = BuildWhereLine(sys, body, docked, inSC, station)
        Dim line2 As String = BuildEngagingLineFromVm() ' empty if no lock or no vm getter
        PushIfChanged(line1, line2)
    End Sub

    ' Safe no-arg update (best-effort from persisted VM vars if a getter was provided)
    Public Sub UpdateFromSnapshot()
        If _vmGet Is Nothing Then Exit Sub

        Dim sys As String = Nz(_vmGet("vm_star_system_p"))
        If sys = "" Then sys = Nz(_vmGet("vm_star_system"))

        Dim body As String = Nz(_vmGet("vm_body_p"))
        If body = "" Then body = Nz(_vmGet("vm_body"))

        Dim docked As Boolean = (_vmGet("vm_docked_p") = "1" OrElse _vmGet("vm_docked") = "1")
        Dim inSC As Boolean = (_vmGet("vm_in_supercruise_p") = "1" OrElse _vmGet("vm_in_supercruise") = "1")

        Dim station As String = Nz(_vmGet("vm_station_name_p"))
        If station = "" Then station = Nz(_vmGet("vm_station_name"))

        UpdateFromSnapshot("", sys, body, docked, inSC, station)
    End Sub


    ' ---------------- Internal helpers ----------------

    Private Function HasOneShot() As Boolean
        ' Prefer explicit ShowOneShot; allow VM-provided oneshot as fallback
        Dim s As String = _oneShot
        If (s = "" OrElse s = "0") AndAlso _vmGet IsNot Nothing Then
            s = Nz(_vmGet("vm_presence_oneshot_s"))
        End If
        Return Not String.IsNullOrEmpty(s) AndAlso s <> "0"
    End Function

    Private Function CurrentOneShot() As String
        Dim s As String = _oneShot
        If (s = "" OrElse s = "0") AndAlso _vmGet IsNot Nothing Then
            s = Nz(_vmGet("vm_presence_oneshot_s"))
        End If
        Return s
    End Function

    Private Function BuildWhereLine(sys As String,
                                    body As String,
                                    docked As Boolean,
                                    inSC As Boolean,
                                    station As String) As String
        If docked Then
            If station <> "" AndAlso sys <> "" Then
                Return "Docked @ " & station & " — " & sys
            ElseIf station <> "" Then
                Return "Docked @ " & station
            ElseIf sys <> "" Then
                Return "Docked — " & sys
            Else
                Return "Docked"
            End If
        End If

        If inSC Then
            If sys <> "" Then
                Return "Supercruise — " & sys
            Else
                Return "Supercruise"
            End If
        End If

        ' Default (no on-foot handling here)
        If sys <> "" Then
            Return "In Space — " & sys
        Else
            Return "In Space"
        End If
    End Function

    ' Uses optional vm getter to show "Engaging …"
    Private Function BuildEngagingLineFromVm() As String
        If _vmGet Is Nothing Then Return ""
        Dim locked As Boolean = (_vmGet("vm_target_locked") = "1")
        If Not locked Then Return ""
        Dim name As String = Nz(_vmGet("vm_target_name_s"))
        If String.IsNullOrEmpty(name) Then Return ""
        Return "Engaging " & name
    End Function

    ' ---------------- Publish with debounce ----------------

    Private Sub PushIfChanged(line1 As String, line2 As String, Optional force As Boolean = False)
        Dim same As Boolean = (String.Equals(line1, _lastLine1, StringComparison.Ordinal) AndAlso
                               String.Equals(line2, _lastLine2, StringComparison.Ordinal))
        If same AndAlso Not force Then Exit Sub

        If Not force Then
            If _sw.ElapsedMilliseconds < _debounceMs Then Exit Sub
        End If
        _sw.Restart()

        _lastLine1 = line1
        _lastLine2 = line2

        Try
            If String.IsNullOrEmpty(line1) AndAlso String.IsNullOrEmpty(line2) Then
                _publish("") ' clear
            ElseIf String.IsNullOrEmpty(line2) Then
                _publish(line1)
            Else
                _publish(line1 & " ⏐ " & line2)
            End If
        Catch ex As Exception
            _log("DiscordPresence publish error: " & ex.Message)
        End Try
    End Sub

    ' ---------------- Utils ----------------

    Private Shared Function Nz(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Trim()
    End Function
End Class
