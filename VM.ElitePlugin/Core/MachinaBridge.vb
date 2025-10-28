Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' -----------------------------------------------------------------------------
' MachinaBridge
' -----------------------------------------------------------------------------
' Minimal helper to expose a simple JSON "signal" for Machina or any external
' tool to read. You can toggle state or publish a snapshot; writes are atomic.
'
' Files:
'   %LocalAppData%\VM.ElitePlugin\Machina.signal.json
' -----------------------------------------------------------------------------
Public Class MachinaBridge
    Private ReadOnly _log As Action(Of String)
    Private ReadOnly _baseDir As String
    Private ReadOnly _signalPath As String

    Public Sub New(log As Action(Of String))
        If log Is Nothing Then
            _log = Sub(__ As String)
                   End Sub
        Else
            _log = log
        End If

        _baseDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VM.ElitePlugin")
        If Not Directory.Exists(_baseDir) Then Directory.CreateDirectory(_baseDir)
        _signalPath = System.IO.Path.Combine(_baseDir, "Machina.signal.json")
    End Sub

    ' Toggle-style signal (e.g., when user runs aion/aioff/aitoggle).
    Public Sub SetEnabled(enabled As Boolean)
        Dim obj As New JObject()
        obj("enabled") = enabled
        obj("kind") = "toggle"
        obj("timestamp_utc") = DateTime.UtcNow.ToString("o")
        WriteAtomic(obj)
    End Sub

    ' Snapshot of variables (optional). Provide whatever subset you want.
    Public Sub PublishSnapshot(values As Dictionary(Of String, String))
        If values Is Nothing Then values = New Dictionary(Of String, String)()
        Dim obj As New JObject()
        obj("kind") = "snapshot"
        obj("timestamp_utc") = DateTime.UtcNow.ToString("o")

        Dim v As New JObject()
        For Each kv In values
            v(kv.Key) = If(kv.Value, "")
        Next
        obj("vars") = v

        WriteAtomic(obj)
    End Sub

    Private Sub WriteAtomic(obj As JObject)
        Try
            Dim json As String = obj.ToString(Formatting.Indented)
            Dim tmp As String = _signalPath & ".tmp"
            Dim bak As String = _signalPath & ".bak"

            File.WriteAllText(tmp, json, Encoding.UTF8)

            If File.Exists(bak) Then
                Try : File.Delete(bak) : Catch : End Try
            End If
            If File.Exists(_signalPath) Then
                Try
                    File.Move(_signalPath, bak)
                Catch
                    Try : File.Delete(_signalPath) : Catch : End Try
                End Try
            End If

            File.Move(tmp, _signalPath)
            _log("[Machina] Wrote signal: " & _signalPath)
        Catch ex As Exception
            _log("[Machina] Write error: " & ex.Message)
        End Try
    End Sub
End Class
