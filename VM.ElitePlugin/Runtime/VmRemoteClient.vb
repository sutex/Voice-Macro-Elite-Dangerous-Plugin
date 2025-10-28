Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Threading

' Minimal HTTP client for VoiceMacro Remote Control.
' Supports:
'   - path-style:  http://host:port/ExecuteMacro=<ProfileName>/<MacroName>
'   - generic:     http://host:port/ExecuteMacro=<MacroName>                (some VM builds reject generic)
Public Class VmRemoteClient

    Private ReadOnly _log As Action(Of String)

    ' Default to localhost per current requirement
    Private _baseUrl As String = "http://localhost:8080/"
    Private _debounceMs As Integer = 350

    ' Per-macro debounce (macro name -> last tick)
    Private ReadOnly _lastSent As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

    Public Sub New(logger As Action(Of String))
        If logger Is Nothing Then
            _log = Sub(s As String)
                   End Sub
        Else
            _log = logger
        End If
    End Sub

    ' -----------------------------
    ' Public config
    ' -----------------------------
    Public Sub SetBaseUrl(baseUrl As String)
        Dim u As String = If(baseUrl, "").Trim()
        If u.Length = 0 Then Return
        If Not u.EndsWith("/") Then u &= "/"
        _baseUrl = u
        _log("[VMRemoteClient] BaseUrl set → " & _baseUrl)
    End Sub

    Public Sub SetDebounce(ms As Integer)
        If ms < 0 Then ms = 0
        _debounceMs = ms
        _log("[VMRemoteClient] Debounce set → " & _debounceMs.ToString() & " ms")
    End Sub

    ' -----------------------------
    ' Generic ExecuteMacro (may be rejected depending on VM build)
    ' -----------------------------
    Public Sub TrySend(macroName As String)
        Dim m As String = If(macroName, "").Trim()
        If m.Length = 0 Then Exit Sub

        ' Per-macro debounce
        Dim now As Integer = Environment.TickCount
        SyncLock _lastSent
            Dim last As Integer = 0
            If _lastSent.TryGetValue(m, last) Then
                Dim delta As Integer = now - last
                If delta >= 0 AndAlso delta < _debounceMs Then Exit Sub
            End If
            _lastSent(m) = now
        End SyncLock

        ThreadPool.QueueUserWorkItem(
            Sub(_state As Object)
                Dim baseA As String = If(_baseUrl.EndsWith("/"), _baseUrl, _baseUrl & "/")
                Dim url As String = baseA & "ExecuteMacro=" & Uri.EscapeDataString(m)

                Try
                    If SendOnce(url) Then Exit Sub
                    Exit Sub

                Catch ex As WebException
                    Dim http As HttpWebResponse = TryCast(ex.Response, HttpWebResponse)
                    Dim is400 As Boolean = (http IsNot Nothing AndAlso CInt(http.StatusCode) = 400)

                    If is400 Then
                        Dim alt As String = SwapLocalhost(baseA)
                        If String.IsNullOrEmpty(alt) Then
                            _log("[VMRemoteClient] 400 Invalid Hostname; base host is not localhost/127.0.0.1 — no retry.")
                            Exit Sub
                        End If

                        Dim altUrl As String = alt & "ExecuteMacro=" & Uri.EscapeDataString(m)
                        _log("[VMRemoteClient] 400 Invalid Hostname → retry " & altUrl)

                        Try
                            If SendOnce(altUrl) Then
                                _baseUrl = alt
                                _log("[VMRemoteClient] Host auto-correct → " & _baseUrl)
                            End If
                        Catch rex As Exception
                            _log("[VMRemoteClient] Retry failed: " & rex.Message)
                        End Try

                        Exit Sub
                    End If

                    _log("[VMRemoteClient] WebException (" & ex.Status.ToString() & "): " & ex.Message)

                Catch ex As Exception
                    _log("[VMRemoteClient] ExecuteMacro failed: " & ex.Message)
                End Try
            End Sub)
    End Sub

    ' -----------------------------
    ' Path-style ExecuteMacro with ProfileName (this is what worked for you)
    '   Example: http://localhost:8080/ExecuteMacro=Elite_Dangerous_main/ED_CommsOpen
    ' -----------------------------
    Public Sub TrySendWithProfile(profileName As String, macroName As String)
        Dim m As String = If(macroName, "").Trim()
        If m.Length = 0 Then Exit Sub

        ' Per-macro debounce (same logic as TrySend)
        Dim now As Integer = Environment.TickCount
        SyncLock _lastSent
            Dim last As Integer = 0
            If _lastSent.TryGetValue(m, last) Then
                Dim delta As Integer = now - last
                If delta >= 0 AndAlso delta < _debounceMs Then Exit Sub
            End If
            _lastSent(m) = now
        End SyncLock

        ThreadPool.QueueUserWorkItem(
    Sub(state As Object)
        Dim baseA As String = If(_baseUrl.EndsWith("/"), _baseUrl, _baseUrl & "/")

        Dim url As String
        If Not String.IsNullOrWhiteSpace(profileName) Then
            Dim right As String = profileName & "/" & m
            url = baseA & "ExecuteMacro=" & Uri.EscapeDataString(right)
        Else
            url = baseA & "ExecuteMacro=" & Uri.EscapeDataString(m)
        End If

        Try
            If SendOnce(url) Then Exit Sub

        Catch ex As WebException
            Dim http = TryCast(ex.Response, HttpWebResponse)
            If http IsNot Nothing AndAlso CInt(http.StatusCode) = 400 Then
                Dim alt As String = SwapLocalhost(baseA)
                If Not String.IsNullOrEmpty(alt) Then
                    Dim altUrl As String
                    If Not String.IsNullOrWhiteSpace(profileName) Then
                        altUrl = alt & "ExecuteMacro=" & Uri.EscapeDataString(profileName & "/" & m)
                    Else
                        altUrl = alt & "ExecuteMacro=" & Uri.EscapeDataString(m)
                    End If
                    _log("[VMRemoteClient] 400 Invalid Hostname → retry " & altUrl)
                    Try
                        If SendOnce(altUrl) Then
                            _baseUrl = alt
                            _log("[VMRemoteClient] Host auto-correct → " & _baseUrl)
                        End If
                    Catch rex As Exception
                        _log("[VMRemoteClient] Retry failed: " & rex.Message)
                    End Try
                Else
                    _log("[VMRemoteClient] 400 Invalid Hostname; no localhost/127 alternative.")
                End If
                Exit Sub
            End If

            _log("[VMRemoteClient] WebException (" & ex.Status.ToString() & "): " & ex.Message)
        Catch ex As Exception
            _log("[VMRemoteClient] ExecuteMacro failed: " & ex.Message)
        End Try
    End Sub)
    End Sub

    ' -----------------------------
    ' Helpers
    ' -----------------------------
    Private Function SendOnce(fullUrl As String) As Boolean
        _log("[VMRemoteClient] GET " & fullUrl)

        Dim req As HttpWebRequest = CType(WebRequest.Create(fullUrl), HttpWebRequest)
        req.Method = "GET"
        req.Timeout = 2500
        req.ReadWriteTimeout = 2500
        req.AllowAutoRedirect = False
        req.Proxy = Nothing
        req.KeepAlive = False
        req.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
        req.UserAgent = "VM.ElitePlugin/1.0"

        Using resp As HttpWebResponse = CType(req.GetResponse(), HttpWebResponse)
            Dim code As Integer = CInt(resp.StatusCode)
            If code = 200 Then Return True

            Dim body As String = ""
            Using s As Stream = resp.GetResponseStream()
                If s IsNot Nothing Then
                    Using r As New StreamReader(s, Encoding.UTF8)
                        body = r.ReadToEnd()
                    End Using
                End If
            End Using
            _log("[VMRemoteClient] HTTP " & code.ToString() & " from VM (ExecuteMacro): " & body)
            Return False
        End Using
    End Function

    ' Flip localhost <-> 127.0.0.1 in the base URL; return new base (with trailing slash) or Nothing.
    Private Shared Function SwapLocalhost(baseWithSlash As String) As String
        Try
            Dim u As Uri = Nothing
            If Not Uri.TryCreate(baseWithSlash, UriKind.Absolute, u) Then Return Nothing
            Dim host As String = u.Host.Trim().ToLowerInvariant()

            Dim altHost As String = Nothing
            If host = "localhost" Then
                altHost = "127.0.0.1"
            ElseIf host = "127.0.0.1" Then
                altHost = "localhost"
            Else
                Return Nothing
            End If

            Dim b As New UriBuilder(u)
            b.Host = altHost
            If String.IsNullOrEmpty(b.Path) OrElse Not b.Path.EndsWith("/") Then b.Path = (If(b.Path, "") & "/")
            Return b.Uri.AbsoluteUri
        Catch
            Return Nothing
        End Try
    End Function

End Class
