Option Explicit On
Option Strict On

Imports System
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' Handles safe JSON persistence with atomic writes and corruption safeguards.
' No vm_* variables are introduced here.
Public Class Persistence
    Private ReadOnly _baseDir As String

    Public Sub New(baseDir As String)
        _baseDir = baseDir
        If Not Directory.Exists(_baseDir) Then Directory.CreateDirectory(_baseDir)
    End Sub

    ' Write JSON atomically: write to temp, then swap in.
    Public Sub WriteJsonAtomic(fileName As String, obj As JObject, log As Action(Of String))
        Try
            Dim fullPath As String = System.IO.Path.Combine(_baseDir, fileName)
            Dim tmpPath As String = fullPath & ".tmp"
            Dim bakPath As String = fullPath & ".bak"

            Dim json As String = obj.ToString(Formatting.Indented)
            File.WriteAllText(tmpPath, json, Encoding.UTF8)

            ' Keep a single backup and ensure atomic-ish replace for our purposes
            If File.Exists(bakPath) Then
                Try : File.Delete(bakPath) : Catch : End Try
            End If
            If File.Exists(fullPath) Then
                Try
                    File.Move(fullPath, bakPath)
                Catch
                    ' If move fails, try delete (last resort)
                    Try : File.Delete(fullPath) : Catch : End Try
                End Try
            End If

            ' Move temp into place
            File.Move(tmpPath, fullPath)
        Catch ex As Exception
            log?.Invoke("[Persistence] WriteJsonAtomic error: " & ex.Message)
        End Try
    End Sub

    ' Read JSON safely: returns empty object on error or missing file.
    Public Function ReadJsonOrEmpty(fileName As String, log As Action(Of String)) As JObject
        Try
            Dim fullPath As String = System.IO.Path.Combine(_baseDir, fileName)
            If Not File.Exists(fullPath) Then Return New JObject()
            Dim text As String = File.ReadAllText(fullPath, Encoding.UTF8)
            If String.IsNullOrWhiteSpace(text) Then Return New JObject()
            Return JObject.Parse(text)
        Catch ex As Exception
            log?.Invoke("[Persistence] ReadJsonOrEmpty error: " & ex.Message)
            Return New JObject()
        End Try
    End Function
End Class
