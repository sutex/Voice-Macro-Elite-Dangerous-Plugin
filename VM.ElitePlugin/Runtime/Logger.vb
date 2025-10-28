Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text

Public Class Logger
    Private ReadOnly _path As String
    Private ReadOnly _sync As New Object
    Private ReadOnly _capBytes As Long

    ' capBytes: max size for the single log (default 1 MB)
    Public Sub New(filePath As String, Optional capBytes As Long = 1048576)
        _path = filePath
        ' enforce a sane minimum cap so we don't trim too aggressively
        _capBytes = If(capBytes < 32768, 32768, capBytes)

        Try
            Dim dir = Path.GetDirectoryName(_path)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
        Catch
            ' best effort; don't throw from logger ctor
        End Try
    End Sub

    Public Sub Write(message As String)
        Dim line As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & " " & message & Environment.NewLine
        SyncLock _sync
            Try
                File.AppendAllText(_path, line, Encoding.UTF8)
                EnsureSizeCap()
            Catch
                ' never throw from logger
            End Try
        End SyncLock
    End Sub

    Private Sub EnsureSizeCap()
        Try
            Dim fi As New FileInfo(_path)
            If Not fi.Exists Then Return
            If fi.Length <= _capBytes Then Return

            ' Keep newest ~75% of cap to avoid trimming too often
            Dim keepBytes As Long = CLng(Math.Max(_capBytes * 0.75, 32768))

            Using fs As New FileStream(_path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)
                If fs.Length <= keepBytes Then Return

                fs.Position = fs.Length - keepBytes

                ' Read the tail we want to keep
                Using sr As New StreamReader(fs, Encoding.UTF8, detectEncodingFromByteOrderMarks:=True, bufferSize:=8192, leaveOpen:=True)
                    Dim tail As String = sr.ReadToEnd()

                    ' Rewrite file with truncation banner + kept tail
                    fs.SetLength(0)
                    Using sw As New StreamWriter(fs, New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False))
                        sw.WriteLine("[log truncated at {0}; kept {1} of {2} bytes]",
                                     DateTime.UtcNow.ToString("u"), keepBytes, fi.Length)
                        sw.Write(tail)
                        sw.Flush()
                    End Using
                End Using
            End Using
        Catch
            ' swallow errors; logging must never crash the plugin
        End Try
    End Sub
End Class
