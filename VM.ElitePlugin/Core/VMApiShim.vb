Option Explicit On
Option Strict On

Imports System
Imports System.Reflection

' VoiceMacro API shim — GUIDP-aware (no fake "scope" param).
' IMPORTANT:
'   - VoiceMacro determines scope **by variable NAME suffix** (_p/_g/_s).
'   - SetVariable(name As String, value As String, Optional GUIDP As String = "")
'   - GetVariable(name As String, Optional GUIDP As String = "")
' We just bind to those signatures and, if a GUIDP overload exists, pass "".
' Bare names (no suffix) are NOT macro-visible – callers should only
' call this shim with *_p/_g/_s when they want macros to see it.
Public NotInheritable Class VMApiShim
    Private Sub New()
    End Sub

    Private Shared _setMI As MethodInfo
    Private Shared _getMI As MethodInfo
    Private Shared _inst As Object

    Private Shared _aritySet As Integer = 0 ' 3 = (name,value,GUIDP) | 2 = (name,value)
    Private Shared _arityGet As Integer = 0 ' 2 = (name,GUIDP) | 1 = (name)
    Private Shared _checked As Boolean

    ' NEW: serialize first bind to be strictly race-free
    Private Shared ReadOnly _bindLock As New Object()

    Private Shared ReadOnly SetNames As String() = {"SetVariable", "SetVar", "SetTextVariable", "SetTextVar"}
    Private Shared ReadOnly GetNames As String() = {"GetVariable", "GetVar", "GetTextVariable", "GetTextVar"}

    ' -----------------------------
    ' Binding
    ' -----------------------------
    Public Shared Sub EnsureBound(log As Action(Of String))
        ' Fast path: if both are already resolved, leave immediately.
        If _setMI IsNot Nothing AndAlso _getMI IsNot Nothing Then Exit Sub

        SyncLock _bindLock
            ' Re-check under lock.
            If _setMI IsNot Nothing AndAlso _getMI IsNot Nothing Then Exit Sub

            If _checked Then Exit Sub
            _checked = True
            Try
                Dim probes As String() = {
                    "vmAPI.vmCommand, vmAPI",
                    "vmAPI.API, vmAPI",
                    "VoiceMacro.API, VoiceMacro",
                    "vmAPI.vmCommand",
                    "vmAPI.API"
                }
                For Each q In probes
                    Dim t = Type.GetType(q, False)
                    If t IsNot Nothing AndAlso TryBindType(t, log) Then Exit Sub
                Next

                ' Fallback: scan loaded assemblies
                For Each a In AppDomain.CurrentDomain.GetAssemblies()
                    Dim ts() As Type = {}
                    Try : ts = a.GetTypes() : Catch : Continue For
                    End Try
                    For Each t In ts
                        If TryBindType(t, log) Then Exit Sub
                    Next
                Next
            Catch ex As Exception
                log?.Invoke("[VMApiShim] EnsureBound error: " & ex.Message)
            Finally
                _checked = False ' allow retry later
            End Try
        End SyncLock
    End Sub

    Private Shared Function TryBindType(t As Type, log As Action(Of String)) As Boolean
        Dim setMI As MethodInfo = Nothing, getMI As MethodInfo = Nothing
        Dim inst As Object = Nothing
        Dim arSet As Integer = 0, arGet As Integer = 0

        ' --- Static preferred ---
        ' SetVariable: prefer 3-arg (name,value,GUIDP) else 2-arg
        For Each n In SetNames
            Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Static)
            If m Is Nothing Then Continue For
            Dim p = m.GetParameters()
            If p.Length = 3 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) AndAlso p(2).ParameterType Is GetType(String) Then
                setMI = m : arSet = 3 : Exit For
            End If
        Next
        If setMI Is Nothing Then
            For Each n In SetNames
                Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Static)
                If m Is Nothing Then Continue For
                Dim p = m.GetParameters()
                If p.Length = 2 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) Then
                    setMI = m : arSet = 2 : Exit For
                End If
            Next
        End If

        ' GetVariable: prefer 2-arg (name,GUIDP) else 1-arg
        For Each n In GetNames
            Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Static)
            If m Is Nothing Then Continue For
            Dim p = m.GetParameters()
            If p.Length = 2 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) Then
                getMI = m : arGet = 2 : Exit For
            End If
        Next
        If getMI Is Nothing Then
            For Each n In GetNames
                Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Static)
                If m Is Nothing Then Continue For
                Dim p = m.GetParameters()
                If p.Length = 1 AndAlso p(0).ParameterType Is GetType(String) Then
                    getMI = m : arGet = 1 : Exit For
                End If
            Next
        End If

        ' --- Instance fallback ---
        If setMI Is Nothing OrElse getMI Is Nothing Then
            Dim iSet As MethodInfo = Nothing, iGet As MethodInfo = Nothing
            Dim iArSet As Integer = 0, iArGet As Integer = 0

            For Each n In SetNames
                Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Instance)
                If m Is Nothing Then Continue For
                Dim p = m.GetParameters()
                If p.Length = 3 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) AndAlso p(2).ParameterType Is GetType(String) Then
                    iSet = m : iArSet = 3 : Exit For
                End If
            Next
            If iSet Is Nothing Then
                For Each n In SetNames
                    Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Instance)
                    If m Is Nothing Then Continue For
                    Dim p = m.GetParameters()
                    If p.Length = 2 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) Then
                        iSet = m : iArSet = 2 : Exit For
                    End If
                Next
            End If

            For Each n In GetNames
                Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Instance)
                If m Is Nothing Then Continue For
                Dim p = m.GetParameters()
                If p.Length = 2 AndAlso p(0).ParameterType Is GetType(String) AndAlso p(1).ParameterType Is GetType(String) Then
                    iGet = m : iArGet = 2 : Exit For
                End If
            Next
            If iGet Is Nothing Then
                For Each n In GetNames
                    Dim m = t.GetMethod(n, BindingFlags.Public Or BindingFlags.Instance)
                    If m Is Nothing Then Continue For
                    Dim p = m.GetParameters()
                    If p.Length = 1 AndAlso p(0).ParameterType Is GetType(String) Then
                        iGet = m : iArGet = 1 : Exit For
                    End If
                Next
            End If

            If iSet IsNot Nothing AndAlso iGet IsNot Nothing Then
                Try
                    Dim prop = t.GetProperty("Instance", BindingFlags.Public Or BindingFlags.Static)
                    If prop IsNot Nothing Then inst = prop.GetValue(Nothing, Nothing)
                    If inst Is Nothing Then
                        Dim ci = t.GetConstructor(Type.EmptyTypes)
                        If ci IsNot Nothing Then inst = ci.Invoke(Nothing)
                    End If
                Catch
                End Try
                setMI = iSet : arSet = iArSet
                getMI = iGet : arGet = iArGet
            End If
        End If

        If setMI IsNot Nothing AndAlso getMI IsNot Nothing Then
            _setMI = setMI : _getMI = getMI : _inst = inst
            _aritySet = arSet : _arityGet = arGet
            Try
                log?.Invoke(If(_aritySet = 3, "[VMApiShim] SetVariable arity=3 (GUIDP)", "[VMApiShim] SetVariable arity=2"))
                log?.Invoke(If(_arityGet = 2, "[VMApiShim] GetVariable arity=2 (GUIDP)", "[VMApiShim] GetVariable arity=1"))
            Catch
            End Try
            log?.Invoke("[VMApiShim] Bound to " & t.FullName)
            Return True
        End If

        Return False
    End Function

    ' -----------------------------
    ' Public API
    ' -----------------------------
    Public Shared Sub SetVariable(name As String, value As String, log As Action(Of String))
        If _setMI Is Nothing Then EnsureBound(log)
        If _setMI Is Nothing Then Exit Sub
        Try
            Dim v As String = If(value, "")
            If _aritySet = 3 Then
                _setMI.Invoke(_inst, New Object() {name, v, ""}) ' GUIDP: current profile
            Else
                _setMI.Invoke(_inst, New Object() {name, v})
            End If
        Catch ex As Exception
            log?.Invoke("[VMApiShim] SetVariable error: " & ex.Message)
        End Try
    End Sub

    Public Shared Function GetVariable(name As String, log As Action(Of String)) As String
        If _getMI Is Nothing Then EnsureBound(log)
        If _getMI Is Nothing Then Return ""
        Try
            Dim v As Object
            If _arityGet = 2 Then
                v = _getMI.Invoke(_inst, New Object() {name, ""}) ' GUIDP: current profile
            Else
                v = _getMI.Invoke(_inst, New Object() {name})
            End If
            If v Is Nothing Then Return ""
            Return v.ToString()
        Catch ex As Exception
            log?.Invoke("[VMApiShim] GetVariable error: " & ex.Message)
            Return ""
        End Try
    End Function
End Class
