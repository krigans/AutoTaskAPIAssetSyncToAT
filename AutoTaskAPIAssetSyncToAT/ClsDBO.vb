Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Public Class ClsDBO
    ' To Open a database sqlconnection with given connectionstring.
    Public Function openConnection(ByRef ConnTemp As SqlConnection, ByVal strConnString As String, ByRef outError As String, Optional ByVal contimeout As Integer = 60) As Boolean
        Dim FileIO As New FileIO
        Try
            Dim result As String
            If (strConnString.ToLower().IndexOf("connection timeout") <= 0) Then
                result = strConnString.Substring(strConnString.Length - 1, 1)
                If (result <> ";") Then
                    strConnString = strConnString & ";connection timeout=" & contimeout & ";"
                Else
                    strConnString = strConnString & "connection timeout=" & contimeout & ";"
                End If
            End If

            If (strConnString.ToLower().IndexOf("multipleactiveresultSets") <= 0) Then
                result = strConnString.Substring(strConnString.Length - 1, 1)
                If (result <> ";") Then
                    strConnString = strConnString & ";MultipleActiveResultSets=True;"
                Else
                    strConnString = strConnString & "MultipleActiveResultSets=True;"
                End If
            End If

            ConnTemp = New SqlConnection(strConnString)
            ConnTemp.Open()
            Return True
        Catch ex As Exception
            FileIO.LogNotify("openConnection", FileIO.NotifyType.ERR, ex.StackTrace)
            Return False
        End Try
    End Function
    Public Function openConnection(ByRef ConnTemp As SqlConnection, ByVal strSection As String, ByVal strKey As String, ByVal strconfigPath As String, ByRef outError As String, Optional ByVal contimeout As Integer = 60)

        'Dim strConnString As String
        'Dim OutKeyValue As String
        'Dim IsGetValue As Boolean
        'Dim result As String
        'outError = ""
        'OutKeyValue = ""
        'IsGetValue = FileIO.getConfigKeyValue(strSection, strKey, strconfigPath, OutKeyValue, outError)
        'If (IsGetValue = True) Then
        '    strConnString = OutKeyValue
        '    If (strConnString.ToLower().IndexOf("connection timeout") <= 0) Then
        '        result = strConnString.Substring(strConnString.Length - 1, 1)
        '        If (result <> ";") Then
        '            strConnString = strConnString + ";connection timeout=" + contimeout + ";"
        '        Else
        '            strConnString = strConnString + "connection timeout=" + contimeout + ";"
        '        End If
        '    End If
        '    If (strConnString.ToLower().IndexOf("multipleactiveresultSets") <= 0) Then
        '        result = strConnString.Substring(strConnString.Length - 1, 1)
        '        If (result <> ";") Then
        '            strConnString = strConnString + ";MultipleActiveResultSets=True;"
        '        Else
        '            strConnString = strConnString + "MultipleActiveResultSets=True;"
        '        End If
        '    End If
        '    ConnTemp = New SqlConnection(strConnString)
        '    ConnTemp.Open()
        '    Return True
        'End If
        Return True
    End Function

End Class
