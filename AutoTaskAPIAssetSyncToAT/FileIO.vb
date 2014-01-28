Imports System
Imports System.Reflection
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.Configuration
Imports System.Diagnostics
Public Class FileIO

    Public Enum NotifyType
        INFO = 1
        ERR = 2
        WARNING = 3
    End Enum

    Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

    Public INIFilePath = Directory.GetCurrentDirectory().ToString() & "\\INI\\"
    Public INIFileName = "dbstring.ini"

    Dim ExeFileFullName = System.Reflection.Assembly.GetEntryAssembly().Location
    Public ExeFileName = System.IO.Path.GetFileNameWithoutExtension(ExeFileFullName)

    Public LogFilePath = Directory.GetCurrentDirectory().ToString() & "\\LOG\\"
    Public LogFileName = ExeFileName & "_" & DateTime.Now.ToString("ddMMyyyy") & ".log"

    Public Function CreateLogFile(ByVal FilePath As String, ByVal FileName As String) As Boolean
        Try
            If Directory.Exists(FilePath) = True Then
                Directory.CreateDirectory(FilePath)
                If File.Exists(FilePath + FileName) = False Then
                    Dim str As String
                    str = ""
                    Using file As StreamWriter = New StreamWriter((FilePath & FileName), True)
                        str = "#SOFTWARE: SAAZ.ITS.DC." + FileName.ToString()
                        file.WriteLine(str)
                        str = "#VERSION: " & System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Major & "." & System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Minor & "." & System.Reflection.Assembly.GetEntryAssembly().GetName().Version.Revision
                        file.WriteLine(str)
                        str = "#DATE: " & DateTime.Now.ToString()
                        file.WriteLine(str)
                        str = "#FIELDS: DTIME" & vbTab & "MODULE" & vbTab & "TYPE" & vbTab & vbTab & "VAL1" & vbTab & "VAL2" & vbTab & "VAL3"
                        file.WriteLine(str)
                    End Using
                End If
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Function LogNotify(ByVal MethodName As String, ByVal NotifyTypes As NotifyType, ByVal Value1 As String, Optional ByVal Value2 As String = "", Optional ByVal Value3 As String = "")
        Try
            If CreateLogFile(LogFilePath, LogFileName) = True Then
                Dim str As String
                Dim MNotiFy As String
                str = ""
                If NotifyTypes.ToString() = "ERR" Then
                    MNotiFy = "ERROR"
                Else
                    MNotiFy = NotifyTypes.ToString()
                End If
                str = "" & DateTime.Now.ToString() & vbTab & MethodName.ToString() & vbTab & MNotiFy & vbTab & Value1.ToString() & vbTab & Value2.ToString() & vbTab & Value3.ToString()
                Using file As StreamWriter = New StreamWriter((LogFilePath & LogFileName), True)
                    file.WriteLine(str)
                End Using
            End If
            Return True
        Catch ex As Exception
        End Try
    End Function

    Public Function CreateDir(ByVal DirNameWithPath As String, ByRef outError As String) As Boolean
        Try
            outError = ""
            If (Directory.Exists(DirNameWithPath)) = False Then
                Try
                    Directory.CreateDirectory(DirNameWithPath)
                    Return True
                Catch ex As Exception
                    outError = ex.Message.ToString()
                    Return False
                End Try
            Else
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function QualifyPath(ByVal sPath As String) As String
        Try
            If Right$(sPath, 1) <> "\" Then
                QualifyPath = sPath & "\"
            Else
                QualifyPath = sPath
            End If
        Catch ex As Exception
            QualifyPath = ""
        End Try
    End Function

    Public Function SetINIKeyValue(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
        WritePrivateProfileString(Section, Key, Value, INIFilePath & INIFileName)
        Return True
    End Function

    Public Function GetINIKeyValue(ByVal Section As String, ByVal Key As String) As String
        Try
            Dim tmpValue As New StringBuilder(255)
            Dim i As Integer
            i = GetPrivateProfileString(Section, Key, "", tmpValue, 255, INIFilePath & INIFileName)
            Return tmpValue.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function getConfigKeyValue(ByVal strSection As String, ByVal strKey As String, ByVal strconfigPath As String, ByRef outResult As String, ByRef outError As String) As Boolean
        outError = ""
        outResult = ""
        'Try
        '    Dim doc As New XmlDocument()
        '    Dim xmlLoc As String
        '    xmlLoc = strconfigPath
        '    doc.Load(xmlLoc)
        '    XmlNodeList(nl = doc.GetElementsByTagName(strSection))
        '    If nl.Count > 0 Then
        '            XmlElement e = (XmlElement)nl[0]
        '            if (e.HasAttribute(strKey) == true)
        '            {
        '                outResult = e.GetAttribute(strKey).ToString();

        '                return true;
        '            }
        '        Else
        '            {
        '                XmlNode node = doc.SelectSingleNode("//" + strSection + "");
        '                XmlElement addElem = (XmlElement)node.SelectSingleNode("//add[@key='" + strKey + "']");
        '                if (addElem != null)
        '                {
        '                    outResult = addElem.GetAttribute("value");
        '                    return true;
        '                }
        '            Else
        '                {
        '                    throw new System.ArgumentException("Key Not Found", "getConfigKeyValue");
        '                }
        '            }
        '        }
        '        else
        '        {
        '            throw new System.ArgumentException("Section Not Found", "getConfigKeyValue");
        '        }
        '    }
        '    catch (Exception ex)
        '    {
        '        outError = ex.Message.ToString();
        '        return false;
        '    }
        Return True
    End Function

End Class
