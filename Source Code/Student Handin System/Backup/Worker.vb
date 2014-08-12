Imports System.IO
Imports Microsoft.Win32

Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public username As String
    Public password As String
    Public context As String
    Public group As String
    Public error_level As String = "none"
  

    Public result As String = ""

    Public Event WorkerErrorEncountered(ByVal ex As Exception, ByVal message As String)
    Public Event WorkerComplete(ByVal Result As String, ByVal Thread As Integer)
    


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
      
        username = ""
        password = ""
        context = ""
        group = ""
        error_level = "none"
    
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal exc As Exception, ByVal message As String)
        Try
            RaiseEvent WorkerErrorEncountered(exc, message)
        Catch ex As Exception
            MsgBox("An error occurred in Student Handin System's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.

            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerLogin)
                    WorkerThread.Start()
                Case 2
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerCurrentUser)
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex, "ChooseThreads")
        End Try
    End Sub

    Public Sub WorkerLogin()
        Try

            Dim apptorun As String
            Dim finfo As FileInfo = New FileInfo((Application.StartupPath & "\UCT Novell Login Shell.exe").Replace("\\", "\"))
            If finfo.Exists = True Then
                apptorun = """" & (Application.StartupPath & "\UCT Novell Login Shell.exe").Replace("\\", "\") & """ """ & username & """ """ & password & """ """ & context & """ """ & group & """ """ & error_level & """"

                result = ApplicationLauncher(apptorun)
            Else
                result = "Failure. Login Script Executable cannot be found"
            End If
            finfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "WorkerLogin")
            result = "Failure. Unknown Reason."
        Finally
            RaiseEvent WorkerComplete(result, 1)
        End Try
    End Sub

    Public Sub WorkerCurrentUser()
        Dim result As String = ""
        Try
            result = (ReturnRegKeyValue("HKEY_CURRENT_USER", "Volatile Environment", "NWUSERNAME"))
        Catch ex As Exception
            Error_Handler(ex, "WorkerCurrentUser")
            result = "Failure. Unknown Reason."
        Finally
            RaiseEvent WorkerComplete(result, 2)
        End Try
    End Sub



    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            'myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If



            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Logger(ex, "DosShellCommand")
            Error_Handler(ex, "DosShellCommand")
        End Try

        Return s
    End Function

    Private Function ApplicationLauncher(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process


            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()


        Catch ex As Exception
            Error_Logger(ex, "ApplicationLauncher")
            Error_Handler(ex, "ApplicationLauncher")
        End Try
        Return s
    End Function

    Public Function ReturnRegKeyValue(ByVal MainKey As String, ByVal RequestedKey As String, ByVal Value As String) As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If continue = False Then
                        Exit For
                    End If
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                result = regkey.GetValue(Value)
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            result = "Fail. Could not locate Value within Registry Key"
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "ReturnRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "ReturnRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function

    Public Function SetRegKeyValue(ByVal MainKey As String, ByVal RequestedKey As String, ByVal Value As String, ByVal Data As String) As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If continue = False Then
                        Exit For
                    End If
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                Next
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                'result = regkey.GetValue(Value)
                                regkey.SetValue(Value, Data)
                                result = "Success."
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            result = "Fail. Could not locate Value within Registry Key"
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "SetRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "SetRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function


    Public Function CreateSubRegKey(ByVal MainKey As String, ByVal RequestedKey As String, Optional ByVal Value As String = "", Optional ByVal Data As String = "") As String
        Dim result As String = "Fail."
        Try
            Dim oReg As RegistryKey
            Dim regkey As RegistryKey
            Try
                Select Case MainKey.ToUpper
                    Case "HKEY_CURRENT_USER"
                        oReg = Registry.CurrentUser
                    Case "HKEY_CLASSES_ROOT"
                        oReg = Registry.ClassesRoot
                    Case "HKEY_LOCAL_MACHINE"
                        oReg = Registry.LocalMachine
                    Case "HKEY_USERS"
                        oReg = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        oReg = Registry.CurrentConfig
                    Case Else
                        oReg = Registry.LocalMachine
                End Select

                regkey = oReg
                oReg.Close()
                If RequestedKey.EndsWith("\") = True Then
                    RequestedKey = RequestedKey.Remove(RequestedKey.Length - 1, 1)
                End If
                Dim subs() As String = (RequestedKey).Split("\")
                Dim continue = True
                For Each stri As String In subs
                    If regkey Is Nothing = False Then
                        Dim skn As String() = regkey.GetSubKeyNames()
                        Dim strin As String

                        continue = False
                        For Each strin In skn
                            If stri = strin Then
                                regkey = regkey.OpenSubKey(stri, True)
                                continue = True
                                Exit For
                            End If
                        Next
                    End If
                    If continue = False Then
                        regkey = regkey.CreateSubKey(stri)
                        continue = True
                    End If
                Next
                If Not Value = "" Then
                    result = "Success."
                End If
                If continue = True Then
                    If regkey Is Nothing = False Then
                        Dim str As String() = regkey.GetValueNames()
                        Dim val As String
                        Dim foundit As Boolean = False
                        For Each val In str
                            If Value = val Then
                                foundit = True
                                'result = regkey.GetValue(Value)
                                regkey.SetValue(Value, Data)
                                result = "Success."
                                Exit For
                            End If
                        Next
                        If foundit = False Then
                            If Not Value = "" Then
                                regkey.SetValue(Value, Data)
                            End If
                            result = "Success."
                        End If
                        regkey.Close()
                    End If
                Else
                    result = "Fail. Key cannot be located"
                End If
            Catch ex As Exception
                Error_Logger(ex, "SetRegKeyValue")
                result = "Fail. Check Error Log for further details"
            End Try
        Catch ex As Exception
            Error_Logger(ex, "SetRegKeyValue")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function


    Public Function MapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Try
            Dim continue As Boolean = True
            Dim letterlist As ArrayList = New ArrayList
            letterlist.Clear()
            Dim i As Integer
            For i = 65 To 90
                letterlist.Add(Chr(i).ToString)
            Next

            Dim runner As IEnumerator
            Dim fso As New Scripting.FileSystemObject
            runner = fso.Drives.GetEnumerator()
            While runner.MoveNext() = True
                Dim d As Scripting.Drive
                d = runner.Current()
                letterlist.RemoveAt((letterlist.IndexOf(d.DriveLetter.ToString)))
            End While
            If letterlist.Count > 1 Then
                letterlist.Reverse()
            End If
            resultdrive = ""
            For Each strp As String In letterlist
                resultdrive = strp
                Exit For
            Next
            If resultdrive.Trim = "" Then
                resultdrive = "Fail. No available drive letter can be found"
                continue = False
            End If

            If continue = True Then
                Dim apptorun As String
                'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
                'net use T: /DELETE
                If pathtomap.EndsWith("\") Then
                    pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
                End If
                apptorun = "net use " & resultdrive.Trim.ToUpper & ": """ & pathtomap & """ /PERSISTENT:NO"

                result = DosShellCommand(apptorun)
                If Not result.IndexOf("The command completed successfully.") = -1 Then
                    resultdrive = resultdrive & ":\"
                Else
                    resultdrive = "Fail. Unable to map the give path to a directory"
                End If
            End If
        Catch ex As Exception
            Error_Logger(ex, "MapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Public Function UnMapDrive(ByVal pathtomap As String) As String
        Dim resultdrive As String = ""
        Try
            Dim apptorun As String
            'net use T: \\Comlab\Vol2\handin /PERSISTENT:NO
            'net use T: /DELETE

            If pathtomap.EndsWith("\") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            If pathtomap.EndsWith(":") = True Then
                pathtomap = pathtomap.Remove(pathtomap.Length - 1, 1)
            End If
            apptorun = "net use " & pathtomap & ":  /DELETE"

            resultdrive = DosShellCommand(apptorun)
        Catch ex As Exception
            Error_Logger(ex, "UnMapDrive")
            resultdrive = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return resultdrive
    End Function

    Private Sub Error_Logger(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then

                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                If identifier_msg = "" Then
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & ex.ToString)
                Else
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                End If

                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin System's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Function UserGroups(ByVal inusername As String, ByVal incontext As String) As String
        Dim result As String
        Try
            Dim apptorun As String
            Dim finfo As FileInfo = New FileInfo((Application.StartupPath & "\UCT Novell Users Group Extractor.exe").Replace("\\", "\"))
            If finfo.Exists = True Then
                apptorun = """" & (Application.StartupPath & "\UCT Novell Users Group Extractor.exe").Replace("\\", "\") & """ """ & inusername & """ """ & incontext & """ """ & error_level & """"

                result = ApplicationLauncher(apptorun)
            Else
                result = "Fail. Group Retrieval Script Executable cannot be found"
            End If
            finfo = Nothing
        Catch ex As Exception
            Error_Logger(ex, "UserGroups")
            result = "Fail. Unknown Reason. Check the log files for further information."
        End Try
        Return result
    End Function
End Class
