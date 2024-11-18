Module MainModule
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    Sub Main()

        If My.Application.CommandLineArgs.Count > 0 Then
            Dim action As String = My.Application.CommandLineArgs(0)

            Select Case action
                Case "-auto"
                    Console.WriteLine("Aplikasi memulai...")
                    Dim form1 As New Form1()
                    Application.Run(form1)
                Case Else
                    MsgBox("Tidak bisa dijalankan langsung")
                    Console.WriteLine("Perintah tidak dikenal")
            End Select
        Else
            Interaction.MsgBox("Tidak bisa dijalankan langsung!")
            Console.WriteLine("Tidak ada argument yang diberikan")

        End If

        Console.ReadLine()
    End Sub
    Public Sub SetIniSettings(ByVal strINIFile As String, ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)
        Try
            WritePrivateProfileString(strSection, strKey, strValue, strINIFile)
        Catch ex As Exception
            If Err.Number <> 0 Then Err.Raise(Err.Number, , "Error form Functions.SetIniSettings " & Err.Description)
        End Try
    End Sub
    Public Function GetIniSetting(ByVal strINIFile As String, ByVal strSection As String, ByVal strKey As String) As String
        Dim strValue As String = ""
        Try
            strValue = Space(1024)
            GetPrivateProfileString(strSection, strKey, "NOT_FOUND", strValue, 1024, strINIFile)
            Do While InStrRev(strValue, " ") = Len(strValue)
                strValue = Mid(strValue, 1, Len(strValue) - 1)
            Loop
            strValue = Mid(strValue, 1, Len(strValue) - 1)
            GetIniSetting = strValue

        Catch ex As Exception
            If Err.Number <> 0 Then Err.Raise(Err.Number, , "Error form Functions.SetIniSettings " & Err.Description)
        End Try
        Return strValue
    End Function
End Module
