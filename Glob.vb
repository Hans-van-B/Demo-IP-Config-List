Module Glob
    Public AppName As String = "IP-Config-List"
    Public AppVer As String = "0.01"

    Public AppRoot As String = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)
    'Public AppRoot As String = My.Application.Info.DirectoryPath
    Public CD As String = My.Computer.FileSystem.CurrentDirectory

    ' Log
    Public Temp As String = Environment.GetEnvironmentVariable("TEMP")
    Public LogFile As String = Temp & "\" & AppName & ".log"
    Public Verbose As Boolean = False
    Public CTrace As Integer = 1                       ' Console Trace Level
    Public LTrace As Integer = 2
    Public G_TL_Sub As Integer = 2  ' Trace level for Subroutines
    Public G_TL_FR As Integer = 2   ' Trace level File Reads
    Public SubLevel As Integer = 0
    Public ErrorCount As Integer = 0
    Public ExitProgram As Boolean = False
    Public StartupDir As String = "$StartupDir$"
    Public Debug As Boolean = False

    'Defaults
    Public WaitForKey As Boolean = False
End Module

' If you get the error: Computer is not a member of 'NetConnect.My'
' Then you selected the wrong template. Use "Console App (.NET Framework)"
'                                       Not "Console App (.Net Core)"
