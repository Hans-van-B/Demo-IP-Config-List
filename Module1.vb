Module Module1

    Sub Main()
        xtrace_init()
        Read_Command_Line_Arg()
        If ExitProgram Then Exit Sub

        xtrace("Demo: IP Config List", 1)
        GetNetworkInformation()
        xtrace("", 1)

        xtrace_TimeStamp()
        If WaitForKey Then AnyKey()
    End Sub

End Module
