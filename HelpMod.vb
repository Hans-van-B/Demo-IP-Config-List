Module HelpMod
    Sub ShowHelp()
        xtrace("Print Help")

        Console.WriteLine(" ")
        Console.WriteLine("---- " & AppName & " V" & AppVer & " -----------------------------------------")
        Console.WriteLine(" Switches:")
        Console.WriteLine("  -d = Debug on")
        Console.WriteLine("  -h = Show this help page")
        Console.WriteLine("  /? = Show this help page")
        Console.WriteLine("  -v = Verbose")
        Console.WriteLine("  -w = Wait for key before exiting")
        Console.WriteLine(" ")
        Console.WriteLine("  --StartupDir = Change the startup directory")
        Console.WriteLine("  --help       = The same as -h")
        Console.WriteLine("  --xhelp      = extended help")
        Console.WriteLine(" ")
        Console.WriteLine(" Arguments:")
        Console.WriteLine(" ")
        Console.WriteLine(" Examples: ")
        Console.WriteLine(" ")
        Console.WriteLine(" Log file = " & LogFile)
        Console.WriteLine(" Hans.van.Buitenen@philips.com")
        Console.WriteLine(" ")
        Console.WriteLine("----------------------------------------------------------------")
        Console.WriteLine(" ")

        wait(2)
    End Sub

    Sub ShowXHelp()
        xtrace("Print Help")

        Console.WriteLine(" ")
        Console.WriteLine("---- " & AppName & " V" & AppVer & " -----------------------------------------")
        Console.WriteLine(" Switches:")
        Console.WriteLine("  -d = Debug on")
        Console.WriteLine("  -h = Show this help page")
        Console.WriteLine("  /? = Show this help page")
        Console.WriteLine("  -v = Verbose")
        Console.WriteLine("  -w = Wait for key before exiting")
        Console.WriteLine(" ")
        Console.WriteLine("  --StartupDir = Change the startup directory")
        Console.WriteLine("  --help       = The same as -h")
        Console.WriteLine("  --xhelp      = extended help")
        Console.WriteLine(" ")
        Console.WriteLine(" Arguments:")
        Console.WriteLine(" ")
        Console.WriteLine(" Log file = " & LogFile)
        Console.WriteLine(" ")
        Console.WriteLine(" Examples: ")
        Console.WriteLine(" ")
        Console.WriteLine(" Log file = " & LogFile)
        Console.WriteLine(" Dev. : C:\Users\nly10677\OneDrive - Philips\VS_Projects\??")
        Console.WriteLine(" Maint: P:\Dev\")
        Console.WriteLine("        https://github.com/Hans-van-B?tab=repositories ")
        Console.WriteLine(" Inst.: Depo\")
        Console.WriteLine(" Hans.van.Buitenen@philips.com")
        Console.WriteLine(" ")
        Console.WriteLine("----------------------------------------------------------------")
        Console.WriteLine(" ")

        wait(2)
    End Sub
End Module
