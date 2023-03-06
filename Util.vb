Module Util
    '==== Wait ================================================================
    Public Sub wait(ByVal interval As Integer)
        xtrace_i("Wait " & interval.ToString)
        interval = interval * 1000

        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval

        Loop
        sw.Stop()
    End Sub

    '==== Press Any Key =======================================================
    Public Sub AnyKey()
        xtrace_subs("AnyKey")
        Console.Write("Press any key to continiue . .")
        Console.ReadKey()
        Console.WriteLine()
        xtrace_sube("AnyKey")
    End Sub
End Module
