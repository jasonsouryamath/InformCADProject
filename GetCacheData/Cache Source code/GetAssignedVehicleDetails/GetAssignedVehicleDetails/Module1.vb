Module Module1

    Sub Main(args As String())
        On Error Resume Next
        Dim oFactory, gOleServer, fs, a, ServerRespId1, ai
        Dim obj0, obj1
        Dim number




        oFactory = CreateObject("MSOS.clsObjectFactory")
        gOleServer = oFactory.CreateCADServer()
        gOleServer.GetAllActiveResponses1()
        ServerRespId1 = gOleServer.GetActiveResponseVehicles(CInt(args(0)), Vehicle_ID:=obj0, Radio_Name:=obj1)
        fs = CreateObject("Scripting.FileSystemObject")
        a = fs.CreateTextFile(Environment.CurrentDirectory() + "\" + args(0) + ".txt")
        Dim stringce As String
        For ai = 0 To (ServerRespId1)
            stringce = obj1(ai).ToString() + "|" + obj0(ai).ToString()
            Console.Write(stringce)
            a.write(stringce)

        Next

        a.close()
        System.Environment.Exit(0)

        Exit Sub
    End Sub

End Module
