Module Module1

    Sub Main()
        Dim oFactory, gOleServer, fs, a, ServerRespId1, ai
        Dim obj0 As Object
        On Error Resume Next
        oFactory = CreateObject("MSOS.clsObjectFactory")
        gOleServer = oFactory.CreateCADServer()
        gOleServer.GetAllActiveResponses1()
        ServerRespId1 = gOleServer.GetAllActiveResponses1(obj0)
        fs = CreateObject("Scripting.FileSystemObject")
        a = fs.CreateTextFile(Environment.CurrentDirectory() + "\" + "IncidentList.txt", True)
        Dim stringce As String
        For ai = 0 To (ServerRespId1)
            stringce = obj0(ai).ToString() + "|"
            Console.Write(stringce)
            a.write(stringce)

        Next

        a.close()
        System.Environment.Exit(0)

        Exit Sub
    End Sub

End Module
