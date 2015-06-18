Imports System
Imports System.Configuration
Imports System.Configuration.ConfigurationManager
Imports System.Data
Imports System.Management
Imports System.Data.OleDb
Imports System.Data.OleDb.OleDbException
Imports System.Data.SqlClient
Imports System.Net.Sockets
Imports System.IO
Imports Microsoft.Win32




Module Module1

    'Public testes = testreg("Keyboard", "InitialKeyboardIndicators")
    Public test = folder_exists("C:\windows")

    Public intEid2 = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "EID", Nothing)

    Public intEid As Integer
    Public strOS As String = My.Computer.Info.OSFullName
    Public strEname As String = My.Computer.Name
    Public strGuid As String = GetWMIProperties("win32_OperatingSystem", "SerialNumber")
    Public strModel As String = GetWMIProperties("win32_computersystem", "model")
    Public strManufacturer As String = GetWMIProperties("win32_computersystem", "manufacturer")
    Sub Main()


        ' If lineinfileexists("D:\Users\bourquer\Desktop\autoexec.bat", "test") Then MsgBox("yes")
        'Call GetTasks(2)
        On Error Resume Next
        Dim intErrorLevel As Integer
        'Dim strPath As String = "D:\Users\bourquer\Desktop\Untitled3.ps1"
        'Shell("powershell.exe -file " & strPath)
        'Check Registry for CID
        If IsNothing(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "EID", Nothing)) Or (My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "GUID", Nothing) <> strGuid) Then

            'If EID/ProductID combo match not found, write ename to db and get back cid - write it to registry
            intEid = GetEID(strEname, strOS, strGuid)

            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "EID", intEid)
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "GUID", strGuid)

        Else

            'Found EID/ProductID assign variable - check to see if there are any updated properties - check to see if any tasks are due
            intEid = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\EasyAdm", "EID", Nothing)

            intErrorLevel = SetProperty("Manufacturer", strManufacturer)
            intErrorLevel = SetProperty("Model", strModel)

        End If


    End Sub


    Public Function RunPowershellViaShell(ByVal scriptText As String) As Integer
        Dim execProcess As New System.Diagnostics.Process
        Dim psScriptTextArg = "-NoExit -Command ""& get-module -list"""
        'Dim psScriptTextArg = "-NoExit -Command ""& set-executionPolicy unrestricted"""
        'Dim psScriptTextArg = ""-NoExit -Command """ & scriptText & """"

        execProcess.StartInfo.WorkingDirectory = Environment.SystemDirectory & "\WindowsPowershell\v1.0\"
        execProcess.StartInfo.FileName = "powershell.exe"
        execProcess.StartInfo.Arguments = psScriptTextArg
        execProcess.StartInfo.UseShellExecute = True
        Return execProcess.Start
    End Function

    Private Sub GetInfo()
        Debug.Print(MsgBox(GetWMIProperties("win32_computersystem", "domain")))
        Debug.Print(GetWMIProperties("win32_computersystem", "domainrole"))
        Debug.Print(GetWMIProperties("win32_computersystem", "systemtype"))
        Debug.Print(GetWMIProperties("win32_computersystem", "manufacturer"))
        Debug.Print(GetWMIProperties("win32_computersystem", "model"))
        Debug.Print(GetWMIProperties("win32_computersystem", "numberofprocessors"))
        Debug.Print(GetWMIProperties("win32_computersystem", "totalphysicalmemory"))
    End Sub
    Private Function lineinfileexists(strFile As String, strtofind As String) As Boolean
        Try
            Using sr As New StreamReader(strFile)
                Dim line As String
                line = sr.ReadToEnd()
                If InStr(line, strtofind) Then
                    Return True
                Else
                    Return False
                End If

            End Using
        Catch e As Exception
            Return False
        End Try

    End Function

    Private Function GetWMIProperties(ByVal ClassName As String, ByVal Selection As String) As String
        Dim tmpstr As String

        Dim myScope As New ManagementScope("\\" & strEname & "\root\cimv2")
        Dim oQuery As New SelectQuery("SELECT " & Selection & " FROM " & ClassName)
        Dim oResults As New ManagementObjectSearcher(myScope, oQuery)
        Dim oItem As ManagementObject
        Dim oProperty As PropertyData

        For Each oItem In oResults.Get()
            For Each oProperty In oItem.Properties
                Try
                    tmpstr = oProperty.Value.ToString
                    Return tmpstr
                Catch ex As Exception
                    tmpstr = ""
                End Try
            Next
        Next

        tmpstr = Nothing
        myScope = Nothing
        oQuery = Nothing
        oResults = Nothing
        oItem = Nothing
        oProperty = Nothing
        Return tmpstr


    End Function
    Public Function file_exists(strFile As String) As Boolean
        Return (If(File.Exists(strFile), True, False))
    End Function
    Public Function folder_exists(strFolder As String) As Boolean
        Return (If(Directory.Exists(strFolder), True, False))
    End Function
    Public Function test3(strPath As String)
        Dim Files = Directory.GetFiles(strPath)
        If (files.Length > 0) Then

        End If
    End Function

    Private Function GetTasks(ByVal strEname As Integer) As String

        Dim settings As ConnectionStringSettings
        settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
        Dim cnn As New SqlConnection(settings.ConnectionString)
        Dim cmd As New SqlCommand
        Dim Return_Val As SqlParameter

        cmd.Connection = cnn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "uspGetTasks"
        cmd.Parameters.AddWithValue("@eid", strEname)
        Return_Val = cmd.Parameters.Add("Return_Val", SqlDbType.NChar)
        Return_Val.Direction = ParameterDirection.ReturnValue

        Try
            cnn.Open()
            cmd.ExecuteScalar()
            Return Return_Val.Value.ToString
        Catch ex As Exception
        End Try

    End Function
    Private Function GetEID(ByVal strEname As String, ByVal strOS As String, ByVal strGuid As String) As Integer

        Dim settings As ConnectionStringSettings
        settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
        Dim cnn As New SqlConnection(settings.ConnectionString)
        Dim cmd As New SqlCommand
        Dim Return_Val As SqlParameter

        cmd.Connection = cnn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "uspGetEID"
        cmd.Parameters.AddWithValue("@ename", strEname)
        cmd.Parameters.AddWithValue("@eos", strOS)
        cmd.Parameters.AddWithValue("@guid", strGuid)
        Return_Val = cmd.Parameters.Add("Return_Val", SqlDbType.Int)
        Return_Val.Direction = ParameterDirection.ReturnValue

        Try
            cnn.Open()
            cmd.ExecuteScalar()
            Return Return_Val.Value
        Catch ex As Exception
        End Try

    End Function

    Private Function testreg(ByVal skey As String, ByVal sval As String)
        Dim test As Microsoft.Win32.RegistryView = Microsoft.Win32.RegistryHive.Users
        Dim RK As Microsoft.Win32.RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(test, Microsoft.Win32.RegistryView.Registry64). _
        OpenSubKey("S-1-5-18\\control panel\\" & skey)
        Dim strRegValue = RK.GetValue(sval)
        Return strRegValue
    End Function
    Private Function ReadProperties(ByVal strCommand As String) As List(Of String)


        Dim settings As ConnectionStringSettings
        settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
        If Not settings Is Nothing Then

            Dim sqlConnection1 As New SqlConnection(settings.ConnectionString)
            Dim sqlcmd As New SqlCommand(strCommand)
            sqlcmd.Connection = sqlConnection1
            sqlConnection1.Open()
            Dim sqldr As SqlDataReader = sqlcmd.ExecuteReader()
            Dim arrResults As New List(Of String)
            Dim arrResults2 As New List(Of String)
            If sqldr.Read() Then
                While sqldr.Read()

                    For count As Integer = 0 To (sqldr.FieldCount - 1)

                        arrResults.Add(sqldr(count))

                    Next
                    arrResults2 = arrResults
                    arrResults.Clear()
                End While
            Else
                Return Nothing
            End If
            Return arrResults
            sqldr.Close()


        End If

    End Function
    Private Function SetProperty(ByVal strPropDesc As String, ByVal strPropVal As String) As Integer
        'examples of strPropdesc and strPropval are "Name"/"Local Area Connection", "Type"/"Network Connection", "Name"/"Microsoft Visio", and "Version"/"10.6.8"
        Dim settings As ConnectionStringSettings
        settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
        Dim cnn As New SqlConnection(settings.ConnectionString)
        Dim cmd As New SqlCommand
        Dim Return_Val As SqlParameter

        cmd.Connection = cnn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "uspProperty"
        cmd.Parameters.AddWithValue("@propdescription", strPropDesc)
        cmd.Parameters.AddWithValue("@propvalue", strPropVal)
        Return_Val = cmd.Parameters.Add("Return_Val", SqlDbType.Int)
        Return_Val.Direction = ParameterDirection.ReturnValue

        Try
            cnn.Open()
            cmd.ExecuteScalar()
            Return Return_Val.Value
        Catch ex As Exception
        End Try

    End Function
    Private Function SetEndpointProperties(ByVal strPropDesc As String, ByVal strPropVal As String, ByVal strCatDesc As String, ByVal strHeadDesc As String, Optional ByVal intECatPropInterval As Integer = 0, Optional ByVal intECatID As Integer = 0) As Integer
        'examples of strPropdesc and strPropval are "Name"/"Local Area Connection", "Type"/"Network Connection", "Name"/"Microsoft Visio", and "Version"/"10.6.8"
        'examples of strCatDesc are "Hardware Information" and "Installed Applications"
        'intECatPropInterval can be provided to set a specific number of hours to update the property - default is 0 (do not update)
        'intECatID can be provided to add to an existing EndpointCategory e.g. each Installed Application would have several properties and values so the first set (Name/Microsoft Visio) would return the intCatID which can be used to add additional sets (Version/10.6.8)
        Dim settings As ConnectionStringSettings
        settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
        Dim cnn As New SqlConnection(settings.ConnectionString)
        Dim cmd As New SqlCommand
        Dim Return_Val As SqlParameter

        cmd.Connection = cnn
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "uspEnterEndpointProperty"
        cmd.Parameters.AddWithValue("@propdescription", strPropDesc)
        cmd.Parameters.AddWithValue("@propvalue", strPropVal)
        cmd.Parameters.AddWithValue("@catdescription", strCatDesc)
        cmd.Parameters.AddWithValue("@headdescription", strHeadDesc)
        cmd.Parameters.AddWithValue("@ecatpropInterval", intECatPropInterval)
        cmd.Parameters.AddWithValue("@ecatid", intECatID)
        cmd.Parameters.AddWithValue("@eid", intEid)
        Return_Val = cmd.Parameters.Add("Return_Val", SqlDbType.Int)
        Return_Val.Direction = ParameterDirection.ReturnValue

        Try
            cnn.Open()
            cmd.ExecuteScalar()
            Return Return_Val.Value
        Catch ex As Exception
        End Try

    End Function


    'Private Function FindProperty(ByVal strCommand As String) As Integer


    '    Dim settings As ConnectionStringSettings
    '    settings = ConfigurationManager.ConnectionStrings("EasyAdm.My.MySettings.Setting")
    '    If Not settings Is Nothing Then

    '        Dim sqlConnection1 As New SqlConnection(settings.ConnectionString)
    '        Dim sqlcmd As New SqlCommand(strCommand)
    '        sqlcmd.Connection = sqlConnection1
    '        sqlConnection1.Open()
    '        Dim sqldr As SqlDataReader = sqlcmd.ExecuteReader()
    '        If sqldr.Read() Then
    '            ' While sqldr.Read()

    '            Return sqldr("PID")

    '            'End While
    '        Else
    '            Return 0
    '        End If
    '        sqldr.Close()


    '    End If

    'End Function

End Module
