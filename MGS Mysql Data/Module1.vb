Imports Microsoft.Win32

Module Module1
    Public Wrapp_My As New Simple3Des("jklghsdoeri")
    Public Conn_str As String = ""
    Public Gl_Pass As String
    Public Function My_Save_Setting(ByVal mgankof As String, ByVal myKey As String, ByVal mValue As String) As Boolean
        Return My_Save_Setting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, mgankof, myKey, mValue)
    End Function

    Public Function My_Save_Setting(ByVal myKey As String, ByVal mValue As String) As Boolean
        Return My_Save_Setting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "Admin", myKey, mValue)
    End Function

    Public Function My_Save_Setting(ByVal mseqcia As String, ByVal mgankof As String, ByVal myKey As String, ByVal mValue As String) As Boolean
        If mseqcia = "" Then
            mseqcia = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name
        End If
        If mgankof = "" Then
            mgankof = "Admin"
        End If
        Try
            Dim CU As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\MGS\" & mseqcia & "\" & mgankof)
            With CU
                .OpenSubKey("SOFTWARE\MGS\" & mseqcia & "\" & mgankof, True)
                .SetValue(myKey, mValue)
            End With
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function My_Get_Setting(ByVal mgankof As String, ByVal myKey As String, ByVal mValue As String) As String
        Return My_Get_Setting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, mgankof, myKey, mValue)
    End Function

    Public Function My_Get_Setting(ByVal myKey As String, ByVal mValue As String) As String
        Return My_Get_Setting(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "Admin", myKey, mValue)
    End Function

    Public Function My_Get_Setting(ByVal mseqcia As String, ByVal mgankof As String, ByVal myKey As String, ByVal mValue As String) As String
        My_Get_Setting = mValue
        If mseqcia.Length = 0 Then
            mseqcia = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name
        End If
        If mgankof.Length = 0 Then
            mgankof = "Admin"
        End If
        Try
            Dim CU As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\MGS\" & mseqcia & "\" & mgankof)
            With CU
                .OpenSubKey("SOFTWARE\MGS\" & mseqcia & "\" & mgankof, True)
                My_Get_Setting = .GetValue(myKey, mValue)
            End With
        Catch
        End Try
    End Function
End Module
