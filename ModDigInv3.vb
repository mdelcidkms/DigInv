Option Strict Off
Option Explicit On
Imports System.Data.OracleClient
'Imports VB = Microsoft.VisualBasic
Module ModDigInv3
    Public Se As Object
    Public Db As Object
    Public conn As OracleConnection

    Public gs_UserName As String
    Public gs_ReqUser As String
    Public gs_Password As String
    Public gs_DSN As String
    Public gs_BranchN As String
    Public gs_BranchName As String
    Public gs_Hoststring As String
    Public gs_SessionID As String
    Public gstrBranchAlpha As String
    Public ShipDate As String
    Public RunDateTime As String
    Public strTab As String
    Public StrTab2 As String 'Second temp table
    Public DBOpened As Boolean

    Public glngParNum As Integer ' Commandline Parameter Holder
    Public gstrParName() As String
    Public gstrParValue() As String


    '=========================
    Public Sub OpenDatabase()
        '=========================
        Init(Command())
        gs_UserName = GetParam("DBUSER")
        gs_ReqUser = GetParam("USERNAME")
        gs_Password = GetParam("DBPASSWD")
        gs_Hoststring = GetParam("TNSNAME")
        gs_BranchN = GetParam("BRANCHID")
        'Se = CreateObject("Oracleinprocserver.xorasession")'04/24/2024 Change for vb.net upgrade
        'Db = Se.OpenDatabase(GetParam("TNSNAME"), GetParam("DBUSER") & "/" & GetParam("DBPASSWD"), 0) '04/24/2024 Change for vb.net upgrade
        Dim connString As String = "Data Source=" + GetParam("TNSNAME") + ";User Id=" + GetParam("DBUSER") + ";Password=" + GetParam("DBPASSWD") + ";" '04/23/2024 Change for vb.net upgrade
        conn = New OracleConnection(connString) '04/24/2024 Change for vb.net upgrade

        conn.Open()
        DBOpened = True


    End Sub

    ''==========================
    Public Sub GetBranchName() '04/24/2024 add for vb.net upgrade
        '==========================
        Dim SQL As String
        Dim OCM As OracleCommand
        Dim ODR As OracleDataReader
        Dim DT As DataTable

        SQL = "select branchname, branchalpha from branch where branchid = '" & gs_BranchN & "'"
        OCM = New OracleCommand(SQL, conn)
        ODR = OCM.ExecuteReader()
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)
        If ODR.HasRows Then
            While ODR.Read()
                gs_BranchName = "" & ODR("branchname").ToString()
                gstrBranchAlpha = "" & ODR("branchalpha").ToString()
            End While
        End If
    End Sub



    ''==========================
    'Public Sub GetBranchName() '04/24/2024 replaced for vb.net upgrade
    '    '==========================
    '    Dim RS As Object
    '    Dim SQL As String
    '    SQL = "select branchname, branchalpha from branch where branchid = '" & gs_BranchN & "'"
    '    RS = Db.CreateDynaset(SQL, &H4)
    '    If RS.RecordCount > 0 Then
    '        gs_BranchName = "" & RS!branchname.Value
    '        gstrBranchAlpha = "" & RS!branchalpha.Value
    '    End If
    '    RS.Close()
    'End Sub
    ''==========================


    Public Sub Init(ByVal strCmdLine As String)
        Dim c, Value As String
        Dim pos As Integer

        glngParNum = 0
        strCmdLine = Trim(strCmdLine)
        While strCmdLine <> ""
            ReDim Preserve gstrParName(glngParNum)
            ReDim Preserve gstrParValue(glngParNum)

            'Get Parameter Name
            pos = InStr(strCmdLine, "=")
            If pos < 2 Then GoTo ERRH
            gstrParName(glngParNum) = Mid(strCmdLine, 1, pos - 1)
            strCmdLine = Mid(strCmdLine, pos + 1)

            'Get Parameter Value
            If Left(strCmdLine, 1) <> """" Then GoTo ERRH
            strCmdLine = Mid(strCmdLine, 2)
            Value = ""
            Do
                c = Mid(strCmdLine, 1, 1) : strCmdLine = Mid(strCmdLine, 2)
                If c = "" Then GoTo ERRH
                If c = """" Then
                    If Mid(strCmdLine, 1) = """" Then
                        strCmdLine = Mid(strCmdLine, 2)
                    Else
                        gstrParValue(glngParNum) = Value
                        strCmdLine = Trim(Mid(strCmdLine, 2)) 'Skip ending double quote
                        Exit Do
                    End If
                End If
                Value = Value & c
            Loop

            glngParNum = glngParNum + 1
        End While

        Exit Sub
ERRH:
        MsgBox("CommandLine Error (JETSLib.Init)", MsgBoxStyle.Critical)
    End Sub


    Public Function GetParam(ByRef strParName As String) As String
        Dim i As Integer

        For i = 0 To glngParNum - 1
            If gstrParName(i) = strParName Then
                GetParam = gstrParValue(i)
                Exit Function
            End If
        Next

        MsgBox("GetParam Error (JETSLib): " & strParName, MsgBoxStyle.Critical)
    End Function
End Module