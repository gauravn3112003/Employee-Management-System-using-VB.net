Imports System.Data.OleDb
Public Class Form1
    Dim con2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Asus\Documents\EMS.accdb")
    'Public Sub addData(ByVal name As String, ByVal department As String, ByVal dateOfJoining As String, ByVal salary As String, ByVal age As String, ByVal phone As String)
    Public Sub addData(ByVal name As String)

        Try
            con2.Open()
            'Dim cmd1 As New OleDbCommand("INSERT INTO employee([Name],[Age],[Department],[DateOfJoining],[Salary],[Phone]) VALUES('" & name & "','" & age & "','" & department & "','" & dateOfJoining & "','" & salary & "','" & phone & "')", con2)

            Dim cmd1 As New OleDbCommand("INSERT INTO ems([Name]) VALUES('" & name & "')", con2)
            cmd1.ExecuteNonQuery()
            cmd1.Dispose()
            con2.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Name, Department, DateOfJoining, Salary, Age, Phone As String
        Name = TextBox1.Text
        Age = TextBox2.Text
        Department = TextBox3.Text
        DateOfJoining = TextBox4.Text
        Salary = "₹ " & TextBox5.Text
        Phone = TextBox6.Text
        'addData(name, Department, DateOfJoining, Salary, Age, Phone)
        addData(Name)
    End Sub
End Class
