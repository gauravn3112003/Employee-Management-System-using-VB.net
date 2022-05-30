Imports System.Data.OleDb
Public Class Form1
    Dim con2 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Employee Management\Database11.accdb")
    Public Sub addData(ByVal name As String, ByVal department As String, ByVal dateOfJoining As String, ByVal salary As String, ByVal age As String, ByVal phone As String)
        'Public Sub addData(ByVal name As String)
        Try
            con2.Open()
            Dim cmd1 As New OleDbCommand("INSERT INTO employee([Name],[Age],[Department],[DateOfJoining],[Salary],[Phone]) VALUES('" & name & "','" & age & "','" & department & "','" & dateOfJoining & "','" & salary & "','" & phone & "')", con2)
            ' Dim cmd1 As New OleDbCommand("INSERT INTO Employee([Name]) VALUES('" & name & "')", con2)
            cmd1.ExecuteNonQuery()
            cmd1.Dispose()
            MsgBox("Record Inserted Successfully !")
            con2.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            con2.Close()
        End Try
    End Sub
    Public Sub showData()
        Try
            con2.Close()
            Dim cmd As New OleDbCommand("SELECT * FROM EMPLOYEE", con2)
            Dim da As New OleDbDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView1.DataSource = dt
            con2.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            con2.Close()

        End Try
    End Sub
    Public Sub updateData(ByVal name As String, ByVal department As String, ByVal dateOfJoining As String, ByVal salary As String, ByVal age As String, ByVal phone As String, ByVal key As String)

        Try
            con2.Open()
            Dim cmd2 As New OleDbCommand("SELECT * FROM Employee where Name = '" & name & "'", con2)
            Dim cmd3 As New OleDbCommand("UPDATE EMPLOYEE SET Name='" & name & "',Age='" & age & "',Department='" & department & "',DateOfJoining='" & dateOfJoining & "',Salary='" & salary & "',Phone='" & phone & "' WHERE Name='" & key & "'", con2)

                cmd3.ExecuteNonQuery()
                con2.Close()
                MsgBox("Record Update Successfully !")

            con2.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            con2.Close()
        End Try
    End Sub
    Public Sub deleteData(ByVal name As String)
        Try
            con2.Open()

            Dim cmd As New OleDbCommand("DELETE FROM EMPLOYEE WHERE Name = '" & name & "'", con2)
            cmd.ExecuteNonQuery()
            MsgBox("Record Deleted Successfully !")
            con2.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            con2.Close()

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        showData()
        Me.TabControl1.SelectedTab = Me.TabPage2
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        Dim Name, Department, DateOfJoining, Salary, Age, Phone As String
        Name = TextBox12.Text
        Age = TextBox11.Text
        Department = TextBox10.Text
        DateOfJoining = TextBox9.Text
        Salary = "₹ " & TextBox8.Text
        Phone = TextBox7.Text
        addData(Name, Department, DateOfJoining, Salary, Age, Phone)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Me.TabControl1.SelectedTab = Me.TabPage3
        Timer1.Stop()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        showData()
        Me.TabControl1.SelectedTab = Me.TabPage2


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Name, Department, DateOfJoining, Salary, Age, Phone, Key As String
        Name = TextBox6.Text
        Age = TextBox5.Text
        Department = TextBox4.Text
        DateOfJoining = TextBox3.Text
        Salary = "₹ " & TextBox2.Text
        Phone = TextBox1.Text
        Key = TextBox13.Text
        updateData(Name, Department, DateOfJoining, Salary, Age, Phone, Key)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.TabControl1.SelectedTab = Me.TabPage4
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        showData()
        Me.TabControl1.SelectedTab = Me.TabPage2
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.TabControl1.SelectedTab = Me.TabPage5
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim name As String
        name = TextBox20.Text
        deleteData(name)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        showData()
        Me.TabControl1.SelectedTab = Me.TabPage2
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.TabControl1.SelectedTab = Me.TabPage3
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.TabControl1.SelectedTab = Me.TabPage3

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.TabControl1.SelectedTab = Me.TabPage3

    End Sub
End Class
