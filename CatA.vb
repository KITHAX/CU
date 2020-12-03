Imports System.Data.OleDb
Public Class CatA

    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim myConnection As OleDbConnection = New OleDbConnection
    Private Sub CatA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'DATOSDataSet.INFORMACION' Puede moverla o quitarla según sea necesario.
        Me.INFORMACIONTableAdapter.Fill(Me.DATOSDataSet.INFORMACION)

    End Sub

    Private Sub Alta_Click(sender As Object, e As EventArgs) Handles Alta.Click
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
        dataFile = "G:\PROYECTO 3.0\ACCESS\DATOS.accdb;"

        connString = provider & dataFile
        myConnection.ConnectionString = connString

        myConnection.Open()

        Dim str As String
        str = "INSERT INTO INFORMACION ([FECHA_DE_INGRESO],[MARCAS],[MODELO],[COLOR],[AÑO],[KILOMETRAJE],[VERSION]) VALUES (@Calen,@LMarc,@Model,@TColor,@LAño,@Kilom,@Versi) "


        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("FECHA_DE_INGRESO", CType(Calen.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("MARCAS", CType(LMarc.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("MODELO", CType(Model.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("COLOR", CType(TColor.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("AÑO", CType(LAño.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("KILOMETRAJE", CType(Kilom.Text, String)))
        cmd.Parameters.Add(New OleDbParameter("VERSION", CType(Versi.Text, String)))

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
            ' Alta.Clear()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Baja_Click(sender As Object, e As EventArgs) Handles Baja.Click

    End Sub

End Class
