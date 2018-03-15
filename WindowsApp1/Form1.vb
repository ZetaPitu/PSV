Imports System.IO
Imports System.Windows.Markup
Imports MS.Internal

Public Class Form1
    Private contador As Int32 = 0
    Private contador1 As Int32 = 0
    Private contador2 As Drawing.Size
    Private Errores As Int32 = 0
    Private Final As Int32
    Private myStream As Stream = Nothing
    Private DirPsv = "http://sophrologie-dynamique.weebly.com/uploads/1/1/3/2/11325139/1334317102.jpg "
    Private img = "1.jpg"
    Private DirImg = Nothing
    Private ArchivoWeb(2000) As String
    Private ArchivoPc(2000) As String
    Private Rapido As Boolean = True


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Rapido = True
        Try

            ReadFile()

        Catch Ex As Exception

            MessageBox.Show(Ex.Message)

        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        myStream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()

        REM openFileDialog1.InitialDirectory = "Y:\"
        openFileDialog1.Filter = "Excel separado por ; (*.csv) |*.csv|Todos los archivos|*.*"
        openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True
        TextBox2.Text = openFileDialog1.FileName
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try

                TextBox2.Text = openFileDialog1.FileName
                Button2.Enabled = False
                Button3.Enabled = True

                REM  Leer_Archivo_Excel()

            Catch Ex As Exception
                MessageBox.Show("Error al abrir archivo: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        REM FolderBrowserDialog1.SelectedPath = "Y:\"

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
            Button3.Enabled = False
            Button1.Enabled = True
            Button5.Enabled = True
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = " Archivo de excel"
        Label2.Text = " Directorio de destino"
        FolderBrowserDialog1.ShowDialog()
    End Sub
    Function Leer_Archivo_Excel() As Boolean



        ReadFile()

        REM MessageBox.Show(TextBox3.Text)
        Return True
    End Function


    Function Leer_Archivo(Web As String, Local As String) As Boolean
        Try
            TextBox3.Text = Web
            REM   My.Computer.Network.DownloadFile(Web, Local, "anonymous", "")
            ArchivoWeb(contador) = New String(Web)
            ArchivoPc(contador) = Local
            contador = contador + 1
            If contador > 1000 Then
                REM   MsgBox(Web)
            End If
            TextBox4.Text = TextBox4.Text + " : " + Web + vbCrLf
        Catch ex As Exception
            MsgBox("Error " & ex.Message & " Error de escritura ")
        End Try
        Return True
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Timer1.Enabled = False

        Dim response = MsgBox("Finalizar", MsgBoxStyle.YesNo, "In Live")
        If response = MsgBoxResult.Yes Then
            Application.Exit()
        End If

        Timer1.Enabled = True

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub ReadFile()
        Label3.Text = "Leyendo logos a ......Espere"
        Button5.Enabled = False
        Button1.Enabled = False
        Using MyReader As New Microsoft.VisualBasic.
            FileIO.TextFieldParser(TextBox2.Text, System.Text.Encoding.Unicode)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";; ")
            Dim currentRow As String()
            Try
                currentRow = MyReader.ReadFields()
            Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                MsgBox("Line " & ex.Message &
                    " Error en linea ")
            End Try
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    Dim currentField As String
                    For Each currentField In currentRow
                        My.Application.DoEvents()
                        If (currentField.Length > 3) Then
                            TextBox3.Text = currentField
                            Label3.Text = "Leyendo logos a ...... " + CType(contador, String)
                            Crear_Archivo(TextBox3.Text)
                            Try
                                ProgressBar1.Maximum = contador + 1
                                ProgressBar1.Value = contador
                            Catch ex As Exception
                                ProgressBar1.Value = 0
                            End Try


                        End If

                    Next
                Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                    "Error en linea")
                End Try

            End While

            Descargar_Logos()
        End Using
    End Sub
    Private Sub ReadLine()
        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(TextBox2.Text)

            MyReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            MyReader.SetDelimiters(";;")
            Dim currentRow As String()
            'Loop through all of the fields in the file. 
            'If any lines are corrupt, report an error and continue parsing. 
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    MsgBox(currentRow)
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                    " is invalid.  Skipping")
                End Try
            End While
        End Using
    End Sub

    Private Sub Crear_Archivo(S As String)

        Dim TestString As String = S
        Dim TestArray() As String = Split(TestString, ";")
        ' TestArray holds {"apple", "", "", "", "pear", "banana", "", ""}
        Dim LastNonEmpty As Integer = -1
        For i As Integer = 0 To TestArray.Length - 1
            If TestArray(i) <> ";" Then
                LastNonEmpty += 1
                TestArray(LastNonEmpty) = TestArray(i)
            End If
        Next
        ReDim Preserve TestArray(LastNonEmpty)
        ' TestArray now holds {"apple", "pear", "banana"}

        TextBox4.Text = TestArray(1) & "  ---> " & TextBox1.Text & "\" & TestArray(4) & ".inlive"
        Leer_Archivo(TestArray(1), TextBox1.Text & "\" & TestArray(4) & ".inlive")
        REM Logo.Url = New Uri(TestArray(1))
    End Sub
    Private Sub Descargar_Logos()
        Dim control As Boolean = True
        MsgBox("Logos --> " + CType(contador, String))
        ProgressBar1.Value = 0
        ProgressBar1.Maximum = contador + 1

        Label3.Text = "Descargando logos  ......Espere"
        For x = 0 To contador
            contador1 = x

            My.Application.DoEvents()
            Try
                Try
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Catch ex As Exception
                    ProgressBar1.Value = 0
                End Try

                TextBox3.Text = ArchivoWeb(x)
                TextBox4.Text = "Registro -->" + CType(x + 1, String) + " --->" + ArchivoWeb(x) + " -->" + ArchivoPc(x) + vbCrLf
                Label3.Text = "Descargando logos  ...... " + CType(x, String) + " Errores -> " + CType(Errores, String)
                My.Computer.Network.DownloadFile(ArchivoWeb(x), ArchivoPc(x), "anonymous", "")
                VerImagen(ArchivoPc(x))
            Catch ex As Exception
                Try
                    VerImagen(ArchivoPc(x))
                    TextBox4.Text = TextBox4.Text + "Error -->" + ex.Message

                Catch ei As Exception
                    PictureBox2.Image = InLive.My.Resources.Resources.Foto0976
                    TextBox4.Text = TextBox4.Text + "Error -->" + ei.Message
                    Errores += 1
                End Try

            End Try
            If (Rapido) Then
                Dim response = MsgBox("Siguiente", MsgBoxStyle.OkCancel, "In Live")
                If response = MsgBoxResult.Cancel Then
                    Application.Exit()
                    x = contador

                End If
            End If
        Next

        PictureBox2.Image = InLive.My.Resources.Resources.cordova
            Label3.Text = "Descarga completada  Errores -> " + CType(Errores, String)
            TextBox4.Text = "                 Descarga completada" + vbLf + vbLf + "                 Pulse Salir para cerrar el programa"



    End Sub
    Public Sub VerImagen(fileToDisplay As String)
        ' Sets up an image object to be displayed.
        Dim MyImage As Bitmap = Nothing
        If (MyImage IsNot Nothing) Then
            MyImage.Dispose()
        End If

        ' Stretches the image to fit the pictureBox. 
        PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
        MyImage = New Bitmap(fileToDisplay)
        PictureBox2.Image = CType(MyImage, Image)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        REM TextBox4.Text = ArchivoWeb(contador1) + vbCrLf
        If My.Computer.Keyboard.CtrlKeyDown Then
            Timer1.Enabled = False

            Dim response = MsgBox("Finalizar", MsgBoxStyle.YesNo, "In Live")
            If response = MsgBoxResult.Yes Then
                Application.Exit()
            End If

            Timer1.Enabled = True

        Else
            REM  MsgBox("CTRL key up")
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Rapido = False
        Try


            ReadFile()

        Catch Ex As Exception

            MessageBox.Show(Ex.Message)

        End Try
    End Sub

    Private Sub ProgressBar1_Click(sender As Object, e As EventArgs) Handles ProgressBar1.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
