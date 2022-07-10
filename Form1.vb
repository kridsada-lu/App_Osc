
Public Class Form1
    Dim x, y1, y2 As Int16
    Dim dum01 As Int16
    Dim i, j, it1, it2 As Int16
    Dim Time(150), CH1(150), CH2(150), t0 As Double
    Dim num As Int16


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        My.Computer.FileSystem.CreateDirectory("D:\Data65_OSC")
        Dim data1 As String
        data1 = "time" & vbTab & "CH1" & vbTab & "CH2" & vbNewLine
        My.Computer.FileSystem.WriteAllText("D:\Data65_OSC\01.txt", data1, True)
        num = 71


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ComboBox1.Items.Clear()
        Dim myPort As Array
        myPort = IO.Ports.SerialPort.GetPortNames()
        ComboBox1.Items.AddRange(myPort)
        'ComboBox1.SelectedIndex = 0
        ComboBox1.DroppedDown = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "Connect" Then
            Button3.BackColor = Color.Red
            Button3.Text = "Disconnect"
            'SerialPort1.BaudRate = 1000000
            'SerialPort1.PortName = COM5
            SerialPort1.BaudRate = ComboBox2.SelectedItem 'Or ComboBox2.Text
            SerialPort1.PortName = ComboBox1.SelectedItem 'Or ComboBox1.Text
            SerialPort1.Open()
            Timer1.Start()
            'dum01 = Val(TextBox1.Text)
            'Timer1.Enabled = True
        Else
            Button3.BackColor = Color.Green
            Button3.Text = "Connect"
            Timer1.Stop()
            'i = 0
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If SerialPort1.BytesToRead > 0 Then  ' ตรวจสอบว่าArduino มีการส่งข้อมูลเข้ามาทาง Port ไหม
            'Dim Time(50), CH1(50), CH2(50) As Double
            Dim DataInp() As String = Split(SerialPort1.ReadLine, ",")   ' เก็บค่าที่ Arduino ส่งมา
            'Dim DataInp As String = SerialPort1.ReadByte  ' เก็บค่าที่ Arduino ส่งมา
            TextBox1.Text = DataInp(0)   'เอาค่าที่ Arduino ส่งมาแสดงใน Textbox
            On Error Resume Next
            TextBox2.Text = DataInp(1)
            On Error Resume Next
            TextBox3.Text = DataInp(2)
            On Error Resume Next



            'Chart1.Series("CH1").Points.AddXY(0.0, 0.0)
            'Chart1.Series("CH2").Points.AddXY(0.0, 0.0)

            Chart1.Series("CH1").Points.AddXY(TextBox1.Text / 1, TextBox2.Text)
            Chart1.Series("CH2").Points.AddXY(TextBox1.Text / 1, TextBox3.Text)

            Dim data2 As String
            data2 = TextBox1.Text & vbTab & TextBox2.Text & vbTab & TextBox3.Text & vbNewLine
            My.Computer.FileSystem.WriteAllText("D:\Data65_OSC\01.txt", data2, True)

            i = i + 1
            'DataGridView1.Rows.Add(i, TextBox1.Text / 1000, (TextBox2.Text * 3.3) / 1023, (TextBox3.Text * 3.3) / 1023)


            Time(i) = TextBox1.Text
            CH1(i) = TextBox2.Text
            CH2(i) = TextBox3.Text
            'DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowMode.AllCellsExceptHeader
            'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            'DataGridView1.Rows.Add(i, Time(i) / 1000, CH1(i), CH2(i))

            If i = num Then Timer1.Enabled = False




            'Dim s_ch1() As Double = New Double() {CH1(1), CH1(2), CH1(3), CH1(4), CH1(5), CH1(6), CH1(7), CH1(8), CH1(9), CH1(10) And
            'CH1(11), CH1(12), CH1(13), CH1(14), CH1(15), CH1(16), CH1(17), CH1(18), CH1(19), CH1(20) And
            'CH1(21), CH1(22), CH1(23), CH1(24), CH1(55), CH1(26), CH1(27), CH1(28), CH1(29), CH1(30) And
            'CH1(31), CH1(32), CH1(33), CH1(34), CH1(35), CH1(36), CH1(37), CH1(38), CH1(39), CH1(40) And
            'CH1(41), CH1(42), CH1(43), CH1(44), CH1(45), CH1(46), CH1(47), CH1(48), CH1(49), CH1(50)}


        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Max------------------------------------------------------------------------------
        TextBox4.Text = (From row As DataGridViewRow In DataGridView1.Rows
                         Where row.Cells(2).FormattedValue.ToString() <> String.Empty
                         Select Convert.ToInt32(row.Cells(2).FormattedValue)
                         ).Max().ToString()

        TextBox17.Text = (From row As DataGridViewRow In DataGridView1.Rows
                          Where row.Cells(3).FormattedValue.ToString() <> String.Empty
                          Select Convert.ToInt32(row.Cells(3).FormattedValue)
                         ).Max().ToString()
        '-----------------------------------------------------------------------------------
        'Min------------------------------------------------------------------------------
        TextBox5.Text = (From row As DataGridViewRow In DataGridView1.Rows
                         Where row.Cells(2).FormattedValue.ToString() <> String.Empty
                         Select Convert.ToInt32(row.Cells(2).FormattedValue)
                        ).Min().ToString()

        TextBox16.Text = (From row As DataGridViewRow In DataGridView1.Rows
                          Where row.Cells(3).FormattedValue.ToString() <> String.Empty
                          Select Convert.ToInt32(row.Cells(3).FormattedValue)
                         ).Min().ToString()
        '---Avg-----------------------------------------------------------------------------
        TextBox6.Text = (Val(TextBox4.Text) + Val(TextBox5.Text)) / 2

        TextBox15.Text = (Val(TextBox17.Text) + Val(TextBox16.Text)) / 2
        TextBox19.Text = (Val(TextBox6.Text) + Val(TextBox15.Text)) / 2

        '---Amp-----------------------------------------------------------------------------
        TextBox7.Text = (Val(TextBox4.Text) - Val(TextBox19.Text))
        TextBox14.Text = (Val(TextBox17.Text) - Val(TextBox19.Text))
        '---Amp-----------------------------------------------------------------------------

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        TextBox25.Text = TextBox22.Text
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        TextBox26.Text = TextBox22.Text
        TextBox27.Text = (Val(TextBox26.Text) - Val(TextBox25.Text))
        TextBox13.Text = 1 / Val(TextBox27.Text) * 1000000

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        'TextBox27.Text = (Val(TextBox26.Text) - Val(TextBox25.Text))
        'TextBox13.Text = 1 / Val(TextBox27.Text) * 1000000

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        'TextBox30.Text = (Val(TextBox29.Text) - Val(TextBox28.Text))
        'TextBox18.Text = 1 / Val(TextBox30.Text) * 1000000
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        TextBox28.Text = TextBox22.Text
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        TextBox29.Text = TextBox22.Text
        TextBox30.Text = (Val(TextBox29.Text) - Val(TextBox28.Text))
        TextBox18.Text = 1 / Val(TextBox30.Text) * 1000000
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)
        'TextBox9.Text = (((Math.Asin(Val(TextBox23.Text) / Val(TextBox7.Text))) * 180) / 3.14159) - ((6.28318 / Val(TextBox27.Text)) * Val(TextBox11.Text))
        'TextBox9.Text = (Math.Asin(Val(TextBox23.Text) / Val(TextBox7.Text))) - ((6.28318 / Val(TextBox27.Text)) * Val(TextBox11.Text))
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)
        'TextBox10.Text = (Math.Asin(Val(TextBox24.Text) / Val(TextBox14.Text))) - ((6.28318 / Val(TextBox30.Text)) * Val(TextBox11.Text))


        'TextBox12.Text = Val(TextBox9.Text) - Val(TextBox10.Text)
        'TextBox8.Text = (Val(TextBox12.Text) * 180) / 3.14159
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)
        'TextBox23.Text = Val(TextBox21.Text) - Val(TextBox19.Text)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs)
        'TextBox24.Text = Val(TextBox20.Text) - Val(TextBox19.Text)
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs)
        'TextBox11.Text = TextBox22.Text
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        TextBox31.Text = TextBox22.Text
        TextBox32.Text = TextBox21.Text
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        TextBox33.Text = TextBox22.Text
        TextBox34.Text = TextBox21.Text
        Dim m As Double
        m = (Val(TextBox32.Text) - Val(TextBox34.Text)) / (Val(TextBox31.Text) - Val(TextBox33.Text))
        TextBox35.Text = m
        Dim C1 As Double
        C1 = Val(TextBox32.Text) - m * Val(TextBox31.Text)
        TextBox45.Text = C1
        TextBox36.Text = FormatNumber((Val(TextBox19.Text) - C1) / m, 5)



    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        TextBox42.Text = TextBox22.Text
        TextBox41.Text = TextBox20.Text
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'For it1 = 100 To 400 Step 1
        'Dim x1 As Int16
        'x1 = Val(TextBox36.Text)
        'Chart1.Series("CH1").Points.AddXY(500, it1)
        'Next it1
    End Sub



    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        TextBox40.Text = TextBox22.Text
        TextBox39.Text = TextBox20.Text
        Dim m2 As Double
        m2 = (Val(TextBox41.Text) - Val(TextBox39.Text)) / (Val(TextBox42.Text) - Val(TextBox40.Text))
        TextBox38.Text = m2
        Dim C2 As Double
        C2 = Val(TextBox41.Text) - m2 * Val(TextBox42.Text)
        TextBox46.Text = C2
        TextBox37.Text = FormatNumber((Val(TextBox19.Text) - C2) / m2, 5)

    End Sub

    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        Dim dt As Double
        dt = Math.Abs(TextBox37.Text - TextBox36.Text)
        TextBox48.Text = dt
        TextBox47.Text = FormatNumber(Val(TextBox48.Text) * 360 / Val(TextBox27.Text), 3)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim index As Int16
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView1.Rows(index)
        TextBox22.Text = selectedRow.Cells(1).Value.ToString()
        TextBox21.Text = selectedRow.Cells(2).Value.ToString()
        TextBox20.Text = selectedRow.Cells(3).Value.ToString()

        ' y1(t) and y2(t)--------------------------------------
        'TextBox23.Text = Val(TextBox21.Text) - Val(TextBox19.Text)
        'TextBox24.Text = Val(TextBox20.Text) - Val(TextBox19.Text)
        ' Y0-------------------------------------------------------
        'TextBox8.Text = Val(TextBox21.Text) - Val(TextBox19.Text)
        'TextBox13.Text = Val(TextBox20.Text) - Val(TextBox19.Text)
        ' มุมเรเดียน--------------------------------------------
        'TextBox9.Text = Math.Asin(Val(TextBox8.Text) / Val(TextBox7.Text))
        'TextBox12.Text = Math.Asin(Val(TextBox13.Text) / Val(TextBox14.Text))
        'มุมองศา------------------------------------------------
        'TextBox10.Text = ((Val(TextBox9.Text) * 180) / 3.14159)
        'TextBox11.Text = ((Val(TextBox12.Text) * 180) / 3.14159)
        'ความต่างเฟส-----------------------------------------------
        'TextBox18.Text = (Val(TextBox10.Text) - Val(TextBox11.Text))

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        For j = 1 To num - 1 Step 1
            t0 = Time(j + 1) - Time(2)
            'DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowMode.AllCellsExceptHeader
            'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            DataGridView1.Rows.Add(j, t0, CH1(j + 1), CH2(j + 1))
            'Chart1.Series("CH1").Points.AddXY(Time(j) - t0, CH1(j))
            'Chart1.Series("CH2").Points.AddXY(Time(j) - t0, CH2(j))

            'TextBox22.Text = j
        Next j
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click, Chart1.SizeChanged

    End Sub
End Class
