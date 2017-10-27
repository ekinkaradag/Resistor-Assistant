Public Class Form1


    'Ekin Karadağ's Term Project
    'Resistor Assistant
    'Copyright© 2013

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Make the ComboBoxes loaded as something for the first run
        ComboBox1.Text = "BLACK"
        ComboBox2.Text = "BLACK"
        ComboBox3.Text = "BLACK"
        ComboBox4.Text = "GOLD"
        ComboBox5.Text = "0"
        ComboBox6.Text = "0"
        ComboBox7.Text = "1"
        ComboBox8.Text = "± %5"

        'Make colors compatible with ComboBoxes
        PictureBox7.BackColor = Color.Black
        PictureBox8.BackColor = Color.Black
        PictureBox10.BackColor = Color.Black
        PictureBox9.BackColor = Color.Gold



        'Get to know color and its values
        Dim dt As New DataTable
            dt.Columns.Add("color")
            dt.Columns.Add("value", GetType(Long))
            dt.Rows.Add(New Object() {"BLACK", 1})
            dt.Rows.Add(New Object() {"BROWN", 10})
            dt.Rows.Add(New Object() {"RED", 100})
            dt.Rows.Add(New Object() {"ORANGE", 1000})
            dt.Rows.Add(New Object() {"YELLOW", 10000})
            dt.Rows.Add(New Object() {"GREEN", 100000})
            dt.Rows.Add(New Object() {"BLUE", 1000000})
            dt.Rows.Add(New Object() {"PURPLE", 10000000})
            dt.Rows.Add(New Object() {"GRAY", 100000000})
            dt.Rows.Add(New Object() {"WHITE", 1000000000})

            Dim colors() As String = {"BLACK", "BROWN", "RED", "ORANGE", "YELLOW", "GREEN", "BLUE", "PURPLE", "GRAY", "WHITE"}

        'The first two ComboBoxes represents the coefficents of the resistor
            ComboBox1.DataSource = colors.Clone
        ComboBox2.DataSource = colors.Clone

        'The third ComboBox represents the multiplier of the resistor
            ComboBox3.DisplayMember = "color"
            ComboBox3.ValueMember = "value"
            ComboBox3.DataSource = dt.Copy

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Make the calculation that depends on the ComboBoxes
        Dim totalVal As Long = (10 * ComboBox1.SelectedIndex + ComboBox2.SelectedIndex) * CLng(ComboBox3.SelectedValue) 'Make the calculation
        ToolTip1.SetToolTip(Label1, totalVal.ToString("n0"))
        Dim values As New List(Of Object())
        'Make the numbers shorter with letters
        values.Add(New Object() {1000, "K"})
        values.Add(New Object() {1000000, "M"})
        values.Add(New Object() {1000000000, "G"})
        Dim output = (From v In values Where CLng(v(0)) <= totalVal).LastOrDefault

        'Give the result
        Label1.Text = "Resistor's value = " & If(output IsNot Nothing, String.Format("{0}{1}", totalVal / CLng(output(0)), output(1)), totalVal.ToString) & " Ω"


        'Make the tolerance value visible
        If ComboBox4.Text = "COLORLESS" Then
            Label2.Text = "Tolerance value = ± %20"
        ElseIf ComboBox4.Text = "SILVER" Then
            Label2.Text = "Tolerance value = ± %10"
        ElseIf ComboBox4.Text = "GOLD" Then
            Label2.Text = "Tolerance value = ± %5"
        ElseIf ComboBox4.Text = "GRAY" Then
            Label2.Text = "Tolerance value = ± %0.05"
        ElseIf ComboBox4.Text = "PURPLE" Then
            Label2.Text = "Tolerance value = ± %0.10"
        ElseIf ComboBox4.Text = "BLUE" Then
            Label2.Text = "Tolerance value = ± %0.25"
        ElseIf ComboBox4.Text = "GREEN" Then
            Label2.Text = "Tolerance value = ± %0.5"
        ElseIf ComboBox4.Text = "RED" Then
            Label2.Text = "Tolerance value = ± %2"
        ElseIf ComboBox4.Text = "BROWN" Then
            Label2.Text = "Tolerance value = ± %1"
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        'Color the tolerance value on the resistor
        PictureBox5.BackColor = Color.FromName(ComboBox4.Text)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        'Color the multiplier band on the resistor
        PictureBox3.BackColor = Color.FromName(ComboBox3.Text)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        'Color the second band on the resistor
        PictureBox2.BackColor = Color.FromName(ComboBox2.Text)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        'Color the first band on the resistor
        PictureBox1.BackColor = Color.FromName(ComboBox1.Text)
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        'Make the TextBoxes numeric-only
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress

        'Make the TextBoxes numeric-only
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress

        'Make the TextBoxes numeric-only
        If e.KeyChar <> ChrW(Keys.Back) Then
            If Char.IsNumber(e.KeyChar) Then
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Get to know the inputs
        Dim a, b, c, d, f, g As Single
        If RadioButton1.Checked = True Then
            'Make the calculation
            a = CSng(TextBox1.Text)
            b = CSng(TextBox2.Text)
            c = CSng(TextBox3.Text)
            g = (c / 1000)
            d = ((a - b) / g)
            f = ((a - b) * g)
            Label3.Text = "Needed resistor = " & d & " Ω"
            Label4.Text = "Needed resistor's power = " & f & " W"
        ElseIf RadioButton1.Checked = False Then
            'Make the calculation
            a = CSng(TextBox1.Text)
            b = CSng(TextBox2.Text)
            c = CSng(TextBox3.Text)
            d = ((a - b) / c)
            f = ((a - b) * c)
            Label3.Text = "Needed resistor = " & d & " Ω"
            Label4.Text = "Needed resistor's power = " & f & " W"
        End If
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        System.Diagnostics.Process.Start("https://ekinkaradag.com")
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        If ComboBox5.Text = "0" Then
            PictureBox8.BackColor = Color.Black
        ElseIf ComboBox5.Text = "1" Then
            PictureBox8.BackColor = Color.Brown
        ElseIf ComboBox5.Text = "2" Then
            PictureBox8.BackColor = Color.Red
        ElseIf ComboBox5.Text = "3" Then
            PictureBox8.BackColor = Color.Orange
        ElseIf ComboBox5.Text = "4" Then
            PictureBox8.BackColor = Color.Yellow
        ElseIf ComboBox5.Text = "5" Then
            PictureBox8.BackColor = Color.Green
        ElseIf ComboBox5.Text = "6" Then
            PictureBox8.BackColor = Color.Blue
        ElseIf ComboBox5.Text = "7" Then
            PictureBox8.BackColor = Color.Purple
        ElseIf ComboBox5.Text = "8" Then
            PictureBox8.BackColor = Color.Gray
        ElseIf ComboBox5.Text = "9" Then
            PictureBox8.BackColor = Color.White
        End If
        Label13.Text = PictureBox8.BackColor.Name
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If ComboBox6.Text = "0" Then
            PictureBox7.BackColor = Color.Black
        ElseIf ComboBox6.Text = "1" Then
            PictureBox7.BackColor = Color.Brown
        ElseIf ComboBox6.Text = "2" Then
            PictureBox7.BackColor = Color.Red
        ElseIf ComboBox6.Text = "3" Then
            PictureBox7.BackColor = Color.Orange
        ElseIf ComboBox6.Text = "4" Then
            PictureBox7.BackColor = Color.Yellow
        ElseIf ComboBox6.Text = "5" Then
            PictureBox7.BackColor = Color.Green
        ElseIf ComboBox6.Text = "6" Then
            PictureBox7.BackColor = Color.Blue
        ElseIf ComboBox6.Text = "7" Then
            PictureBox7.BackColor = Color.Purple
        ElseIf ComboBox6.Text = "8" Then
            PictureBox7.BackColor = Color.Gray
        ElseIf ComboBox6.Text = "9" Then
            PictureBox7.BackColor = Color.White
        End If
        Label15.Text = PictureBox7.BackColor.Name
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        If ComboBox7.Text = "1" Then
            PictureBox10.BackColor = Color.Black
        ElseIf ComboBox7.Text = "10" Then
            PictureBox10.BackColor = Color.Brown
        ElseIf ComboBox7.Text = "100" Then
            PictureBox10.BackColor = Color.Red
        ElseIf ComboBox7.Text = "1k" Then
            PictureBox10.BackColor = Color.Orange
        ElseIf ComboBox7.Text = "10k" Then
            PictureBox10.BackColor = Color.Yellow
        ElseIf ComboBox7.Text = "100k" Then
            PictureBox10.BackColor = Color.Green
        ElseIf ComboBox7.Text = "1M" Then
            PictureBox10.BackColor = Color.Blue
        ElseIf ComboBox7.Text = "10M" Then
            PictureBox10.BackColor = Color.Purple
        ElseIf ComboBox7.Text = "0.1" Then
            PictureBox10.BackColor = Color.Gold
        ElseIf ComboBox7.Text = "0.01" Then
            PictureBox10.BackColor = Color.Silver
        End If
        Label17.Text = PictureBox10.BackColor.Name
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        If ComboBox8.Text = "± %1" Then
            PictureBox9.BackColor = Color.Brown
        ElseIf ComboBox8.Text = "± %2" Then
            PictureBox9.BackColor = Color.Red
        ElseIf ComboBox8.Text = "± %0.5" Then
            PictureBox9.BackColor = Color.Green
        ElseIf ComboBox8.Text = "± %0.25" Then
            PictureBox9.BackColor = Color.Blue
        ElseIf ComboBox8.Text = "± %0.10" Then
            PictureBox9.BackColor = Color.Purple
        ElseIf ComboBox8.Text = "± %0.05" Then
            PictureBox9.BackColor = Color.Gray
        ElseIf ComboBox8.Text = "± %5" Then
            PictureBox9.BackColor = Color.Gold
        ElseIf ComboBox8.Text = "± %10" Then
            PictureBox9.BackColor = Color.Silver
        ElseIf ComboBox8.Text = "± %20" Then
            PictureBox9.Visible = False
        End If
        If ComboBox8.Text = "± %20" = True Then
            Label19.Text = "Colorless (No Color)"
        ElseIf ComboBox8.Text = "± %20" = False Then
            PictureBox9.Visible = True
            Label19.Text = PictureBox9.BackColor.Name
        End If
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        System.Diagnostics.Process.Start("https://github.com/ekinkaradag/Resistor-Assistant")
    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click
        System.Diagnostics.Process.Start("https://twitter.com/karadagekin")
    End Sub

    Private Sub PictureBox15_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click
        System.Diagnostics.Process.Start("https://instagram.com/karadagekin")
    End Sub

    Private Sub PictureBox16_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click
        System.Diagnostics.Process.Start("https://www.linkedin.com/in/ekinkaradag/en")
    End Sub
End Class
