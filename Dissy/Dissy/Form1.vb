Imports System.IO
Imports System.Text
Imports System.Math
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word
Imports System.Management

Public Class Form1
    'according to DIN6885-1
    Public Shared shaft_key() As String = {
    "6;8;2x2;1.2;1.0",
    "8;10;3x3;1.8;1.4",
    "10;12;4x4;2.5;1.8",
    "12;17;5x5;3.0;2.3",
    "17;22;6x6;3.5;2.8",
    "22;30;8x7;4.0;3.3",
    "30;38;10x8;5.0;3.3",
    "38;44;12x8;5.0;3.3",
    "44;50;14x9;5.5;3.8",
    "50;58;16x10;6;4.3",
    "58;65;18x11;7;4.4",
    "65;75;20x12;7.5;4.9",
    "75;85;22x14;9;5.4",
    "85;95;25x14;9;5.4",
    "95;110;28x16;10;6.4",
    "110;130;32x18;11;7.4",
    "130;150;36x20;12;8.4",
    "150;170;40x22;13;9.4",
    "170;200;45x25;15;10.4",
    "200;230;50x28;17;11.4",
    "230;260;56x32;20;12.4",
    "260;290;63x32;20;12.4",
    "290;330;70x36;22;14.4",
    "330;380;80x40;25;15.4",
    "380;440;90x45;28;17.4",
    "440;550;100x50;31;19.5"}

    Public words() As String
    Public separators() As String = {";"}

    Private Sub Button1_Click(sender As Object, E As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown18.ValueChanged, ComboBox1.SelectedIndexChanged
        Calc_tab1()
    End Sub

    Private Sub Calc_tab1()
        Dim power, rpm, rad, torque, dia_beater As Double
        Dim l_wet, l_add, l_tot As Double
        Dim tip_speed, acc, acc_time As Double
        Dim lump_dia, lump_weight, density, f_tip, lump_torque As Double
        Dim key_h, key_l, shaft_radius, max_key_torque, max_key_force As Double
        Dim allowed_stress As Double

        If ComboBox1.SelectedIndex > -1 Then
            words = shaft_key(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)

            TextBox13.Text = words(3)       '(t1) Key depth in shaft
            TextBox15.Text = words(4)       '(t2) Key above shaft
            TextBox16.Text = words(1)       'Max shaft diameter [mm]
            TextBox14.Text = words(2)       'Key size
        End If

        Double.TryParse(TextBox13.Text, key_h)      '[mm]
        key_l = NumericUpDown9.Value                '[mm]

        power = NumericUpDown1.Value
        rpm = NumericUpDown2.Value
        dia_beater = NumericUpDown8.Value / 1000    '[m]
        lump_dia = NumericUpDown14.Value / 1000     '[m]
        acc_time = NumericUpDown15.Value
        density = NumericUpDown16.Value

        shaft_radius = NumericUpDown12.Value / 2000   '[mm]
        allowed_stress = NumericUpDown18.Value / (1.5 * 1.25)  '[N/mm2]

        rad = rpm / 60 * 2 * PI
        torque = power * 1000 / rad

        l_wet = NumericUpDown3.Value
        l_add = NumericUpDown4.Value
        l_tot = (l_add + l_wet) * 1000 / 3600   '[kg/s]

        tip_speed = dia_beater * rpm * PI / 60  '[m/s]

        lump_weight = 4 / 3 * PI * (lump_dia / 2) ^ 3 * density

        acc = tip_speed / acc_time              '[m/s2]
        f_tip = lump_weight * acc               '[N]
        lump_torque = f_tip * (dia_beater / 2)  '[N.m]

        max_key_force = key_h * key_l * allowed_stress
        max_key_torque = max_key_force * shaft_radius
        max_key_torque *= 2     'two keys 

        TextBox1.Text = l_tot.ToString("0.0")
        TextBox2.Text = rad.ToString("0.0")
        TextBox3.Text = torque.ToString("0")
        TextBox4.Text = tip_speed.ToString("0.0")
        TextBox5.Text = lump_weight.ToString("0.0")
        TextBox6.Text = acc.ToString("0")
        TextBox7.Text = f_tip.ToString("0")
        TextBox8.Text = key_l.ToString("0")
        TextBox9.Text = (max_key_torque / 1000).ToString("0")
        TextBox10.Text = lump_torque.ToString("0")
        TextBox11.Text = (max_key_force / 1000).ToString("0")
        TextBox12.Text = allowed_stress.ToString("0")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        ComboBox1.Items.Clear()                     'Note Combobox1 contains"startup" to prevent exceptions

        For hh = 0 To (shaft_key.Length - 1)  'Fill combobox4 Motor data
            words = shaft_key(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 23, -1)) 'Select ..

    End Sub
End Class
