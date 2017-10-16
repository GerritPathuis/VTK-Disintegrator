Imports System.IO
Imports System.Text
Imports System.Math
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word
Imports System.Management

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged
        Calc_tab1()
    End Sub

    Private Sub Calc_tab1()
        Dim power, rpm, rad, torque, dia_beater As Double
        Dim l_wet, l_add, l_tot As Double
        Dim tip_speed, acc, acc_time As Double
        Dim lump_dia, lump_weight, density, f_tip As Double
        Dim key_h, key_l, shaft_radius, max_key_torque As Double
        Dim yield As Double

        power = NumericUpDown1.Value
        rpm = NumericUpDown2.Value
        dia_beater = NumericUpDown8.Value / 1000  '[m]
        lump_dia = NumericUpDown14.Value / 1000  '[m]
        acc_time = NumericUpDown15.Value
        density = NumericUpDown16.Value
        key_h = NumericUpDown17.Value   '[mm]
        key_l = NumericUpDown9.Value   '[mm]
        shaft_radius = NumericUpDown12.Value / 2000   '[mm]
        yield = NumericUpDown18.Value / 1.5 / 1.25  '[N/mm2]

        rad = rpm / 60 * 2 * PI
        torque = power * 1000 / rad

        l_wet = NumericUpDown3.Value
        l_add = NumericUpDown4.Value
        l_tot = (l_add + l_wet) * 1000 / 3600   '[kg/s]

        tip_speed = dia_beater * rpm * PI / 60  '[m/s]

        lump_weight = 4 / 3 * PI * (lump_dia / 2) ^ 3 * density

        acc = tip_speed / acc_time      '[m/s2]
        f_tip = lump_weight * acc       '[N]

        max_key_torque = key_h * key_l * yield * shaft_radius

        TextBox1.Text = l_tot.ToString("0.0")
        TextBox2.Text = rad.ToString("0.0")
        TextBox3.Text = torque.ToString("0")
        TextBox4.Text = tip_speed.ToString("0.0")
        TextBox5.Text = lump_weight.ToString("0.0")
        TextBox6.Text = acc.ToString("0")
        TextBox7.Text = f_tip.ToString("0")
        TextBox8.Text = key_l.ToString("0")
        TextBox9.Text = max_key_torque.ToString("0")
    End Sub


End Class
