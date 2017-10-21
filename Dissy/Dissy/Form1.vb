﻿Imports System.IO
Imports System.Text
Imports System.Math

Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Management

Public Class Form1

    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Dissy_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Dissy_rapport_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

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

    Private Sub Button1_Click(sender As Object, E As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown1.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown17.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, ComboBox1.SelectedIndexChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown10.ValueChanged
        Calc_tab1()
    End Sub

    Private Sub Calc_tab1()
        Dim Installed_power, rpm, rad, motor_torque, dia_beater As Double
        Dim l_wet, l_add, l_tot As Double
        Dim tip_speed, acc, acc_time As Double
        Dim lump_dia, lump_weight, density, f_tip, lump_torque, safety_lump As Double
        Dim key_h, key_l, shaft_radius, max_key_torque, max_key_force As Double
        Dim start_torque, allowed_stress As Double
        Dim specific_load, load_beater_tip As Double
        Dim no_beaters As Double

        If ComboBox1.SelectedIndex > -1 Then
            words = shaft_key(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox13.Text = words(3)       '(t1) Key depth in shaft
            TextBox15.Text = words(4)       '(t2) Key above shaft
            TextBox16.Text = words(1)       'Max shaft diameter [mm]
            TextBox14.Text = words(2)       'Key size
        End If

        If ComboBox2.SelectedIndex > -1 Then
            words = shaft_key(ComboBox2.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox17.Text = words(1)       'Max shaft diameter [mm]
            TextBox18.Text = words(2)       'Key size
            TextBox21.Text = words(3)       '(t1) Key depth in shaft
        End If

        no_beaters = NumericUpDown7.Value
        Double.TryParse(TextBox13.Text, key_h)      '[mm]
        key_l = NumericUpDown9.Value                '[mm]

        Installed_power = NumericUpDown1.Value * 1000    '[W]
        rpm = NumericUpDown2.Value
        dia_beater = NumericUpDown8.Value / 1000    '[m]
        lump_dia = NumericUpDown14.Value / 1000     '[m]
        acc_time = NumericUpDown15.Value
        density = NumericUpDown16.Value

        shaft_radius = NumericUpDown12.Value / 2000   '[mm]
        allowed_stress = NumericUpDown18.Value / (1.5 * 1.25)  '[N/mm2]

        rad = rpm / 60 * 2 * PI
        motor_torque = Installed_power / rad
        start_torque = motor_torque * 2.0

        l_wet = NumericUpDown3.Value
        l_add = NumericUpDown4.Value
        l_tot = (l_add + l_wet) * 1000 / 3600   '[kg/s]

        specific_load = 3600 * l_tot / Installed_power  '[ton/(kW.hr)]
        load_beater_tip = l_tot / (rpm / 60 * no_beaters * 2)

        tip_speed = dia_beater * rpm * PI / 60  '[m/s]

        '---- Lump calculation--------
        lump_weight = 4 / 3 * PI * (lump_dia / 2) ^ 3 * density

        acc = tip_speed / acc_time              '[m/s2]
        f_tip = lump_weight * acc               '[N]
        lump_torque = f_tip * (dia_beater / 2)  '[N.m]

        max_key_force = key_h * key_l * allowed_stress
        max_key_torque = max_key_force * shaft_radius
        max_key_torque *= 2     'two keys 
        safety_lump = max_key_torque / lump_torque

        '--------- max Allowed power on coupling key --
        Dim drive_key_force, drive_l, drive_b, drive_r As Double
        Dim drive_power_max, safety_coupling_key, safety_nut As Double

        Double.TryParse(TextBox21.Text, drive_b)            '[mm] key t1
        drive_l = NumericUpDown17.Value                     '[mm]  
        drive_key_force = allowed_stress * drive_b * drive_l '[N]
        drive_r = NumericUpDown13.Value / 2000              '[m] radius
        drive_power_max = drive_key_force * drive_r * rad   '[W]
        safety_coupling_key = drive_power_max / start_torque

        '--------- Hydraulic nut (spacer = friction disk) --
        Dim spacer_od, spacer_id, spacer_radius, fric As Double
        Dim max_torque, delta_l, shaft_l, pull_force, area As Double
        pull_force = NumericUpDown19.Value
        spacer_od = NumericUpDown21.Value
        fric = NumericUpDown6.Value
        spacer_id = NumericUpDown12.Value
        spacer_radius = (spacer_od + spacer_id) / 4
        max_torque = pull_force * fric * (spacer_radius / 1000)     '[kNm]
        safety_nut = max_torque / start_torque * 1000                     '[-]
        area = PI / 4 * spacer_id ^ 2                               '[mm2]
        shaft_l = (NumericUpDown9.Value + NumericUpDown11.Value) * no_beaters '[mm]

        delta_l = pull_force * 10 ^ 3 * shaft_l / (190000 * area)            '[mm]

        '-------- present-------
        TextBox1.Text = l_tot.ToString("0")
        TextBox2.Text = rad.ToString("0.0")
        TextBox3.Text = (motor_torque / 1000).ToString("0.0") '[kNm]
        TextBox4.Text = tip_speed.ToString("0.0")
        TextBox5.Text = lump_weight.ToString("0.00")
        TextBox6.Text = acc.ToString("0")
        TextBox7.Text = f_tip.ToString("0")
        TextBox8.Text = key_l.ToString("0")
        TextBox9.Text = (max_key_torque / 1000).ToString("0.0") '[kNm]
        TextBox10.Text = (lump_torque / 1000).ToString("0.0")   '[kNm]
        TextBox11.Text = (max_key_force / 1000).ToString("0")   '[kN]
        TextBox12.Text = allowed_stress.ToString("0")
        TextBox19.Text = allowed_stress.ToString("0")
        TextBox20.Text = (drive_power_max / 1000).ToString("0") '[kNm]
        TextBox22.Text = (drive_key_force / 1000).ToString("0") '[km]
        TextBox23.Text = spacer_id.ToString("0")                '[mm]
        TextBox24.Text = spacer_radius.ToString("0")            '[mm]
        TextBox25.Text = max_torque.ToString("0")               '[kNm]
        TextBox29.Text = (start_torque / 1000).ToString("0.0")  '[kNm]
        TextBox32.Text = safety_coupling_key.ToString("0.0")    '[kNm]
        TextBox33.Text = specific_load.ToString("0.0")          '[]
        TextBox34.Text = delta_l.ToString("0.00")               '[mm]
        TextBox35.Text = shaft_l.ToString("0")                  '[mm]
        TextBox36.Text = load_beater_tip.ToString("0.0")        '[kg]
        TextBox37.Text = spacer_od.ToString("0")                '[mm]
        TextBox39.Text = safety_nut.ToString("0.0")             '[-]
        TextBox40.Text = safety_lump.ToString("0.0")            '[-]

        '------- checks---------
        'TextBox25.BackColor = IIf(max_torque < start_torque, Color.LightGreen, Color.Red)
        TextBox39.BackColor = IIf(safety_nut > 1.2, Color.LightGreen, Color.Red)
        TextBox40.BackColor = IIf(safety_lump > 1.2, Color.LightGreen, Color.Red)
        Calc_inertia()
        Calc_shaft_coupling()
        Calc_shaft_beaters()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------- directory's-----------

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        ComboBox1.Items.Clear()                     'Note Combobox1 contains
        ComboBox2.Items.Clear()                     'Note Combobox1 contains

        For hh = 0 To (shaft_key.Length - 1)  'Fill combobox4 Motor data
            words = shaft_key(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 13, -1)) 'Select ..
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 18, -1)) 'Select ..
    End Sub
    Private Sub Calc_inertia()
        Dim overall_length, half_beater_weight, I_mass_inert, I_mass_in_tot, thick As Double
        Dim no_beaters, B, H, H2, tip_width As Double
        Dim tb, th, I_missing_tip, tip_weight As Double

        '-------- mass moment of --------------
        'see http://www.dtic.mil/dtic/tr/fulltext/u2/274936.pdf
        no_beaters = NumericUpDown7.Value
        overall_length = NumericUpDown8.Value / 1000
        H = overall_length / 2     '[m]
        B = NumericUpDown21.Value / 1000
        thick = NumericUpDown9.Value / 1000
        tip_width = NumericUpDown22.Value / 1000

        '---- top triangle is cut off------------
        th = tip_width / B * H          'missing_tip_height
        tb = tip_width                  'missing_tip_base
        tip_weight = th * tb * thick * 7850 / 2    '[kg] (triangle)
        I_missing_tip = (half_beater_weight / 24) * ((7 * tb ^ 2) + (4 * th ^ 2)) '[kg.m2] one triangle

        '---- Beater triangle including missing tip ---------
        H2 = H + th
        half_beater_weight = H2 * B * thick * 7850 / 2    '[kg] (triangle)
        I_mass_inert = (half_beater_weight / 24) * ((7 * B ^ 2) + (4 * H2 ^ 2)) '[kg.m2] one triangle

        '---- now subtract the missing tio
        I_mass_inert = I_mass_inert - I_missing_tip
        I_mass_inert *= 2                                   'two triangles is one beater
        half_beater_weight = half_beater_weight - tip_weight
        I_mass_in_tot = I_mass_inert * no_beaters '[kg.m2]


        '----present--------
        TextBox26.Text = I_mass_inert.ToString("0")                   '[kg.m2]
        TextBox27.Text = I_mass_in_tot.ToString("0")                  '[kg.m2]
        TextBox28.Text = (half_beater_weight * 2).ToString("0")     '[kg]
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Write_to_word()
    End Sub

    Private Sub Write_to_word()

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim row As Integer
        Dim ufilename As String

        Try
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            oDoc.PageSetup.TopMargin = 35
            oDoc.PageSetup.BottomMargin = 20
            oDoc.PageSetup.RightMargin = 20
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait
            oDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
            'oDoc.PageSetup.VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalCenter

            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = 14
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 0.5                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Disintegrator stress calculation" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            oTable.Cell(1, 1).Range.Text = "Project Name"
            oTable.Cell(1, 2).Range.Text = TextBox30.Text
            oTable.Cell(2, 1).Range.Text = "Item number"
            oTable.Cell(2, 2).Range.Text = TextBox31.Text
            oTable.Cell(3, 1).Range.Text = "Author "
            oTable.Cell(3, 2).Range.Text = Environment.UserName
            oTable.Cell(4, 1).Range.Text = "Date "
            oTable.Cell(4, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Drive Details----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 6, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Electric motor "
            row += 1
            oTable.Cell(row, 1).Range.Text = "Installed Power"
            oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[kW]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Speed "
            oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Radial speed"
            oTable.Cell(row, 2).Range.Text = TextBox2.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Nominal Motor torque"
            oTable.Cell(row, 2).Range.Text = TextBox3.Text
            oTable.Cell(row, 3).Range.Text = "[kNm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Locked Motor torque"
            oTable.Cell(row, 2).Range.Text = TextBox29.Text
            oTable.Cell(row, 3).Range.Text = "[kNm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ material----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Material"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Wet material"
            oTable.Cell(row, 2).Range.Text = NumericUpDown3.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[ton/hr]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Dry material"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[ton/hr]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total material"
            oTable.Cell(row, 2).Range.Text = TextBox1.Text
            oTable.Cell(row, 3).Range.Text = "[kg/sn]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Density"
            oTable.Cell(row, 2).Range.Text = NumericUpDown16.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Coupling key----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Coupling key"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown13.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key size"
            oTable.Cell(row, 2).Range.Text = TextBox18.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "key length"
            oTable.Cell(row, 2).Range.Text = TextBox17.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Allowed stress"
            oTable.Cell(row, 2).Range.Text = TextBox19.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key t1"
            oTable.Cell(row, 2).Range.Text = TextBox21.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max force 1 key"
            oTable.Cell(row, 2).Range.Text = TextBox22.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max Allowed power 1 key"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[kW]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Safety factor (locked motor)"
            oTable.Cell(row, 2).Range.Text = TextBox32.Text
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Beaters----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 11, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Beaters"
            row += 1
            oTable.Cell(row, 1).Range.Text = "No of beaters"
            oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Overall length"
            oTable.Cell(row, 2).Range.Text = NumericUpDown8.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Center width"
            oTable.Cell(row, 2).Range.Text = NumericUpDown21.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Tip width"
            oTable.Cell(row, 2).Range.Text = NumericUpDown22.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Beater plate thickness"
            oTable.Cell(row, 2).Range.Text = NumericUpDown9.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction disk OD"
            oTable.Cell(row, 2).Range.Text = TextBox37.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction disk thckness"
            oTable.Cell(row, 2).Range.Text = NumericUpDown11.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Tip speed"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Beater weight"
            oTable.Cell(row, 2).Range.Text = TextBox28.Text
            oTable.Cell(row, 3).Range.Text = "[kg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Beater inertia"
            oTable.Cell(row, 2).Range.Text = TextBox26.Text
            oTable.Cell(row, 3).Range.Text = "[kg.m2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total inertia"
            oTable.Cell(row, 2).Range.Text = TextBox27.Text
            oTable.Cell(row, 3).Range.Text = "[kg.m2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Beaters shaft ----------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Beater shaft key"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown12.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key size"
            oTable.Cell(row, 2).Range.Text = TextBox14.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "key length"
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Allowed stress"
            oTable.Cell(row, 2).Range.Text = TextBox12.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key t1"
            oTable.Cell(row, 2).Range.Text = TextBox13.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max force 1 key"
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max Torque 2 keys"
            oTable.Cell(row, 2).Range.Text = TextBox9.Text
            oTable.Cell(row, 3).Range.Text = "[kN.m]"


            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Material lump ----------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Material lump (Teufelsei)"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Egg diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown14.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Egg weight"
            oTable.Cell(row, 2).Range.Text = TextBox5.Text
            oTable.Cell(row, 3).Range.Text = "[kg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Acceleration time"
            oTable.Cell(row, 2).Range.Text = NumericUpDown15.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[sec]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Acceleration"
            oTable.Cell(row, 2).Range.Text = TextBox6.Text
            oTable.Cell(row, 3).Range.Text = "[m/s2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Acceleration force"
            oTable.Cell(row, 2).Range.Text = TextBox7.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Generated torque"
            oTable.Cell(row, 2).Range.Text = TextBox10.Text
            oTable.Cell(row, 3).Range.Text = "[kNm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Safety factor"
            oTable.Cell(row, 2).Range.Text = TextBox40.Text
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Hydaulic Nut ----------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Hydraulic Nut"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction disk OD"
            oTable.Cell(row, 2).Range.Text = NumericUpDown21.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction disk ID"
            oTable.Cell(row, 2).Range.Text = TextBox23.Text
            oTable.Cell(row, 3).Range.Text = "[kg]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction radius"
            oTable.Cell(row, 2).Range.Text = TextBox24.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Friction coef."
            oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Strech force"
            oTable.Cell(row, 2).Range.Text = NumericUpDown19.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[KN]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max. allowable torque"
            oTable.Cell(row, 2).Range.Text = TextBox25.Text
            oTable.Cell(row, 3).Range.Text = "[kNm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Stretch length"
            oTable.Cell(row, 2).Range.Text = TextBox34.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------- store rapport------------------
            ufilename = "Fan_select_report_" & TextBox30.Text & "_" & TextBox31.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".docx"

            '---- if path not exist then create one----------
            Try
                If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
                If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
                If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
            Catch ex As Exception
            End Try

            If Directory.Exists(dirpath_Rap) Then
                ufilename = dirpath_Rap & ufilename
            Else
                ufilename = dirpath_Home & ufilename
            End If
            oWord.ActiveDocument.SaveAs(ufilename.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " Problem storing file to " & dirpath_Rap)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click, TabPage2.Enter
        Calc_inertia()
    End Sub
    Private Sub Calc_shaft_coupling()
        Dim dia, dia_calc, t1, f1, m1 As Double
        Dim J, area As Double
        Dim τ, σ_allowed As Double

        dia = NumericUpDown13.Value         'dia shaft
        Double.TryParse(TextBox21.Text, t1) 'depth key
        dia_calc = dia - t1                 'shaft calculation diameter
        f1 = 0                              'pulling force
        Double.TryParse(TextBox29.Text, m1) 'torque locked motor
        m1 *= 10 ^ 6                        '[kN.m]-->[N.mm]
        σ_allowed = NumericUpDown10.Value / 1.5 / 1.25 'σ_allowed

        '--------- Calc Polar Moment of Inertia of Area   -----------
        'https://www.engineeringtoolbox.com/torsion-shafts-d_947.html
        J = (PI * dia_calc ^ 4) / 32    '[mm4] Solid shaft

        '--------- calc τ -----------
        τ = m1 * (dia_calc / 2) / J

        '---------- present --------------
        TextBox41.Text = dia.ToString("0.0")
        TextBox42.Text = t1.ToString("0.0")
        TextBox43.Text = dia_calc.ToString("0.0")
        TextBox44.Text = (m1 / 10 ^ 6).ToString("0")    '[kNm]
        TextBox45.Text = f1.ToString("0")
        TextBox46.Text = τ.ToString("0")
        TextBox48.Text = σ_allowed.ToString("0")  'allowed stress
        TextBox60.Text = J.ToString("0")
        TextBox63.Text = area.ToString("0")

        '--------- checks ---------
        TextBox46.BackColor = IIf(τ < σ_allowed, Color.LightGreen, Color.Red)
    End Sub
    Private Sub Calc_shaft_beaters()
        Dim dia, dia_calc, t1, f1, m1, m2, length As Double
        Dim J, I, area As Double
        Dim σd, σb, τ, σ2, σ_allowed As Double
        Dim dia_fric As Double
        Dim wght, q As Double

        dia = NumericUpDown12.Value         '[mm] dia shaft
        Double.TryParse(TextBox13.Text, t1) '[mm] depth key
        dia_calc = dia - t1                 '[mm]shaft calculation diameter
        f1 = NumericUpDown19.Value * 10 ^ 3 '[N]pulling force
        Double.TryParse(TextBox29.Text, m1) 'torque locked motor
        m1 *= 10 ^ 6                        '[kN.m]-->[N.mm]
        length = NumericUpDown8.Value       '[mm] bearing-bearing length
        dia_fric = NumericUpDown21.Value    '[mm] spacer plate
        σ_allowed = NumericUpDown10.Value / 1.5 / 1.25 'σ_allowed
        wght = NumericUpDown5.Value         '[kg]

        '--------- Calc Polar Moment of Inertia   -----------
        'https://www.engineeringtoolbox.com/torsion-shafts-d_947.html
        J = PI * dia_calc ^ 4 / 32    '[mm4] Solid shaft

        '--------- Calc Area Moment of Inertia    -----------
        I = PI * dia_calc ^ 4 / 64    '[mm4] Solid shaft

        '--------- calc σd (pull force) -----------
        area = PI / 4 * dia_calc ^ 2    '[mm2]
        σd = f1 / area

        '--------- calc σb (bend force) -----------
        ' http://www-classes.usc.edu/engr/ce/457/moment_table.pdf
        'Uniform load 
        q = wght * 10 / length              '[n/mm] uniform load
        m2 = q * length ^ 2 / 8             '[Nmm] Max moment
        σb = m2 * (dia_calc / 2) / I

        '--------- calc τ -----------
        τ = m1 * (dia_calc / 2) / J

        '--------- calc combined stress -----------
        σ2 = 0.5 * Sqrt(((σd - σb) / 2) ^ 2 + 4 * τ ^ 2)

        '---------- present --------------
        TextBox47.Text = m2.ToString("0")               '[Nm]
        TextBox56.Text = dia.ToString("0.0")            '[mm]
        TextBox57.Text = t1.ToString("0.0")             '[mm]
        TextBox58.Text = dia_calc.ToString("0.0")       '[mm]
        TextBox53.Text = (m1 / 10 ^ 6).ToString("0")    '[kNm]
        TextBox54.Text = (f1 / 1000).ToString("0")      '[kN]
        TextBox55.Text = length.ToString("0.0")
        TextBox51.Text = τ.ToString("0")
        TextBox52.Text = σd.ToString("0")
        TextBox59.Text = σb.ToString("0")
        TextBox50.Text = σ2.ToString("0")
        TextBox61.Text = J.ToString("0")
        TextBox62.Text = I.ToString("0")
        TextBox64.Text = dia_fric.ToString("0")
        '--------- checks ---------
        TextBox50.BackColor = IIf(σ2 < σ_allowed, Color.LightGreen, Color.Red)
    End Sub

End Class