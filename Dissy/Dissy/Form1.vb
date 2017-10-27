Imports System.IO
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

    Private Sub Button1_Click(sender As Object, E As EventArgs) Handles Button1.Click, TabPage1.Enter, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown1.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown17.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, ComboBox1.SelectedIndexChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown18.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged
        Calc_tab1()
    End Sub

    Private Sub Calc_tab1()
        Dim tot_Instal_power, rpm, rad, motor_torque, dia_beater As Double
        Dim l_wet, l_add, l_tot As Double
        Dim tip_speed, acc, acc_time As Double
        Dim lump_dia, lump_weight, density, f_tip, lump_torque As Double
        Dim FOS_lump_key, FOS_lump_nut As Double
        Dim key_h, key_l, beater_shaft_radius, max_key_torque, max_key_force As Double
        Dim start_torque, service_factor As Double
        Dim σ_yield, allowed_σ_stress, allowed_τ_stress As Double
        Dim specific_load, load_beater_tip As Double
        Dim no_beaters, actual_egg_key_force As Double
        Dim drive_key_radius As Double
        Dim no_motors As Double

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
            TextBox18.Text = words(2)       'Key sizeallowed_σ_stress
            TextBox21.Text = words(3)       '(t1) Key depth in shaft
        End If

        no_motors = NumericUpDown27.Value
        service_factor = NumericUpDown24.Value
        no_beaters = NumericUpDown7.Value
        Double.TryParse(TextBox13.Text, key_h)      '[mm]
        key_l = NumericUpDown9.Value                '[mm]

        tot_Instal_power = NumericUpDown1.Value * 1000    '[kW]
        tot_Instal_power *= no_motors
        rpm = NumericUpDown2.Value
        dia_beater = NumericUpDown8.Value / 1000    '[m]
        lump_dia = NumericUpDown14.Value / 1000     '[m]
        drive_key_radius = NumericUpDown13.Value / 2000     '[m]
        acc_time = NumericUpDown15.Value
        density = NumericUpDown16.Value
        σ_yield = NumericUpDown18.Value

        beater_shaft_radius = NumericUpDown12.Value / 2000   '[mm]
        allowed_σ_stress = σ_yield / service_factor  '[N/mm2]
        allowed_τ_stress = allowed_σ_stress * 0.8   '[N/mm2]

        '-------- motor----------
        rad = rpm / 60 * 2 * PI
        motor_torque = tot_Instal_power / rad
        start_torque = motor_torque * 2.0
        tip_speed = dia_beater * rpm * PI / 60  '[m/s]

        '-------- process ------
        l_wet = NumericUpDown3.Value
        l_add = NumericUpDown4.Value
        l_tot = (l_add + l_wet) * 1000 / 3600   '[kg/s]
        specific_load = 3600 * l_tot / tot_Instal_power  '[ton/(kW.hr)]
        load_beater_tip = l_tot / (rpm / 60 * no_beaters * 2)


        '---- Lump calculation--------
        lump_weight = 4 / 3 * PI * (lump_dia / 2) ^ 3 * density

        acc = tip_speed / acc_time              '[m/s2]
        f_tip = lump_weight * acc               '[N]
        lump_torque = f_tip * (dia_beater / 2)  '[N.m]

        actual_egg_key_force = lump_torque / beater_shaft_radius
        actual_egg_key_force /= 2    'two keys 

        max_key_force = key_h * key_l * allowed_σ_stress
        max_key_torque = max_key_force * beater_shaft_radius
        max_key_torque *= 2     'two keys 

        FOS_lump_key = max_key_torque / lump_torque

        '--------- max Allowed power on coupling key --
        Dim actual_drive_key_force, drive_l, drive_b As Double
        Dim actual_drive_key_stress, FOS_coupling_key As Double

        Double.TryParse(TextBox21.Text, drive_b)            '[mm] key t1
        drive_l = NumericUpDown17.Value                     '[mm] key length

        actual_drive_key_force = start_torque / drive_key_radius / no_motors      '[kN]
        actual_drive_key_stress = actual_drive_key_force / (drive_b * drive_l)   '[N/mm2]
        FOS_coupling_key = σ_yield / actual_drive_key_stress      '[-]

        '--------- Hydraulic nut (spacer = friction disk) --
        Dim spacer_od, spacer_id, spacer_radius, fric As Double
        Dim max_torque_nut, delta_l, shaft_l, pull_force, area As Double
        pull_force = NumericUpDown19.Value
        spacer_od = NumericUpDown21.Value
        fric = NumericUpDown6.Value
        spacer_id = NumericUpDown12.Value
        spacer_radius = (spacer_od + spacer_id) / 4
        max_torque_nut = pull_force * fric * (spacer_radius / 1000)     '[kNm]

        area = PI / 4 * spacer_id ^ 2                                   '[mm2]
        shaft_l = NumericUpDown28.Value                                 '[mm] stretch indicator
        delta_l = pull_force * 10 ^ 3 * shaft_l / (215000 * area)       '[mm]
        FOS_lump_nut = max_torque_nut * 10 ^ 3 / lump_torque

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
        TextBox12.Text = allowed_σ_stress.ToString("0")         '[N/mm2]
        TextBox19.Text = allowed_σ_stress.ToString("0")         '[N/mm2]
        TextBox20.Text = (actual_drive_key_stress).ToString("0") '[kNm]
        TextBox22.Text = (actual_drive_key_force / 10 ^ 3).ToString("0") '[km]
        TextBox23.Text = spacer_id.ToString("0")                '[mm]
        TextBox24.Text = spacer_radius.ToString("0")            '[mm]
        TextBox25.Text = max_torque_nut.ToString("0")           '[kNm]
        TextBox29.Text = (start_torque / 1000).ToString("0.0")  '[kNm]
        TextBox32.Text = FOS_coupling_key.ToString("0.0")       '[-]
        TextBox33.Text = specific_load.ToString("0.00")         '[]
        TextBox34.Text = delta_l.ToString("0.00")               '[mm]
        TextBox35.Text = (tot_Instal_power / 1000).ToString("0") '[kW]
        TextBox36.Text = load_beater_tip.ToString("0.00")       '[kg]
        TextBox37.Text = spacer_od.ToString("0")                '[mm]
        TextBox40.Text = FOS_lump_key.ToString("0.0")           '[-]
        TextBox65.Text = FOS_lump_nut.ToString("0.0")           '[-]
        TextBox67.Text = allowed_τ_stress.ToString("0")         '[N/mm2]
        TextBox70.Text = (actual_egg_key_force / 10 ^ 3).ToString("0.0")     '[kN]


        '------- checks---------
        TextBox32.BackColor = IIf(FOS_coupling_key > 3, Color.LightGreen, Color.Red)
        TextBox40.BackColor = IIf(FOS_lump_key > 3, Color.LightGreen, Color.Red)
        TextBox65.BackColor = IIf(FOS_lump_nut > 3, Color.LightGreen, Color.Red)
        Calc_inertia()
        Calc_shaft_coupling()
        Calc_shaft()
        Calc_beater()
        Calc_emotor_4P()
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

        TextBox69.Text =
        "Factors of Design and Safety" & vbCrLf &
        "FOS= Yield stress/Working stress" & vbCrLf &
        "FOS must be bigger than 3.0" & vbCrLf &
        "Design load is maximum load of part the part will ever see in service" & vbCrLf &
        "Note Yield not Ultimate strength is used" & vbCrLf
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
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Electric motor "
            row += 1
            oTable.Cell(row, 1).Range.Text = "Total installed Power"
            oTable.Cell(row, 2).Range.Text = TextBox35.Text
            oTable.Cell(row, 3).Range.Text = "[kW]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "No. motors"
            oTable.Cell(row, 2).Range.Text = NumericUpDown27.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Speed "
            oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[rpm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Radial speed"
            oTable.Cell(row, 2).Range.Text = TextBox2.Text
            oTable.Cell(row, 3).Range.Text = "[rad/s]"
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

            '------------------ Selected Steel ----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Selected steel"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key,σ design compression stress"
            oTable.Cell(row, 2).Range.Text = TextBox12.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key,τ design shear stress"
            oTable.Cell(row, 2).Range.Text = TextBox67.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft,σ design compr.stress"
            oTable.Cell(row, 2).Range.Text = TextBox48.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft,τ design shear stress"
            oTable.Cell(row, 2).Range.Text = TextBox68.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"


            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ material----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
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
            oTable.Cell(row, 3).Range.Text = "[kg/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Average density"
            oTable.Cell(row, 2).Range.Text = NumericUpDown16.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Specific  load"
            oTable.Cell(row, 2).Range.Text = TextBox33.Text
            oTable.Cell(row, 3).Range.Text = "[ton/kw]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Specific tip load"
            oTable.Cell(row, 2).Range.Text = TextBox36.Text
            oTable.Cell(row, 3).Range.Text = "[kg]"

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
            oTable.Cell(row, 1).Range.Text = "Key length"
            oTable.Cell(row, 2).Range.Text = TextBox17.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Design σ stress"
            oTable.Cell(row, 2).Range.Text = TextBox19.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key t1"
            oTable.Cell(row, 2).Range.Text = TextBox21.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Actual force 1 key"
            oTable.Cell(row, 2).Range.Text = TextBox22.Text
            oTable.Cell(row, 3).Range.Text = "[kN]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Actual key stress"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Factor of Safety (locked motor)"
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
            oTable.Cell(row, 1).Range.Text = "Overall beater diameter"
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
            oTable.Cell(row, 1).Range.Text = "Friction disk thickness"
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
            oTable.Cell(row, 1).Range.Text = "Working key length"
            oTable.Cell(row, 2).Range.Text = TextBox8.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Allowed σ compression stress"
            oTable.Cell(row, 2).Range.Text = TextBox12.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key t1"
            oTable.Cell(row, 2).Range.Text = TextBox13.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Maximum force 1 key"
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            oTable.Cell(row, 3).Range.Text = "[N]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Maximum Torque 2 keys"
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
            oTable.Cell(row, 1).Range.Text = "Factor of Safety, key only"
            oTable.Cell(row, 2).Range.Text = TextBox40.Text
            oTable.Cell(row, 3).Range.Text = "[-]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Factor of Safety, hydraulic nut"
            oTable.Cell(row, 2).Range.Text = TextBox65.Text
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Hydraulic Nut ----------------------
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
            oTable.Cell(row, 3).Range.Text = "[mm]"
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

            '------------------ Shaft ----------------------
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 13, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Shaft"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft drive diameter"
            oTable.Cell(row, 2).Range.Text = TextBox41.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key depth t1"
            oTable.Cell(row, 2).Range.Text = TextBox42.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Calc diameter"
            oTable.Cell(row, 2).Range.Text = TextBox43.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Drive key τ stress"
            oTable.Cell(row, 2).Range.Text = TextBox46.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft beater diameter"
            oTable.Cell(row, 2).Range.Text = TextBox56.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Key depth"
            oTable.Cell(row, 2).Range.Text = TextBox57.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Calc diameter"
            oTable.Cell(row, 2).Range.Text = TextBox58.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Torsion τ stress"
            oTable.Cell(row, 2).Range.Text = TextBox51.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Pull σd stress"
            oTable.Cell(row, 2).Range.Text = TextBox52.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bend σb stress"
            oTable.Cell(row, 2).Range.Text = TextBox59.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Principle combined stress"
            oTable.Cell(row, 2).Range.Text = TextBox50.Text
            oTable.Cell(row, 3).Range.Text = "[N/mm2]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Factor of Safety stress"
            oTable.Cell(row, 2).Range.Text = TextBox45.Text
            oTable.Cell(row, 3).Range.Text = "[-]"

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
        Dim dia, dia_calc, t1, m1 As Double
        Dim j, τ, σ_design_cpl As Double

        dia = NumericUpDown13.Value         'dia shaft
        Double.TryParse(TextBox21.Text, t1) 'depth key
        dia_calc = dia - t1                 'shaft calculation diameter
        Double.TryParse(TextBox29.Text, m1) 'torque locked motor
        m1 *= 10 ^ 6                        '[kN.m]-->[N.mm]
        Double.TryParse(TextBox12.Text, σ_design_cpl) 'desing stress


        '--------- Calc Polar Moment of Inertia of Area   -----------
        'https://www.engineeringtoolbox.com/torsion-shafts-d_947.html
        j = (PI * dia_calc ^ 4) / 32    '[mm4] Solid shaft

        '--------- calc τ -----------
        τ = m1 * (dia_calc / 2) / J

        '---------- present --------------
        TextBox41.Text = dia.ToString("0.0")
        TextBox42.Text = t1.ToString("0.0")
        TextBox43.Text = dia_calc.ToString("0.0")
        TextBox44.Text = (m1 / 10 ^ 3).ToString("0")    '[Nm]
        TextBox46.Text = τ.ToString("0")


        TextBox60.Text = J.ToString("0")
        '--------- checks ---------
        TextBox46.BackColor = IIf(τ < σ_design_cpl, Color.LightGreen, Color.Red)
    End Sub
    Private Sub Calc_shaft()
        Dim dia, dia_calc, t1, f1, dm1, dm2 As Double
        Dim length_l, length_a, length_b, R1 As Double
        Dim J, I, area As Double
        Dim σd, σb, τ, FOS_stress As Double
        Dim σ_design_shft, σ_yield, τ_design_shft, service_fac As Double
        Dim τmax, σ12 As Double
        Dim dia_fric As Double
        Dim wght, w As Double

        dia = NumericUpDown12.Value         '[mm] dia shaft
        Double.TryParse(TextBox13.Text, t1) '[mm] depth key
        dia_calc = dia - 2 * t1             '[mm]shaft calculation diameter
        f1 = NumericUpDown19.Value * 10 ^ 3 '[N] pulling force
        Double.TryParse(TextBox29.Text, dm1) 'torque locked motor
        dm1 *= 10 ^ 6                        '[kN.m]-->[N.mm]
        length_l = NumericUpDown29.Value    '[mm] bearing-bearing length
        length_b = NumericUpDown20.Value    '[mm] beater shaft key
        dia_fric = NumericUpDown21.Value    '[mm] spacer plate

        service_fac = NumericUpDown25.Value
        σ_yield = NumericUpDown10.Value
        σ_design_shft = σ_yield / service_fac 'σ_design
        τ_design_shft = σ_design_shft * 0.58
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
        'http://www-classes.usc.edu/engr/ce/457/moment_table.pdf
        'http://www.awc.org/pdf/codes-standards/publications/design-aids/AWC-DA6-BeamFormulas-0710.pdf
        'https://theconstructor.org/structural-engg/solid-mechanics/combined-bending-direct-and-torsional-stresses/3704/
        'Simple support with Partial Uniform load 

        w = wght * 9.81 / length_b         '[N/mm] uniform load
        length_a = (length_l - length_b) / 2
        R1 = wght * 9.81 / 2               'R1=R2= (half weight * 9.81)
        dm2 = R1 * (length_a + R1 / (2 * w))

        σb = dm2 * (dia_calc / 2) / I

        '--------- calc τ -----------
        τ = dm1 * (dia_calc / 2) / J

        '--------- calc combined principle stress -----------
        '---- Stress and Strain formula (2.3-23)-------------
        τmax = 0.5 * Sqrt(((σd - σb) * 0.5) ^ 2 + 4 * τ ^ 2)
        σ12 = ((σd + σb) * 0.5) + Sqrt(((σd - σb) * 0.5) ^ 2 + τ ^ 2)  'max priciple stress

        FOS_stress = σ_yield / σ12

        '---------- present --------------
        TextBox47.Text = (dm2 / 1000).ToString("0")     '[Nm]
        TextBox56.Text = dia.ToString("0.0")            '[mm]
        TextBox57.Text = t1.ToString("0.0")             '[mm]
        TextBox58.Text = dia_calc.ToString("0.0")       '[mm]
        TextBox53.Text = (dm1 / 10 ^ 3).ToString("0")   '[Nm]
        TextBox54.Text = (f1 / 1000).ToString("0")      '[kN]
        TextBox51.Text = τ.ToString("0")
        TextBox52.Text = σd.ToString("0")
        TextBox59.Text = σb.ToString("0")
        TextBox50.Text = σ12.ToString("0")
        TextBox61.Text = J.ToString("0")
        TextBox62.Text = I.ToString("0")
        TextBox63.Text = area.ToString("0")
        TextBox64.Text = dia_fric.ToString("0")
        TextBox45.Text = FOS_stress.ToString("0.0")
        TextBox49.Text = τmax.ToString("0.0")

        TextBox48.Text = σ_design_shft.ToString("0")  'allowed stress
        TextBox68.Text = τ_design_shft.ToString("0")  'allowed stress

        '--------- checks ---------
        TextBox49.BackColor = IIf(τmax < τ_design_shft, Color.LightGreen, Color.Red)
        TextBox50.BackColor = IIf(σ12 < σ_design_shft, Color.LightGreen, Color.Red)
        TextBox52.BackColor = IIf(σd < σ_design_shft, Color.LightGreen, Color.Red)
        TextBox51.BackColor = IIf(τ < τ_design_shft, Color.LightGreen, Color.Red)
        TextBox59.BackColor = IIf(σb < σ_design_shft, Color.LightGreen, Color.Red)
        TextBox45.BackColor = IIf(FOS_stress > 3.0, Color.LightGreen, Color.Red)
        '--------- beater shaft/key length input wrong--------- 
        If (length_l < length_b) Then
            NumericUpDown8.BackColor = Color.Red
            NumericUpDown20.BackColor = Color.Red
        Else
            NumericUpDown8.BackColor = Color.Yellow
            NumericUpDown20.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Calc_beater()
        Dim service_fac, σ_design_beat As Double
        service_fac = NumericUpDown26.Value

        σ_design_beat = NumericUpDown10.Value / service_fac 'σ_design
        TextBox66.Text = σ_design_beat.ToString("0")  'allowed stress
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox30.Text.Length > 0 And TextBox31.Text.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Naam en of Item Tag" & vbCrLf & "Is niet ingevuld !")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Read_file()
    End Sub
    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string As String
        Dim filename As String = "Dissy_calc_" & TextBox30.Text & "_" & TextBox31.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".vtk"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox10.Text) Then TextBox10.Text = "-"
        If String.IsNullOrEmpty(TextBox11.Text) Then TextBox11.Text = "-"

        temp_string = TextBox30.Text & ";" & TextBox31.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric, combobox, checkbox and radiobutton controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 5062, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function
    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim ttt As Double
        Dim k As Integer = 0
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Dissy_*"
        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK Files|*.vtk"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split first line the read file content
            TextBox30.Text = words(0)                  'Project number
            TextBox31.Text = words(1)                 'Item name

            '---------- terugzetten numeric controls -----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal numeric controls--
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        grbx.Value = grbx.Minimum       'NOK
                        MessageBox.Show("Numeric controls value out of ousode min-max range, Minimum value is used")
                    End If
                Else
                    MessageBox.Show("Warning last Numeric controls not found in file")  'NOK
                End If
            Next

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As CheckBox = CType(all_check(i), CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last radiobutton not found in file")
                End If
            Next
        End If
    End Sub
    Private Sub Calc_emotor_4P()        '
        Dim Ins_power, req_pow_safety, aanlooptijd, n_motor, rad As Double
        Dim m_torque_inrush, m_torque_max, m_torque_rated, m_torque_average As Double
        Dim impeller_inertia, motor_inertia, total_inertia As Double
        Dim ang_acceleration, C_acc, inertia_torque, fan_load_torque As Double

        '--------- motor torque-------------
        'see http://ecatalog.weg.net/files/wegnet/WEG-specification-of-electric-motors-50039409-manual-english.pdf
        'see http://electrical-engineering-portal.com/calculation-of-motor-startin

        Ins_power = NumericUpDown1.Value * 10 ^ 3           'Geinstalleerd vermogen [Watt]
        n_motor = NumericUpDown2.Value
        rad = n_motor / 60 * 2 * PI                        'Hoeksnelheid [rad/s]
        fan_load_torque = req_pow_safety / rad             '[N.m]
        m_torque_rated = Ins_power / rad
        m_torque_inrush = m_torque_rated * 0.95
        m_torque_max = m_torque_rated * 2.5

        m_torque_max *= 0.8 ^ 2                                     'Starting voltage is 80%
        m_torque_average = 0.45 * (m_torque_inrush + m_torque_max)  'Average torque motor

        '------------- inertia motor + inertia impeller-------------------
        motor_inertia = CDbl(Emotor_4P_inert(n_motor, Ins_power))
        Double.TryParse(TextBox27.Text, impeller_inertia)       '[kg.m2]
        total_inertia = impeller_inertia + motor_inertia        '[kg.m2]
        inertia_torque = total_inertia * ang_acceleration       '[N.m]

        '-------------- aanlooptijd--------------------------------
        C_acc = m_torque_average - (2.5 * fan_load_torque)
        aanlooptijd = 2 * PI * n_motor * total_inertia / (60 * C_acc)
        TextBox39.Text = aanlooptijd.ToString("0") 'Aanlooptijd [s]
    End Sub
    ' see http://ecatalog.weg.net/files/wegnet/WEG-specification-of-electric-motors-50039409-manual-english.pdf
    Function Emotor_4P_inert(rpm As Double, kw As Double) As Double
        Dim motor_inertia As Double
        If rpm < 750 Then rpm = 750
        Select Case True
            Case rpm = 3000
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 1 ^ 2.5    '2 poles (1 pair) (3000 rpm) [kg.m2]
            Case rpm = 1500
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 2 ^ 2.5    '4 poles (2 pair) (1500 rpm) [kg.m2]
            Case rpm = 1000
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 3 ^ 2.5    '6 poles (3 pair) (1000 rpm) [kg.m2]
            Case rpm = 750
                motor_inertia = 0.04 * (kw / 1000) ^ 0.9 * 4 ^ 2.5    '8 poles (4 pair) (750 rpm) [kg.m2]
            Case Else
                MessageBox.Show("Error occured in Motor Inertia calculation ")
        End Select
        Return (motor_inertia)
    End Function

End Class
