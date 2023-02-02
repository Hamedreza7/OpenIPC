'This codes are written by Dr. Hamed Hosseinzadeh (2021 and 2022) in Prof. Lang Yuan's Research group at The UofSc .
'The codes are available through MIT License.
'If you change the codes, please distribute the updated codes as MIT License.
'Please do not sell this version or your updated version!
'Please do not hesitate to email: Hamed@uwalumni.com, if you have any questions.


Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports System.Windows.Forms.DataVisualization.Charting
Imports Microsoft.Office.Tools
Imports System.IO

Public Class MDIParent1

    Dim Pore_Pixel(0 To 5000, 0 To 5000) As Boolean
    Dim Crack_Pixel(0 To 5000, 0 To 5000) As Boolean
    Dim Crack_Pixel_checked(0 To 5000, 0 To 5000) As Boolean
    Dim Crack_Pixel_length(0 To 5000) As Integer
    Dim Crack_Pixel_X(0 To 5000) As Integer
    Dim Crack_Pixel_Y(0 To 5000) As Integer
    Dim Crack_Pixel_Color(0 To 5000) As Color
    Dim Crack_No As Integer
    Dim Black_test_Pixel(0 To 5000, 0 To 5000) As Boolean

    Dim Pore_No(0 To 5000, 0 To 5000) As Integer
    Dim Pore_Boundary(0 To 5000, 0 To 5000) As Boolean
    Dim Pore_Boundary_no(0 To 5000, 0 To 5000) As Integer
    Dim Pore_Perimeter(0 To 10000) As Single
    Dim Pore_Color(0 To 5000, 0 To 5000) As Color
    Dim Pore_Area(0 To 10000) As Single

    Dim TopLeft_X As Integer
    Dim TopLeft_Y As Integer
    Dim BottomRight_X As Integer
    Dim BottomRight_Y As Integer

    Dim _painter As Boolean
    Dim painter As Boolean

    Dim Color_R As Integer
    Dim Color_G As Integer
    Dim Color_B As Integer

    Dim Color_R_y As Integer
    Dim Color_G_y As Integer
    Dim Color_B_y As Integer

    Dim Color_R_x As Integer
    Dim Color_G_x As Integer
    Dim Color_B_x As Integer

    Dim Color_R_i As Integer
    Dim Color_G_i As Integer
    Dim Color_B_i As Integer

    Dim Layer As Integer
    Dim Grain_area As Integer
    Dim Grain_width_size As Integer
    Dim Grain_height_size As Integer
    Dim Grain_numbers As Integer
    Dim Grain_number As Integer
    Dim Grain_numbers_layer(0 To 1000) As Integer
    Dim Grain_size(0 To 1000, 0 To 1000) As String
    Dim Grain_boundary_color(0 To 10000, 0 To 10000) As Boolean
    Dim Grain(0 To 10000, 0 To 10000) As Integer
    Dim Grain_height(0 To 10000) As Single
    Dim Grain_width(0 To 10000) As Single
    Dim Counter_i(0 To 10000) As Integer

    Dim width_start, width_end, height_start, height_end As Integer
    Dim Calibration_start, Calibration_end As Boolean
    Dim Calibration_start_x, Calibration_end_x As Integer
    Dim pixel_size As Single
    Dim Region_start, Region_end As Boolean

    Dim Start_grain As Boolean
    Dim End_grain As Boolean

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub



    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        TextBox5.Text = TrackBar1.Value
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim ofd As New OpenFileDialog
        ofd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
        ofd.Filter = "JPEG files (*.jpg)|*.jpg|Bitmap files (*.bmp)|*.bmp|TIF files (*.tif)|*.tif"
        Dim result As DialogResult = ofd.ShowDialog
        If Not (PictureBox2) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox2.Image = Image.FromFile(ofd.FileName)
        End If
        If Not (PictureBox1) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox1.Image = Image.FromFile(ofd.FileName)
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If TSDDB1.Text = "Black & White" Then
            If CheckBox4.Checked = True Then
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                Dim MyBitmap As Bitmap
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    Layer += 1
                    TSCB1.Items.Add(Layer)
                    Start_grain = False
                    End_grain = False
                    Grain_width_size = 0
                    Grain_numbers = 0
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                            If Start_grain = False Then
                                Start_grain = True
                                Grain_numbers = 1
                            End If
                        End If
                        If Start_grain = True Then
                            If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                MyBitmap.SetPixel(i, j, Color.Green)
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                Grain_width_size += 1
                                If i = BottomRight_X Then
                                    Grain_numbers = Grain_numbers - 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                End If
                            End If
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                For x = i - 2 To i + 2
                                    For y = j - 2 To j + 2
                                        MyBitmap.SetPixel(x, y, Color.Red)
                                    Next
                                Next
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                If Grain_width_size > Val(TextBox5.Text) Then
                                    Grain_size(Layer, Grain_numbers) = Grain_width_size
                                    ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain" & Grain_numbers & "): " & Grain_width_size)
                                    Grain_numbers += 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                    Grain_width_size = 0
                                End If
                            End If

                        End If

                    Next
                Next

            End If
        End If
        If TSDDB1.Text = "Grayscale" Then
            TSCB1.Items.Clear()
            ListBox1.Items.Clear()
            Dim MyBitmap As Bitmap
            MyBitmap = PictureBox2.Image
            Layer = 0
            TPB1.Value = 0
            TPB1.Maximum = BottomRight_Y + 1
            For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                TPB1.Value = j
                For i = TopLeft_X To BottomRight_X
                    For x = i - 1 To i + 1
                        For y = j - 1 To j + 1
                            Color_R_i = MyBitmap.GetPixel(i, j).R
                            Color_G_i = MyBitmap.GetPixel(i, j).G
                            Color_B_i = MyBitmap.GetPixel(i, j).B
                            Color_R_x = MyBitmap.GetPixel(x, y).R
                            Color_G_x = MyBitmap.GetPixel(x, y).G
                            Color_B_x = MyBitmap.GetPixel(x, y).B

                            If MyBitmap.GetPixel(x, y) <> MyBitmap.GetPixel(i, j) Then
                                If Math.Abs(Color_R_x - Color_R_i) < 10 Then GoTo h2
                                If Math.Abs(Color_G_x - Color_G_i) < 10 Then GoTo h2
                                If Math.Abs(Color_B_x - Color_B_i) < 10 Then GoTo h2
                                Grain_boundary_color(x, y) = True
                                GoTo h3
                            End If
h2:
                        Next
                    Next
h3:
                Next
            Next
            PictureBox2.Image = MyBitmap
            PictureBox2.Refresh()
            TPB1.Value = 0
            For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                TPB1.Value = j
                For i = TopLeft_X To BottomRight_X
                    If Grain_boundary_color(i, j) = True Then
                        MyBitmap.SetPixel(i, j, Color.FromArgb(0, 0, 0))
                    Else
                        MyBitmap.SetPixel(i, j, Color.FromArgb(255, 255, 255))
                    End If
                    PictureBox2.Image = MyBitmap
                    PictureBox2.Refresh()
                Next
            Next
            If CheckBox4.Checked = True Then
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    Layer += 1
                    TSCB1.Items.Add(Layer)
                    Start_grain = False
                    End_grain = False
                    Grain_width_size = 0
                    Grain_numbers = 0
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                            If Start_grain = False Then
                                Start_grain = True
                                Grain_numbers = 1
                            End If
                        End If
                        If Start_grain = True Then
                            If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                MyBitmap.SetPixel(i, j, Color.Green)
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                Grain_width_size += 1
                                If i = BottomRight_X Then
                                    Grain_numbers = Grain_numbers - 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                End If
                            End If
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                For x = i - 2 To i + 2
                                    For y = j - 2 To j + 2
                                        MyBitmap.SetPixel(x, y, Color.Red)
                                    Next
                                Next
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                If Grain_width_size > Val(TextBox5.Text) Then
                                    Grain_size(Layer, Grain_numbers) = Grain_width_size
                                    ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain" & Grain_numbers & "): " & Grain_width_size)
                                    Grain_numbers += 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                    Grain_width_size = 0
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        End If
        If TSDDB1.Text = "Color" Then
            TSCB1.Items.Clear()
            ListBox1.Items.Clear()
            Dim MyBitmap As Bitmap
            Dim Color_gray As Integer
            MyBitmap = PictureBox2.Image
            Layer = 0
            TPB1.Value = 0
            TPB1.Maximum = BottomRight_Y + 1
            For j = TopLeft_Y To BottomRight_Y 'Step Val(TextBox4.Text)
                TPB1.Value = j
                For i = TopLeft_X To BottomRight_X
                    Color_R_i = MyBitmap.GetPixel(i, j).R
                    Color_G_i = MyBitmap.GetPixel(i, j).G
                    Color_B_i = MyBitmap.GetPixel(i, j).B
                    Color_gray = (Color_R_i + Color_G_i + Color_B_i) \ 3
                    MyBitmap.SetPixel(i, j, Color.FromArgb(Color_gray, Color_gray, Color_gray))
                Next
            Next
            PictureBox2.Image = MyBitmap
            PictureBox2.Refresh()
            For j = TopLeft_Y To BottomRight_Y 'Step Val(TextBox4.Text)
                TPB1.Value = j
                For i = TopLeft_X To BottomRight_X
                    For x = i - 1 To i + 1
                        For y = j - 1 To j + 1
                            Color_R_i = MyBitmap.GetPixel(i, j).R
                            Color_G_i = MyBitmap.GetPixel(i, j).G
                            Color_B_i = MyBitmap.GetPixel(i, j).B
                            Color_R_x = MyBitmap.GetPixel(x, y).R
                            Color_G_x = MyBitmap.GetPixel(x, y).G
                            Color_B_x = MyBitmap.GetPixel(x, y).B
                            If MyBitmap.GetPixel(x, y) <> MyBitmap.GetPixel(i, j) Then
                                If Math.Abs(Color_R_x - Color_R_i) < 10 Then GoTo h4
                                If Math.Abs(Color_G_x - Color_G_i) < 10 Then GoTo h4
                                If Math.Abs(Color_B_x - Color_B_i) < 10 Then GoTo h4
                                Grain_boundary_color(x, y) = True
                                GoTo h5
                            End If
h4:
                        Next
                    Next
h5:
                Next
            Next
            PictureBox2.Image = MyBitmap
            PictureBox2.Refresh()
            TPB1.Value = 0
            For j = TopLeft_Y To BottomRight_Y 'Step Val(TextBox4.Text)
                TPB1.Value = j
                For i = TopLeft_X To BottomRight_X
                    If Grain_boundary_color(i, j) = True Then
                        MyBitmap.SetPixel(i, j, Color.FromArgb(0, 0, 0))
                    Else
                        MyBitmap.SetPixel(i, j, Color.FromArgb(255, 255, 255))
                    End If
                Next
            Next
            PictureBox2.Image = MyBitmap
            PictureBox2.Refresh()
            If CheckBox4.Checked = True Then
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    Layer += 1
                    'TSCB1.Items.Add("Layer" & Layer)
                    TSCB1.Items.Add(Layer)
                    Start_grain = False
                    End_grain = False
                    Grain_width_size = 0
                    Grain_numbers = 0
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                            If Start_grain = False Then
                                Start_grain = True
                                Grain_numbers = 1
                            End If
                        End If
                        If Start_grain = True Then
                            If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                MyBitmap.SetPixel(i, j, Color.Green)
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                Grain_width_size += 1
                                If i = BottomRight_X Then
                                    Grain_numbers = Grain_numbers - 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                End If
                            End If
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                For x = i - 2 To i + 2
                                    For y = j - 2 To j + 2
                                        MyBitmap.SetPixel(x, y, Color.Red)
                                    Next
                                Next
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                If Grain_width_size > Val(TextBox5.Text) Then
                                    Grain_size(Layer, Grain_numbers) = Grain_width_size
                                    ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain" & Grain_numbers & "): " & Grain_width_size)
                                    Grain_numbers += 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                    Grain_width_size = 0
                                End If
                            End If
                        End If
                    Next
                Next

            End If
        End If
        If TSDDB1.Text = "Black & White" Then
            Dim temp_Grain_height_size As Single
            Dim Grain_height_counter As Single
            Dim color_y As Color, color_grain As Color
            color_y = Color.Green
            color_grain = Color.Blue
            If CheckBox3.Checked = True Then
                For j = TopLeft_Y To BottomRight_Y
                    For i = TopLeft_X To BottomRight_X
                        Grain_height(Grain(i, j)) = 0
                    Next
                Next
                For p = 0 To Grain_numbers
                    Counter_i(p) = 0
                Next
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                Dim MyBitmap As Bitmap
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 10 And Color_G < 10 And Color_B < 10 Then
                            MyBitmap.SetPixel(i, j, Color.Black)
                        Else
                            MyBitmap.SetPixel(i, j, Color.White)
                        End If
                    Next
                Next

                PictureBox2.Image = MyBitmap
                PictureBox2.Refresh()
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    Layer += 1
                    TSCB1.Items.Add(Layer)
                    Start_grain = True
                    End_grain = False
                    Grain_width_size = 0
                    Grain_number = 0
                    Grain_height_size = 0
                    Grain_height_counter = 0
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Start_grain = True Then
                            If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                MyBitmap.SetPixel(i, j, color_y)
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                Grain_width_size += 1
                                MyBitmap.SetPixel(i, j, color_grain)
                                If i Mod 5 = 0 Then
                                    Grain_height_size = 0
                                    Grain_height_counter += 1
                                    For y = j To TopLeft_Y Step -1 'upward scan
                                        Color_R_y = MyBitmap.GetPixel(i, y).R
                                        Color_G_y = MyBitmap.GetPixel(i, y).G
                                        Color_B_y = MyBitmap.GetPixel(i, y).B
                                        If Color_R_y >= 5 And Color_G_y >= 5 And Color_B_y >= 5 Then
                                            Grain_height_size += 1
                                            MyBitmap.SetPixel(i, y, color_y)
                                        ElseIf Color_R_y < 5 And Color_G_y < 5 And Color_B_y < 5 Then
                                            GoTo e1
                                        End If
                                    Next
e1:
                                    For y = j To BottomRight_Y 'downward scan
                                        Color_R_y = MyBitmap.GetPixel(i, y).R
                                        Color_G_y = MyBitmap.GetPixel(i, y).G
                                        Color_B_y = MyBitmap.GetPixel(i, y).B
                                        If Color_R_y >= 5 And Color_G_y >= 5 And Color_B_y >= 5 Then
                                            Grain_height_size += 1
                                            MyBitmap.SetPixel(i, y, color_y)
                                        ElseIf Color_R_y < 5 And Color_G_y < 5 And Color_B_y < 5 Then
                                            GoTo e2
                                        End If
                                    Next
e2:
                                    temp_Grain_height_size = temp_Grain_height_size + Grain_height_size
                                    Grain_height(Grain(i, j)) = Grain_height_size
                                    Grain_number = Grain(i, j)
                                End If
                            End If
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                For x = i - 2 To i + 2
                                    For y = j - 2 To j + 2
                                        MyBitmap.SetPixel(x, y, Color.Red)
                                    Next
                                Next
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                If Grain_width_size > Val(TextBox5.Text) Then
                                    Counter_i(Grain_number) += 1
                                    Grain_width(Grain_number) = (Grain_width(Grain_number) + Grain_width_size) / Counter_i(Grain_number)
                                    Grain_height(Grain_number) = (Grain_height(Grain_number) + temp_Grain_height_size / Grain_height_counter) / Counter_i(Grain_number)
                                    ListBox1.Items.Add("Grain width " & Grain_number & "): " & Grain_width(Grain_number) & " - Grain height " & Grain_number & "): " & Grain_height(Grain_number))
                                    Grain_width_size = 0
                                    Grain_height_size = 0
                                    temp_Grain_height_size = 0
                                    Grain_height_counter = 0
                                    color_y = Color.FromArgb(250 * Rnd(), 250 * Rnd(), 250 * Rnd())
                                    color_grain = Color.FromArgb(250 * Rnd(), 250 * Rnd(), 250 * Rnd())
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        End If
        If TSDDB1.Text = "Grayscale" Then
            If CheckBox3.Checked = True Then
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                Dim MyBitmap As Bitmap
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    For i = TopLeft_X To BottomRight_X
                        For x = i - 1 To i + 1
                            For y = j - 1 To j + 1
                                Color_R_i = MyBitmap.GetPixel(i, j).R
                                Color_G_i = MyBitmap.GetPixel(i, j).G
                                Color_B_i = MyBitmap.GetPixel(i, j).B
                                Color_R_x = MyBitmap.GetPixel(x, y).R
                                Color_G_x = MyBitmap.GetPixel(x, y).G
                                Color_B_x = MyBitmap.GetPixel(x, y).B
                                If MyBitmap.GetPixel(x, y) <> MyBitmap.GetPixel(i, j) Then
                                    If Math.Abs(Color_R_x - Color_R_i) < 10 Then GoTo h6
                                    If Math.Abs(Color_G_x - Color_G_i) < 10 Then GoTo h6
                                    If Math.Abs(Color_B_x - Color_B_i) < 10 Then GoTo h6
                                    Grain_boundary_color(x, y) = True
                                    GoTo h7
                                End If
h6:
                            Next
                        Next
h7:
                    Next
                Next
                PictureBox2.Image = MyBitmap
                PictureBox2.Refresh()
                TPB1.Value = 0
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    For i = TopLeft_X To BottomRight_X
                        If Grain_boundary_color(i, j) = True Then
                            MyBitmap.SetPixel(i, j, Color.FromArgb(0, 0, 0))
                        Else
                            MyBitmap.SetPixel(i, j, Color.FromArgb(255, 255, 255))
                        End If
                        PictureBox2.Image = MyBitmap
                        PictureBox2.Refresh()
                    Next
                Next
                If CheckBox4.Checked = True Then
                    TSCB1.Items.Clear()
                    ListBox1.Items.Clear()
                    MyBitmap = PictureBox2.Image
                    Layer = 0
                    TPB1.Value = 0
                    TPB1.Maximum = BottomRight_Y + 1
                    For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                        TPB1.Value = j
                        Layer += 1
                        TSCB1.Items.Add(Layer)
                        Start_grain = False
                        End_grain = False
                        Grain_width_size = 0
                        Grain_numbers = 0
                        For i = TopLeft_X To BottomRight_X
                            Color_R = MyBitmap.GetPixel(i, j).R
                            Color_G = MyBitmap.GetPixel(i, j).G
                            Color_B = MyBitmap.GetPixel(i, j).B
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                If Start_grain = False Then
                                    Start_grain = True
                                    Grain_numbers = 1
                                End If
                            End If
                            If Start_grain = True Then
                                If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                    MyBitmap.SetPixel(i, j, Color.Green)
                                    PictureBox2.Image = MyBitmap
                                    PictureBox2.Refresh()
                                    Grain_width_size += 1
                                    If i = BottomRight_X Then
                                        Grain_numbers = Grain_numbers - 1
                                        Grain_numbers_layer(Layer) = Grain_numbers
                                    End If
                                End If
                                If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                    For x = i - 2 To i + 2
                                        For y = j - 2 To j + 2
                                            MyBitmap.SetPixel(x, y, Color.Red)
                                        Next
                                    Next
                                    PictureBox2.Image = MyBitmap
                                    PictureBox2.Refresh()
                                    If Grain_width_size > Val(TextBox5.Text) Then
                                        Grain_size(Layer, Grain_numbers) = Grain_width_size
                                        ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain" & Grain_numbers & "): " & Grain_width_size)
                                        Grain_numbers += 1
                                        Grain_numbers_layer(Layer) = Grain_numbers
                                        Grain_width_size = 0
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End If
        End If
        If TSDDB1.Text = "Color" Then
            If CheckBox3.Checked = True Then
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                Dim MyBitmap As Bitmap
                Dim Color_gray As Integer
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                For j = TopLeft_Y To BottomRight_Y 'Step Val(TextBox4.Text)
                    TPB1.Value = j
                    For i = TopLeft_X To BottomRight_X
                        Color_R_i = MyBitmap.GetPixel(i, j).R
                        Color_G_i = MyBitmap.GetPixel(i, j).G
                        Color_B_i = MyBitmap.GetPixel(i, j).B
                        Color_gray = (Color_R_i + Color_G_i + Color_B_i) \ 3
                        MyBitmap.SetPixel(i, j, Color.FromArgb(Color_gray, Color_gray, Color_gray))
                    Next
                Next
                PictureBox2.Image = MyBitmap
                PictureBox2.Refresh()
                For j = TopLeft_Y To BottomRight_Y
                    TPB1.Value = j
                    For i = TopLeft_X To BottomRight_X
                        For x = i - 1 To i + 1
                            For y = j - 1 To j + 1
                                Color_R_i = MyBitmap.GetPixel(i, j).R
                                Color_G_i = MyBitmap.GetPixel(i, j).G
                                Color_B_i = MyBitmap.GetPixel(i, j).B
                                Color_R_x = MyBitmap.GetPixel(x, y).R
                                Color_G_x = MyBitmap.GetPixel(x, y).G
                                Color_B_x = MyBitmap.GetPixel(x, y).B
                                If MyBitmap.GetPixel(x, y) <> MyBitmap.GetPixel(i, j) Then
                                    If Math.Abs(Color_R_x - Color_R_i) < 10 Then GoTo h8
                                    If Math.Abs(Color_G_x - Color_G_i) < 10 Then GoTo h8
                                    If Math.Abs(Color_B_x - Color_B_i) < 10 Then GoTo h8
                                    Grain_boundary_color(x, y) = True
                                    GoTo h9
                                End If
h8:
                            Next
                        Next
h9:
                    Next
                Next
                PictureBox2.Image = MyBitmap
                PictureBox2.Refresh()
                TPB1.Value = 0
                For j = TopLeft_Y To BottomRight_Y
                    TPB1.Value = j
                    For i = TopLeft_X To BottomRight_X
                        If Grain_boundary_color(i, j) = True Then
                            MyBitmap.SetPixel(i, j, Color.FromArgb(0, 0, 0))
                        Else
                            MyBitmap.SetPixel(i, j, Color.FromArgb(255, 255, 255))
                        End If
                    Next
                Next
                PictureBox2.Image = MyBitmap
                PictureBox2.Refresh()
                If CheckBox4.Checked = True Then
                    TSCB1.Items.Clear()
                    ListBox1.Items.Clear()
                    MyBitmap = PictureBox2.Image
                    Layer = 0
                    TPB1.Value = 0
                    TPB1.Maximum = BottomRight_Y + 1
                    For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                        TPB1.Value = j
                        Layer += 1
                        TSCB1.Items.Add(Layer)
                        Start_grain = False
                        End_grain = False
                        Grain_width_size = 0
                        Grain_numbers = 0
                        For i = TopLeft_X To BottomRight_X
                            Color_R = MyBitmap.GetPixel(i, j).R
                            Color_G = MyBitmap.GetPixel(i, j).G
                            Color_B = MyBitmap.GetPixel(i, j).B
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                If Start_grain = False Then
                                    Start_grain = True
                                    Grain_numbers = 1
                                End If
                            End If
                            If Start_grain = True Then
                                If Color_R >= 5 And Color_G >= 5 And Color_B >= 5 Then
                                    MyBitmap.SetPixel(i, j, Color.Green)
                                    PictureBox2.Image = MyBitmap
                                    PictureBox2.Refresh()
                                    Grain_width_size += 1
                                    If i = BottomRight_X Then
                                        Grain_numbers = Grain_numbers - 1
                                        Grain_numbers_layer(Layer) = Grain_numbers
                                    End If
                                End If
                                If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                    For x = i - 2 To i + 2
                                        For y = j - 2 To j + 2
                                            MyBitmap.SetPixel(x, y, Color.Red)
                                        Next
                                    Next
                                    PictureBox2.Image = MyBitmap
                                    PictureBox2.Refresh()
                                    If Grain_width_size > Val(TextBox5.Text) Then
                                        Grain_size(Layer, Grain_numbers) = Grain_width_size
                                        ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain" & Grain_numbers & "): " & Grain_width_size)
                                        Grain_numbers += 1
                                        Grain_numbers_layer(Layer) = Grain_numbers
                                        Grain_width_size = 0
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End If
        End If
        If TSDDB1.Text = "Black & White" Then
            Dim temp_Grain_height_size As Single
            Dim Grain_height_counter As Single
            Dim color_y As Color, color_grain As Color
            color_y = Color.Green
            color_grain = Color.Blue
            If CheckBox1.Checked = True Then
                Grain_numbers = 0
                TSCB1.Items.Clear()
                ListBox1.Items.Clear()
                Dim MyBitmap As Bitmap
                MyBitmap = PictureBox2.Image
                Layer = 0
                TPB1.Value = 0
                TPB1.Maximum = BottomRight_Y + 1
                Dim bold_grain_boundary(0 To BottomRight_X + 10, 0 To BottomRight_Y + 10) As Boolean
                For j = TopLeft_Y To BottomRight_Y
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                            For x = i - Math.Round(Val(TextBox5.Text) / 2, 0) To i + Math.Round(Val(TextBox5.Text) / 2, 0)
                                For y = j - Math.Round(Val(TextBox5.Text) / 2, 0) To j + Math.Round(Val(TextBox5.Text) / 2, 0)
                                    bold_grain_boundary(x, y) = True
                                Next
                            Next
                        End If
                    Next
                Next
                For j = TopLeft_Y To BottomRight_Y
                    For i = TopLeft_X To BottomRight_X
                        If bold_grain_boundary(i, j) = True Then
                            MyBitmap.SetPixel(i, j, Color.Black)
                        End If
                    Next
                    PictureBox2.Image = MyBitmap
                    PictureBox2.Refresh()
                Next
                For j = TopLeft_Y To BottomRight_Y Step Val(TextBox4.Text)
                    TPB1.Value = j
                    Layer += 1
                    TSCB1.Items.Add(Layer)
                    Start_grain = False
                    End_grain = False
                    Grain_width_size = 0
                    Grain_height_size = 0
                    Grain_height_counter = 0
                    Grain_area = 0
                    For i = TopLeft_X To BottomRight_X
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                            If Start_grain = False Then
                                Start_grain = True
                                Grain_numbers = 1
                            End If
                        End If
                        If Start_grain = True Then
                            If Color_R >= 240 And Color_G >= 240 And Color_B >= 240 Then
                                MyBitmap.SetPixel(i, j, color_y)
                                PictureBox2.Image = MyBitmap
                                PictureBox2.Refresh()
                                MyBitmap.SetPixel(i, j, color_grain)
                                Dim c_1 As Integer
                                Dim c_2 As Integer
                                Dim c_3 As Integer
                                If i Mod 10 = 0 Then
                                    ' For ii = 0 To 2
                                    For si = TopLeft_X To BottomRight_X
                                        For sj = TopLeft_Y To BottomRight_Y
                                            c_1 = MyBitmap.GetPixel(si, sj).R
                                            c_2 = MyBitmap.GetPixel(si, sj).G
                                            c_3 = MyBitmap.GetPixel(si, sj).B
                                            If color_grain.R = c_1 And color_grain.G = c_2 And color_grain.B = c_3 Then
                                                For x = si - Math.Round(Val(TextBox5.Text) / 2, 0) To si + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                    For y = sj - Math.Round(Val(TextBox5.Text) / 2, 0) To sj + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                        Color_R_y = MyBitmap.GetPixel(x, y).R
                                                        Color_G_y = MyBitmap.GetPixel(x, y).G
                                                        Color_B_y = MyBitmap.GetPixel(x, y).B
                                                        If Color_R_y >= 240 And Color_G_y >= 240 And Color_B_y >= 240 Then
                                                            MyBitmap.SetPixel(x, y, color_grain)
                                                            Grain_area += 1
                                                            Grain(x, y) = Grain_numbers
                                                        End If
                                                    Next
                                                Next
                                            End If
                                        Next
                                    Next
                                    For si = BottomRight_X To TopLeft_X Step -1
                                        For sj = BottomRight_Y To TopLeft_Y Step -1
                                            c_1 = MyBitmap.GetPixel(si, sj).R
                                            c_2 = MyBitmap.GetPixel(si, sj).G
                                            c_3 = MyBitmap.GetPixel(si, sj).B
                                            If color_grain.R = c_1 And color_grain.G = c_2 And color_grain.B = c_3 Then
                                                For x = si - Math.Round(Val(TextBox5.Text) / 2, 0) To si + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                    For y = sj - Math.Round(Val(TextBox5.Text) / 2, 0) To sj + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                        Color_R_y = MyBitmap.GetPixel(x, y).R
                                                        Color_G_y = MyBitmap.GetPixel(x, y).G
                                                        Color_B_y = MyBitmap.GetPixel(x, y).B
                                                        If Color_R_y >= 240 And Color_G_y >= 240 And Color_B_y >= 240 Then
                                                            MyBitmap.SetPixel(x, y, color_grain)
                                                            Grain_area += 1
                                                            Grain(x, y) = Grain_numbers
                                                        End If
                                                    Next
                                                Next
                                            End If
                                        Next
                                    Next
                                    For si = BottomRight_X To TopLeft_X Step -1
                                        For sj = TopLeft_Y To BottomRight_Y
                                            c_1 = MyBitmap.GetPixel(si, sj).R
                                            c_2 = MyBitmap.GetPixel(si, sj).G
                                            c_3 = MyBitmap.GetPixel(si, sj).B
                                            If color_grain.R = c_1 And color_grain.G = c_2 And color_grain.B = c_3 Then
                                                For x = si - Math.Round(Val(TextBox5.Text) / 2, 0) To si + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                    For y = sj - Math.Round(Val(TextBox5.Text) / 2, 0) To sj + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                        Color_R_y = MyBitmap.GetPixel(x, y).R
                                                        Color_G_y = MyBitmap.GetPixel(x, y).G
                                                        Color_B_y = MyBitmap.GetPixel(x, y).B
                                                        If Color_R_y >= 240 And Color_G_y >= 240 And Color_B_y >= 240 Then
                                                            MyBitmap.SetPixel(x, y, color_grain)
                                                            Grain_area += 1
                                                            Grain(x, y) = Grain_numbers
                                                        End If
                                                    Next
                                                Next
                                            End If
                                        Next
                                    Next
                                    For si = TopLeft_X To BottomRight_X
                                        For sj = BottomRight_Y To TopLeft_Y Step -1
                                            c_1 = MyBitmap.GetPixel(si, sj).R
                                            c_2 = MyBitmap.GetPixel(si, sj).G
                                            c_3 = MyBitmap.GetPixel(si, sj).B
                                            If color_grain.R = c_1 And color_grain.G = c_2 And color_grain.B = c_3 Then
                                                For x = si - Math.Round(Val(TextBox5.Text) / 2, 0) To si + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                    For y = sj - Math.Round(Val(TextBox5.Text) / 2, 0) To sj + Math.Round(Val(TextBox5.Text) / 2, 0)
                                                        Color_R_y = MyBitmap.GetPixel(x, y).R
                                                        Color_G_y = MyBitmap.GetPixel(x, y).G
                                                        Color_B_y = MyBitmap.GetPixel(x, y).B
                                                        If Color_R_y >= 240 And Color_G_y >= 240 And Color_B_y >= 240 Then
                                                            MyBitmap.SetPixel(x, y, color_grain)
                                                            Grain_area += 1
                                                            Grain(x, y) = Grain_numbers
                                                        End If
                                                    Next
                                                Next
                                            End If
                                        Next
                                    Next
                                    PictureBox2.Image = MyBitmap
                                    PictureBox2.Refresh()
                                End If
                                If i = BottomRight_X Then
                                    Grain_numbers = Grain_numbers - 1
                                    Grain_numbers_layer(Layer) = Grain_numbers
                                End If
                            End If
                            If Color_R < 5 And Color_G < 5 And Color_B < 5 Then
                                Grain_size(Layer, Grain_numbers) = Grain_width_size
                                ListBox1.Items.Add("Grain size (Layer" & Layer & " - " & "Grain area " & Grain_numbers & "): " & Grain_area)
                                Grain_numbers += 1
                                Grain_numbers_layer(Layer) = Grain_numbers
                                Grain_width_size = 0
                                Grain_height_size = 0
                                temp_Grain_height_size = 0
                                Grain_height_counter = 0
                                Grain_area = 0
                                color_y = Color.FromArgb(250 * Rnd(), 250 * Rnd(), 250 * Rnd())
                                color_grain = Color.FromArgb(250 * Rnd(), 250 * Rnd(), 250 * Rnd())
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        TextBox4.Text = TrackBar2.Value
    End Sub

    Private Sub Graph_Click(sender As Object, e As EventArgs) Handles Graph.Click
        Chart1.Series.Clear()
        Dim Series1 As New DataVisualization.Charting.Series
        With Series1
            .Name = "Grain Size"
            .ChartType = SeriesChartType.Bar
            .XValueMember = "Grian Number"
            .YValueMembers = "Grain Size"
        End With
        Chart1.Series.Add(Series1)
        For zn = 1 To Grain_numbers_layer(CInt(TSCB1.Text))
            Me.Chart1.Series(0).Points.AddXY(zn, Grain_size(CInt(TSCB1.Text), zn))
        Next
        Me.Chart1.ResetAutoValues()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Calibration_start = False
        Calibration_end = True
        pixel_size = Val(TextBox16.Text) / (Math.Abs(Calibration_start_x - Calibration_end_x))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Region_start = True
        Region_end = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Region_start = False
        Region_end = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Calibration_start = True
        Calibration_end = False
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        Form1.Show()
    End Sub

    Private Sub LicenseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LicenseToolStripMenuItem.Click
        Form2.show
    End Sub

    Private Sub ManualToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ManualToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Dim ofd As New OpenFileDialog
        ofd.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyPictures
        ofd.Filter = "JPEG files (*.jpg)|*.jpg|Bitmap files (*.bmp)|*.bmp|TIF files (*.tif)|*.tif"
        Dim result As DialogResult = ofd.ShowDialog
        If Not (PictureBox2) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox2.Image = Image.FromFile(ofd.FileName)
        End If
        If Not (PictureBox1) Is Nothing And ofd.FileName <> String.Empty Then
            PictureBox1.Image = Image.FromFile(ofd.FileName)
        End If
    End Sub

    Private Sub MeasuringgrainSizeToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim Crack_no_temp As Integer
        Dim crack_start, ttt As Boolean
        Dim Crack_length_temp As Integer
        Dim Color_i As Color
        Dim MyBitmap As Bitmap
        Dim Last_i, Last_j As Integer
        'Dim Last_i(), Last_j() As Integer
        MyBitmap = PictureBox2.Image
        TPB1.Value = 0
        TPB1.Maximum = PictureBox2.Width + 1
        crack_start = False
        Crack_no_temp = 0
        Crack_length_temp = 0
        For i = 1 To PictureBox2.Width
            TPB1.Value = i
            For j = 1 To PictureBox2.Height
                If Pore_Pixel(i, j) = True Then
                    MyBitmap.SetPixel(i, j, Color.White)
                    PictureBox2.Image = MyBitmap
                    PictureBox2.Refresh()
                End If
                If Crack_Pixel(i, j) = True Then
                    If crack_start = False Then
                        Crack_no_temp += 1
                        Crack_length_temp += 1
                        crack_start = True
                        Color_i = Color.FromArgb(Rnd() * 255, Rnd() * 255, Rnd() * 255)
                    End If
                    ttt = False
                    Dim fgh As Integer = 100
                    Dim counter_ As Integer

                    Last_i = i
                    Last_j = j
                    'ReDim Last_i(Crack_length_temp)
                    'ReDim Last_j(Crack_length_temp)

                    'Last_i(Crack_length_temp) = i
                    'Last_j(Crack_length_temp) = j
find_2:
                    counter_ = 0
                    For ii = Last_i - fgh To Last_i + fgh
                        For jj = Last_j - fgh To Last_j + fgh
                            If ii > 1 And ii < PictureBox2.Width - 1 And jj > 1 And jj < PictureBox2.Height - 1 Then
                                If Crack_Pixel_checked(ii, jj) = False Then
                                    If Crack_Pixel(ii, jj) = True Then
                                        If Math.Abs(Last_i - ii) <= 7.2 And Math.Abs(Last_j - jj) <= 7.2 Then
                                            Crack_length_temp += 1
                                            counter_ += 1
                                            Last_i = ii
                                            Last_j = jj
                                            'Last_i(Crack_length_temp) = ii
                                            'Last_j(Crack_length_temp) = jj
                                            Crack_Pixel_checked(ii, jj) = True
                                            MyBitmap.SetPixel(ii, jj, Color_i)
                                            PictureBox2.Image = MyBitmap
                                            PictureBox2.Refresh()
                                            ttt = True
                                            GoTo find_1
                                        End If
                                    Else
                                    End If
                                End If
                            End If
                        Next
                    Next
find_1:
                    If counter_ > 0 Then GoTo find_2
                    ' For ii = i - fgh To i + fgh
                    'For jj = j - fgh To j + fgh
                    'If ii > 1 And ii < PictureBox2.Width - 1 And jj > 1 And jj < PictureBox2.Height - 1 Then
                    'Crack_Pixel_checked(ii, jj) = False
                    'End If
                    'Next
                    'Next
                    If ttt = False Then
                        If Crack_length_temp < 11 Then
                            Crack_no_temp -= 1
                            GoTo h_1
                        End If
                        Crack_Pixel_length(Crack_no_temp) = Crack_length_temp
                        Crack_No = Crack_no_temp
                        crack_start = False
                        ListBox3.Items.Add("Crack No.: " & Crack_No & " - " & "Size: " & Crack_length_temp)
h_1:
                        crack_start = False
                        Crack_length_temp = 0
                    End If
                End If
            Next
        Next
        Crack_length_temp = Crack_length_temp
        Dim crack_min, crack_max, crack_average As Single
        crack_min = 20000
        crack_max = 0
        crack_average = 0
        For i = 1 To Crack_No
            If Crack_Pixel_length(i) > crack_max Then
                crack_max = Crack_Pixel_length(i)
            End If
            If Crack_Pixel_length(i) < crack_min Then
                crack_min = Crack_Pixel_length(i)
            End If
            crack_average = crack_average + Crack_Pixel_length(i)
        Next
        crack_average = crack_average / Crack_No

        Dim FILE_NAME As String = Directory.GetCurrentDirectory() & "\Results_crack_size.txt"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, IO.FileMode.Append)
        objWriter.WriteLine("Crack Max Size [Pixel]: " & Math.Round(crack_max, 2))
        objWriter.WriteLine("Crack Min Size [Pixel]: " & Math.Round(crack_min, 2))
        objWriter.WriteLine("Crack Average Size [Pixel]: " & Math.Round(crack_average, 2))

        For n = 1 To Crack_No
            objWriter.WriteLine("Crack " & n & " " & Crack_Pixel_length(n))
        Next

        PictureBox2.Image.Save(Directory.GetCurrentDirectory() & "\Crack_size_No.jpg")
        objWriter.Close()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click

        Dim gray_scale As Integer
        Dim gray_scale_2 As Integer
        Dim gray_scale_1 As Integer
        Dim gray_scale_0 As Integer
        Dim MyBitmap As Bitmap
        MyBitmap = PictureBox2.Image
        TPB1.Value = 0
        TPB1.Maximum = PictureBox2.Height + 1
        TPB2.Maximum = PictureBox2.Height + 1
        TPB3.Maximum = PictureBox2.Height + 1
        TPB4.Maximum = PictureBox2.Height + 1
        TPB5.Maximum = PictureBox2.Height + 1
        gray_scale_0 = 255
        gray_scale = Val(TextBox6.Text) '230 'crack
        gray_scale_1 = Val(TextBox7.Text) '180 'pore
        gray_scale_2 = Val(TextBox15.Text) '253 'precipitate

        If CheckBox15.Checked = False Then
            width_start = 0
            width_end = PictureBox2.Width
            height_start = 0
            height_end = PictureBox2.Height
        End If
        If CheckBox15.Checked = True And Region_end = True Then
            width_start = 0
            width_end = PictureBox2.Width
            height_start = 0
            height_end = height_end
        End If
        For j = height_start To height_end 'PictureBox2.Height
            For i = width_start To width_end 'PictureBox2.Width
                Pore_Pixel(i, j) = False
                Crack_Pixel(i, j) = False
                Black_test_Pixel(i, j) = False
            Next
        Next
        Dim jk As Integer, Contrast As Integer
        jk = Val(TextBox8.Text) '7
        Contrast = Val(TextBox9.Text) '7
        If CheckBox6.Checked = True Then
            If CheckBox15.Checked = True Then 'White Ppecipitates detection
                For j = height_start + 3 To height_end - 3 ' For j = 3 To PictureBox2.Height - 3
                    TPB2.Value = j
                    For i = width_start + 3 To width_end - 3 'For i = 3 To PictureBox2.Width - 3
                        If i >= 1 And i < PictureBox2.Width Then
                            Color_R = MyBitmap.GetPixel(i, j).R
                            Color_G = MyBitmap.GetPixel(i, j).G
                            Color_B = MyBitmap.GetPixel(i, j).B
                            If Color_R < gray_scale_2 Then
                                MyBitmap.SetPixel(i, j, Color.White)
                            Else
                                MyBitmap.SetPixel(i, j, Color.Black)
                            End If
                        End If

                    Next
                Next
                GoTo Hamed_1
            End If

            For j = height_start + 1 To height_end - 1
                TPB2.Value = j
                For i = width_start + 1 To width_end
                    If i - (jk - 1) >= 1 And i + (jk - 1) < PictureBox2.Width Then
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Math.Abs(Color_R - MyBitmap.GetPixel(i - (jk - 1), j).R) > Contrast And Color_R < MyBitmap.GetPixel(i - (jk - 1), j).R Then
                            Black_test_Pixel(i, j) = True
                        End If
                        If Math.Abs(Color_R - MyBitmap.GetPixel(i + (jk - 1), j).R) > Contrast And Color_R < MyBitmap.GetPixel(i + (jk - 1), j).R Then
                            Black_test_Pixel(i, j) = True
                        End If
                    End If
                Next
            Next
            For j = height_start + 1 To height_end
                TPB2.Value = j
                For i = width_start + 1 To width_end - 1
                    If j - (jk - 1) >= 1 And j + (jk - 1) < PictureBox2.Height Then
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If Math.Abs(Color_R - MyBitmap.GetPixel(i, j - (jk - 1)).R) < Contrast Then
                            Black_test_Pixel(i, j) = False
                        End If
                        If Math.Abs(Color_R - MyBitmap.GetPixel(i, j + (jk - 1)).R) < Contrast Then
                            Black_test_Pixel(i, j) = False
                        End If
                        If Color_R < Val(TextBox7.Text) Then
                            Black_test_Pixel(i, j) = True
                        End If
                        If Color_R >= Val(TextBox6.Text) Then
                            Black_test_Pixel(i, j) = False
                        End If
                    End If
                Next
            Next
            For j = height_start + 3 To height_end - 3
                TPB2.Value = j
                For i = width_start + 3 To width_end - 3
                    If Black_test_Pixel(i, j) = True Then

                    Else
                        MyBitmap.SetPixel(i, j, Color.White)
                    End If
                Next
            Next
Hamed_1:
        End If
        If CheckBox7.Checked = False Then GoTo ki1
        If CheckBox11.Checked = True Then
            Dim c_i As Integer
            c_i = 0
            For p = 1 To Val(TextBox12.Text) '3
                For j = height_start + 3 To height_end - 3
                    TPB3.Value = j
                    For i = width_start + 3 To width_end - 3
                        Color_R = MyBitmap.GetPixel(i, j).R
                        Color_G = MyBitmap.GetPixel(i, j).G
                        Color_B = MyBitmap.GetPixel(i, j).B
                        If MyBitmap.GetPixel(i + 1, j).R = 255 And MyBitmap.GetPixel(i - 1, j).R = 255 And MyBitmap.GetPixel(i, j + 1).R = 255 And MyBitmap.GetPixel(i, j - 1).R = 255 Then

                        End If
                        If MyBitmap.GetPixel(i + 1, j).R = 255 And MyBitmap.GetPixel(i, j + 1).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i + 1, j).R = 255 And MyBitmap.GetPixel(i, j - 1).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i - 1, j).R = 255 And MyBitmap.GetPixel(i, j - 1).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i - 1, j).R = 255 And MyBitmap.GetPixel(i, j + 1).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i + 1, j).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i, j + 1).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i - 1, j).R = 255 Then
                            c_i += 1
                        End If
                        If MyBitmap.GetPixel(i, j - 1).R = 255 Then
                            c_i += 1
                        End If

                        If c_i > 5 Then
                            Black_test_Pixel(i, j) = False
                        End If
                        c_i = 0
                    Next
                Next

                For j = 3 To PictureBox2.Height - 3
                    TPB4.Value = j
                    For i = 3 To PictureBox2.Width - 3
                        If Black_test_Pixel(i, j) = True Then

                        Else
                            MyBitmap.SetPixel(i, j, Color.White)
                        End If
                    Next
                Next
            Next
        End If
        If CheckBox12.Checked = True Then
            Dim c_i As Integer
            c_i = 0

            For j = height_start + 3 To height_end - 3
                TPB3.Value = j
                For i = width_start + 3 To width_end - 3
                    Color_R = MyBitmap.GetPixel(i, j).R
                    Color_G = MyBitmap.GetPixel(i, j).G
                    Color_B = MyBitmap.GetPixel(i, j).B
                    For p_2 = j - Val(TextBox12.Text) To j + Val(TextBox12.Text) '3
                        For p_1 = i - Val(TextBox12.Text) To i + Val(TextBox12.Text) '3
                            If j - Val(TextBox12.Text) > 0 And i - Val(TextBox12.Text) > 0 Then
                                If j + Val(TextBox12.Text) < PictureBox2.Height - 3 And i + Val(TextBox12.Text) < PictureBox2.Width - 3 Then
                                    If MyBitmap.GetPixel(p_1, p_2).R > Val(TextBox12.Text) Then
                                        'If MyBitmap.GetPixel(p_1, p_2).R > 253 Then
                                        c_i += 1
                                    End If
                                End If
                            End If
                        Next
                    Next

                    If c_i > Val(TextBox13.Text) * (Val(TextBox12.Text) * 2 + 1) ^ 2 Then
                        Black_test_Pixel(i, j) = False
                    End If
                    c_i = 0
                Next
            Next

            For j = height_start + 3 To height_end - 3
                TPB4.Value = j
                For i = width_start + 3 To width_end - 3
                    If Black_test_Pixel(i, j) = True Then

                    Else
                        MyBitmap.SetPixel(i, j, Color.White)
                    End If
                Next
            Next

        End If
ki1:
        '/Small dots remover
        GoTo l1
        GoTo l1_1
        For p = 1 To 2 'Val(TextBox12.Text)
            For j = height_start + jk To height_end - jk
                TPB5.Value = j
                For i = width_start + jk To width_end - jk
                    Color_R = MyBitmap.GetPixel(i, j).R
                    Color_G = MyBitmap.GetPixel(i, j).G
                    Color_B = MyBitmap.GetPixel(i, j).B
                    If MyBitmap.GetPixel(i + 1, j).R <= 250 And MyBitmap.GetPixel(i + 1, j).G <= 255 And MyBitmap.GetPixel(i + 1, j).B <= 255 Then
                        ' Crack_Pixel(i, j) = False
                    ElseIf MyBitmap.GetPixel(i - 1, j).R >= 250 And MyBitmap.GetPixel(i - 1, j).G = 0 And MyBitmap.GetPixel(i - 1, j).B = 0 Then
                        ' Crack_Pixel(i, j) = False
                    ElseIf MyBitmap.GetPixel(i, j + 1).R >= 250 And MyBitmap.GetPixel(i, j + 1).G = 0 And MyBitmap.GetPixel(i, j + 1).B = 0 Then
                        ' Crack_Pixel(i, j) = False
                    ElseIf MyBitmap.GetPixel(i, j - 1).R >= 250 And MyBitmap.GetPixel(i, j - 1).G = 0 And MyBitmap.GetPixel(i, j - 1).B = 0 Then
                        ' Crack_Pixel(i, j) = False
                    End If
                Next
            Next
        Next
l1_1:
        For p = 1 To 2 'Val(TextBox12.Text)
            For j = height_start + jk To height_end - jk
                TPB5.Value = j
                For i = width_start + jk To width_end - jk
                    Color_R = MyBitmap.GetPixel(i, j).R
                    Color_G = MyBitmap.GetPixel(i, j).G
                    Color_B = MyBitmap.GetPixel(i, j).B

                    If MyBitmap.GetPixel(i, j).R > 0 And MyBitmap.GetPixel(i, j).G > 0 And MyBitmap.GetPixel(i, j).B > 0 Then
                        If MyBitmap.GetPixel(i + 1, j).R >= 250 And MyBitmap.GetPixel(i + 1, j).G >= 250 And MyBitmap.GetPixel(i + 1, j).B >= 250 Then
                            If MyBitmap.GetPixel(i - 1, j).R >= 250 And MyBitmap.GetPixel(i - 1, j).G >= 250 And MyBitmap.GetPixel(i - 1, j).B >= 250 Then
                                If MyBitmap.GetPixel(i, j + 1).R >= 250 And MyBitmap.GetPixel(i, j + 1).G >= 250 And MyBitmap.GetPixel(i, j + 1).B >= 250 Then
                                    If MyBitmap.GetPixel(i, j - 1).R >= 250 And MyBitmap.GetPixel(i, j - 1).G >= 250 And MyBitmap.GetPixel(i, j - 1).B >= 250 Then
                                        ' Crack_Pixel(i, j) = False
                                    End If
                                End If
                            End If
                        End If
                    End If

                    Dim counter_pixel_red As Integer
                    If MyBitmap.GetPixel(i + 1, j).R <= 250 And MyBitmap.GetPixel(i + 1, j).G <= 255 And MyBitmap.GetPixel(i + 1, j).B <= 255 Then
                        For ii = i - Val(TextBox12.Text) / 2 To i + Val(TextBox12.Text) / 2
                            'For ii = i - Val(TextBox12.Text) / 2 To j + Val(TextBox12.Text) / 2
                            If MyBitmap.GetPixel(i + 1, j).R <= 250 And MyBitmap.GetPixel(i + 1, j).G <= 255 And MyBitmap.GetPixel(i + 1, j).B <= 255 Then
                                counter_pixel_red += 1
                            End If
                        Next
                    End If
                    If counter_pixel_red > 0 And counter_pixel_red <= Val(TextBox12.Text) Then
                        For ii = i - Val(TextBox12.Text) / 2 To i + Val(TextBox12.Text) / 2

                            Crack_Pixel(ii, j) = False
                            Pore_Pixel(ii, j) = False
                            MyBitmap.SetPixel(ii, j, Color.White)
                            'Next
                        Next
                    End If

                Next
            Next
        Next
l1:
        '/Small dots remover



        PictureBox2.Image = MyBitmap
        PictureBox2.Refresh()
        PictureBox2.Image.Save(Directory.GetCurrentDirectory() & "\Removed_background.jpg")
        If CheckBox9.Checked = True Then GoTo l_pore
        Dim counter_x As Integer, counter_y As Integer
        For j = height_start + 2 To height_end - 2 '/Pore
            TPB5.Value = j
            For i = width_start + 2 To width_end - 2
                Color_R = MyBitmap.GetPixel(i, j).R
                Color_G = MyBitmap.GetPixel(i, j).G
                Color_B = MyBitmap.GetPixel(i, j).B
                counter_x = 0
                counter_y = 0
                If Color_R < gray_scale_1 And Color_G < gray_scale_1 And Color_B < gray_scale_1 Then
                    For x = i - 30 To i + 30
                        If x > 0 And x < PictureBox2.Width Then
                            If MyBitmap.GetPixel(x, j).R < gray_scale_1 And MyBitmap.GetPixel(x, j).G < gray_scale_1 And MyBitmap.GetPixel(x, j).B < gray_scale_1 Then
                                'counter_x += 1
                                If MyBitmap.GetPixel(i + 1, j).R < gray_scale_1 Or MyBitmap.GetPixel(i - 1, j).R < gray_scale_1 Then
                                    counter_x += 1
                                End If
                            End If
                        End If
                    Next
                    For y = j - 30 To j + 30
                        If y > 0 And y < PictureBox2.Height Then
                            If MyBitmap.GetPixel(i, y).R < gray_scale_1 And MyBitmap.GetPixel(i, y).G < gray_scale_1 And MyBitmap.GetPixel(i, y).B < gray_scale_1 Then
                                'counter_y += 1
                                If MyBitmap.GetPixel(i, j + 1).R < gray_scale_1 Or MyBitmap.GetPixel(i, j - 1).R < gray_scale_1 Then
                                    counter_y += 1
                                End If
                            End If
                        End If
                    Next

                    If CheckBox10.Checked = False Then
                        If counter_x >= Val(TextBox10.Text) And counter_y >= Val(TextBox10.Text) Then
                            Pore_Pixel(i, j) = True
                        End If
                    Else
                        If counter_x >= Val(TextBox10.Text) Or counter_y >= Val(TextBox10.Text) Then
                            Pore_Pixel(i, j) = True
                        End If
                    End If
                End If
            Next
        Next
l_pore:
        If CheckBox8.Checked = True Then GoTo l_crack
        For j = height_start + 1 To height_end - 1
            TPB5.Value = j
            For i = width_start + 1 To width_end - 1
                Color_R = MyBitmap.GetPixel(i, j).R
                Color_G = MyBitmap.GetPixel(i, j).G
                Color_B = MyBitmap.GetPixel(i, j).B
                counter_x = 0
                counter_y = 0
                Dim Min_color As Integer
                Dim Min_x As Integer
                Dim Min_y As Integer
                Min_color = 258
                If Color_R < gray_scale And Color_G < gray_scale And Color_B < gray_scale Then
                    If Crack_Pixel(i, j) = False Then
                        If Pore_Pixel(i, j) = False Then
                            For x = i - 5 To i + 5
                                If x > 0 And x < PictureBox2.Width Then
                                    If Pore_Pixel(x, j) = True Then GoTo l0
                                End If
                            Next
                            For y = j - 5 To j + 5
                                If y > 0 And y < PictureBox2.Height Then
                                    If Pore_Pixel(i, y) = True Then GoTo l0
                                End If
                            Next
                            For x = i - 5 To i + 5 '/Crack
                                If x > 0 And x < PictureBox2.Width Then
                                    If MyBitmap.GetPixel(x, j).R < gray_scale Or MyBitmap.GetPixel(x, j).G < gray_scale Or MyBitmap.GetPixel(x, j).B < gray_scale Then
                                        counter_x += 1
                                    End If
                                    If MyBitmap.GetPixel(x, j).R < Min_color Then
                                        Min_color = MyBitmap.GetPixel(x, j).R
                                        Min_x = x
                                        Min_y = j
                                    End If
                                    If MyBitmap.GetPixel(x, j).G < Min_color Then
                                        Min_color = MyBitmap.GetPixel(x, j).R
                                        Min_x = x
                                        Min_y = j
                                    End If
                                    If MyBitmap.GetPixel(x, j).B < Min_color Then
                                        Min_color = MyBitmap.GetPixel(x, j).R
                                        Min_x = x
                                        Min_y = j
                                    End If
                                End If
                            Next
                            For y = j - 30 To j + 30 '/Crack
                                If y > 0 And y < PictureBox2.Height Then
                                    If MyBitmap.GetPixel(i, y).R < gray_scale Or MyBitmap.GetPixel(i, y).G < gray_scale Or MyBitmap.GetPixel(i, y).B < gray_scale Then
                                        counter_y += 1
                                    End If
                                End If
                            Next
                            If CheckBox13.Checked = True Then
                                If counter_x > Val(TextBox11.Text) Or counter_y > Val(TextBox11.Text) Then
                                    Crack_Pixel(i, j) = True
                                    Pore_Pixel(i, j) = False
                                End If
                            End If
                            If CheckBox14.Checked = True Then
                                Crack_Pixel(Min_x, Min_y) = True
                                Pore_Pixel(Min_x, Min_y) = False
                            End If
                        End If

                    End If
l0:
                End If
            Next
        Next
l_crack:
        For j = height_start + jk To height_end - jk
            For i = width_start + jk To width_end - jk
                If Pore_Pixel(i, j) = True Then MyBitmap.SetPixel(i, j, Color.Green)
                If Crack_Pixel(i, j) = True Then MyBitmap.SetPixel(i, j, Color.Red)
                If Pore_Pixel(i, j) = False And Crack_Pixel(i, j) = False Then MyBitmap.SetPixel(i, j, Color.White)
            Next
        Next

        PictureBox2.Image = MyBitmap
        PictureBox2.Refresh()
        PictureBox2.Image.Save(Directory.GetCurrentDirectory() & "\Pore_crack.jpg")
        Dim counter_p As Integer, PoreNo As Integer
        counter_p = 0
        PoreNo = 0
        jk = jk * 1
        For j = height_start + jk To height_end - jk
            For i = width_start + jk To width_end - jk
                counter_p = 0
                If Pore_Pixel(i, j) = True Then
                    For x = i - (jk - 1) To i + (jk - 1)
                        For y = j - (jk - 1) To j + (jk - 1)
                            If Pore_Pixel(x, y) = True Then
                                If Pore_No(x, y) <> 0 Then
                                    counter_p = 1
                                    Pore_No(i, j) = Pore_No(x, y)
                                    Pore_Area(PoreNo) = Pore_Area(PoreNo) + 1
                                    Pore_Color(i, j) = Pore_Color(x, y)
                                    MyBitmap.SetPixel(i, j, Pore_Color(i, j))
                                    GoTo U1
                                End If
                            End If
                        Next
                    Next
U1:
                    If counter_p = 0 Then
                        PoreNo += 1
                        Pore_No(i, j) = PoreNo
                        Pore_Area(PoreNo) = Pore_Area(PoreNo) + 1
                        Pore_Color(i, j) = Color.FromArgb(255 * Rnd(), 255 * Rnd(), 255 * Rnd())
                        MyBitmap.SetPixel(i, j, Pore_Color(i, j))
                    End If
                    For x = i - 1 To i + 1
                        For y = j - 1 To j + 1
                            If Pore_Pixel(x, y) = False Then
                                Pore_Boundary(i, j) = True
                                MyBitmap.SetPixel(i, j, Color.Black)
                                Pore_Boundary_no(i, j) = Pore_No(i, j)
                                Pore_Perimeter(Pore_No(i, j)) = Pore_Perimeter(Pore_No(i, j)) + 1
                                GoTo U2
                            End If
                        Next
                    Next
U2:
                End If
            Next
        Next
        Dim FILE_NAME As String = Directory.GetCurrentDirectory() & "\Results.txt"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, IO.FileMode.Append)
        Dim Counter_pore As Integer
        Dim Counter_crack As Integer
        Dim crack_percent As Single
        Dim pore_percant As Single
        Counter_pore = 0
        Counter_crack = 0
        For j = height_start + jk To height_end - jk
            For i = width_start + jk To width_end - jk
                If Pore_Pixel(i, j) = True Then
                    Counter_pore += 1
                End If
                If Crack_Pixel(i, j) = True Then
                    Counter_crack += 1
                End If
            Next
        Next

        pixel_size = 1

        crack_percent = Counter_crack * 100 / ((PictureBox2.Width - 2 * jk) * (height_end - 2 * jk))
        pore_percant = Counter_pore * 100 / ((PictureBox2.Width - 2 * jk) * (height_end - 2 * jk))
        ListBox2.Items.Add("Crack Percentage (%): " & Math.Round(crack_percent, 2))
        objWriter.WriteLine("Crack Percentage (%): " & Math.Round(crack_percent, 2))
        ListBox2.Items.Add("Pore Percentage (%): " & Math.Round(pore_percant, 2))
        objWriter.WriteLine("Pore Percentage (%): " & Math.Round(pore_percant, 2))
        ListBox2.Items.Add("Total Porosity Percentage (%): " & Math.Round(pore_percant + crack_percent, 2))
        objWriter.WriteLine("Total Porosity Percentage (%): " & Math.Round(pore_percant + crack_percent, 2))
        ListBox2.Items.Add("Number of Pores: " & PoreNo)
        objWriter.WriteLine("Number of Pores: " & PoreNo)
        If CheckBox15.Checked = False Then
            For n = 1 To PoreNo
                ListBox2.Items.Add("Pore" & n & " --- " & "Equivalent radius(um): " & Math.Round(Math.Sqrt(Pore_Area(n) * (pixel_size ^ 2) / 3.14), 7) & " --- " & " (Area(um2): " & Pore_Area(n) * pixel_size ^ 2 & " --- " & "Perimeter: " & Pore_Perimeter(n) & " --- " & "Sphericity (4*pi*Area/(Perimeter)^2): " & Math.Round(4 * 3.14 * (Pore_Area(n) / Pore_Perimeter(n) ^ 2), 2) & " --- " & "Sphericity (Area/(perimeter*equivalent radius)): " & Math.Round(Pore_Area(n) / (Pore_Perimeter(n) * Math.Sqrt(Pore_Area(n) / 3.14)), 2) & " )")
                objWriter.WriteLine("Pore" & n & " --- " & "Equivalent radius(um): " & Math.Round(Math.Sqrt(Pore_Area(n) * (pixel_size ^ 2) / 3.14), 7) & " --- " & " (Area(um2): " & Pore_Area(n) * pixel_size ^ 2 & " --- " & "Perimeter: " & Pore_Perimeter(n) & " --- " & "Sphericity (4*pi*Area/(Perimeter)^2): " & Math.Round(4 * 3.14 * (Pore_Area(n) / Pore_Perimeter(n) ^ 2), 2) & " --- " & "Sphericity (Area/(perimeter*equivalent radius)): " & Math.Round(Pore_Area(n) / (Pore_Perimeter(n) * Math.Sqrt(Pore_Area(n) / 3.14)), 2) & " )")
            Next
        Else
            For n = 1 To PoreNo
                ListBox2.Items.Add("Pore" & n & " --- " & "Equivalent radius(um): " & Math.Round(Math.Sqrt(Pore_Area(n) * (pixel_size ^ 2) / 3.14), 7) & " --- " & " (Area(um2): " & Pore_Area(n) * pixel_size ^ 2 & " --- " & "Perimeter: " & Pore_Perimeter(n) & " --- " & "Sphericity (4*pi*Area/(Perimeter)^2): " & Math.Round(4 * 3.14 * (Pore_Area(n) / Pore_Perimeter(n) ^ 2), 2) & " --- " & "Sphericity (Area/(perimeter*equivalent radius)): " & Math.Round(Pore_Area(n) / (Pore_Perimeter(n) * Math.Sqrt(Pore_Area(n) / 3.14)), 2) & " )")
                objWriter.WriteLine("Pore" & n & " --- " & "Equivalent radius(um): " & Math.Round(Math.Sqrt(Pore_Area(n) * (pixel_size ^ 2) / 3.14), 7) & " --- " & " (Area(um2): " & Pore_Area(n) * pixel_size ^ 2 & " --- " & "Perimeter: " & Pore_Perimeter(n) & " --- " & "Sphericity (4*pi*Area/(Perimeter)^2): " & Math.Round(4 * 3.14 * (Pore_Area(n) / Pore_Perimeter(n) ^ 2), 2) & " --- " & "Sphericity (Area/(perimeter*equivalent radius)): " & Math.Round(Pore_Area(n) / (Pore_Perimeter(n) * Math.Sqrt(Pore_Area(n) / 3.14)), 2) & " )")
            Next
        End If
        objWriter.Close()
        PictureBox2.Image = MyBitmap
        PictureBox2.Refresh()
        PictureBox2.Image.Save(Directory.GetCurrentDirectory() & "\Processed.jpg")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub PictureBox2_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox2.MouseDown
        TopLeft_X = e.X
        TopLeft_Y = e.Y
        _painter = True

        If Region_start = True And Region_end = False Then
            height_end = e.Y
        End If

        If Calibration_start = True And Calibration_end = False Then '/precipitation
            Calibration_start_x = e.X
        End If

    End Sub

    Private Sub PictureBox2_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox2.MouseUp
        BottomRight_X = e.X
        BottomRight_Y = e.Y
        _painter = False
        painter = True

        If Calibration_start = True And Calibration_end = False Then '/precipitation
            Calibration_end_x = e.X
        End If
    End Sub

    Private Sub PictureBox2_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox2.MouseMove
        If _painter = True Then
            BottomRight_X = e.X
            BottomRight_Y = e.Y
            PictureBox2.Refresh()
        End If
        If _painter = True Then
            Dim rect As New Rectangle(TopLeft_X, TopLeft_Y, BottomRight_X - TopLeft_X, BottomRight_Y - TopLeft_Y)
            Using pen As New Pen(Color.Blue, 2)
                Dim graphics1 As Graphics = PictureBox2.CreateGraphics()
                graphics1.DrawRectangle(pen, rect)
            End Using
        End If
        If painter = True Then
            'Dim MyBitmap As Bitmap
            'MyBitmap = CType(PictureBox2.Image, Bitmap)
            'Color_R = MyBitmap.GetPixel(e.X, e.Y).R
            'Color_G = MyBitmap.GetPixel(e.X, e.Y).G
            'Color_B = MyBitmap.GetPixel(e.X, e.Y).B
            'TextBox1.Text = Color_R
            'TextBox2.Text = Color_G
            'TextBox3.Text = Color_B
            'PictureBox3.BackColor = MyBitmap.GetPixel(e.X, e.Y)
        End If
    End Sub

End Class
