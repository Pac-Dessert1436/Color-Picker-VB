Option Strict On
Option Infer On

Public Class Form1
    Inherits Form

    Private WithEvents RedSlider As TrackBar
    Private WithEvents GreenSlider As TrackBar
    Private WithEvents BlueSlider As TrackBar
    Private WithEvents ColorDisplay As Panel
    Private WithEvents HexCode As TextBox
    Private WithEvents CopyButton As Button

    Private redLabel As Label
    Private greenLabel As Label
    Private blueLabel As Label
    Private redValue As Label
    Private greenValue As Label
    Private blueValue As Label
    Private copyMessage As Label
    Private titleLabel As Label

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Text = "RGB Color Picker"
        Size = New Size(600, 450)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        BackColor = Color.FromArgb(245, 245, 245)

        titleLabel = New Label With {
            .Location = New Point(20, 20),
            .Size = New Size(550, 50),
            .Text = "RGB Color Picker",
            .Font = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = Color.FromArgb(44, 62, 80)
        }
        Controls.Add(titleLabel)

        redLabel = New Label With {
            .Location = New Point(40, 80),
            .Size = New Size(60, 20),
            .Text = "Red",
            .ForeColor = Color.Red,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        Controls.Add(redLabel)

        RedSlider = New TrackBar With {
            .Minimum = 0,
            .Maximum = 255,
            .TickFrequency = 25,
            .Value = 128,
            .Location = New Point(110, 80),
            .Size = New Size(250, 30)
        }
        Controls.Add(RedSlider)

        redValue = New Label With {
            .Location = New Point(370, 80),
            .Size = New Size(50, 20),
            .ForeColor = Color.Red,
            .Text = "128",
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Consolas", 9)
        }
        Controls.Add(redValue)

        greenLabel = New Label With {
            .Location = New Point(40, 140),
            .Size = New Size(60, 20),
            .Text = "Green",
            .ForeColor = Color.Green,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        Controls.Add(greenLabel)

        GreenSlider = New TrackBar With {
            .Minimum = 0,
            .Maximum = 255,
            .TickFrequency = 25,
            .Value = 128,
            .Location = New Point(110, 140),
            .Size = New Size(250, 30)
        }
        Controls.Add(GreenSlider)

        greenValue = New Label With {
            .Location = New Point(370, 140),
            .Size = New Size(50, 20),
            .ForeColor = Color.Green,
            .Text = "128",
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Consolas", 9)
        }
        Controls.Add(greenValue)

        blueLabel = New Label With {
            .Location = New Point(40, 200),
            .Size = New Size(60, 20),
            .Text = "Blue",
            .ForeColor = Color.Blue,
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        Controls.Add(blueLabel)

        BlueSlider = New TrackBar With {
            .Minimum = 0,
            .Maximum = 255,
            .TickFrequency = 25,
            .Value = 128,
            .Location = New Point(110, 200),
            .Size = New Size(250, 30)
        }
        Controls.Add(BlueSlider)

        blueValue = New Label With {
            .Location = New Point(370, 200),
            .Size = New Size(50, 20),
            .ForeColor = Color.Blue,
            .Text = "128",
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Consolas", 9)
        }
        Controls.Add(blueValue)

        Dim ColorDisplayGroup As New GroupBox With {
            .Location = New Point(440, 80),
            .Size = New Size(120, 165),
            .Text = "Color Preview",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        Controls.Add(ColorDisplayGroup)

        ColorDisplay = New Panel With {
            .Location = New Point(10, 25),
            .Size = New Size(100, 100),
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.FromArgb(128, 128, 128)
        }
        ColorDisplayGroup.Controls.Add(ColorDisplay)

        Dim hexGroup As New GroupBox With {
            .Location = New Point(40, 270),
            .Size = New Size(520, 80),
            .Text = "Color Code",
            .Font = New Font("Segoe UI", 9, FontStyle.Regular)
        }
        Controls.Add(hexGroup)

        HexCode = New TextBox With {
            .Location = New Point(20, 25),
            .Size = New Size(200, 25),
            .ReadOnly = True,
            .Text = "#808080",
            .Font = New Font("Consolas", 10)
        }
        hexGroup.Controls.Add(HexCode)

        CopyButton = New Button With {
            .Location = New Point(230, 25),
            .Size = New Size(80, 35),
            .Text = "Copy Code",
            .Font = New Font("Segoe UI", 9),
            .BackColor = Color.FromArgb(52, 152, 219),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        hexGroup.Controls.Add(CopyButton)

        copyMessage = New Label With {
            .Location = New Point(320, 28),
            .Size = New Size(180, 30),
            .Text = "Copied to clipboard!",
            .ForeColor = Color.Green,
            .Visible = False,
            .Font = New Font("Segoe UI", 9, FontStyle.Italic)
        }
        hexGroup.Controls.Add(copyMessage)
        UpdateColor(sender, e)
    End Sub

    Private Sub UpdateColor(sender As Object, e As EventArgs) Handles _
            RedSlider.Scroll, GreenSlider.Scroll, BlueSlider.Scroll

        Dim r As Integer = RedSlider.Value
        Dim g As Integer = GreenSlider.Value
        Dim b As Integer = BlueSlider.Value
        redValue.Text = r.ToString()
        greenValue.Text = g.ToString()
        blueValue.Text = b.ToString()

        redValue.ForeColor = Color.FromArgb(r, 0, 0)
        greenValue.ForeColor = Color.FromArgb(0, g, 0)
        blueValue.ForeColor = Color.FromArgb(0, 0, b)

        ColorDisplay.BackColor = Color.FromArgb(r, g, b)
        HexCode.Text = RgbToHex(r, g, b)
    End Sub

    Private Function RgbToHex(r As Integer, g As Integer, b As Integer) As String
        r = Math.Max(0, Math.Min(255, r))
        g = Math.Max(0, Math.Min(255, g))
        b = Math.Max(0, Math.Min(255, b))

        Return $"#{r:X2}{g:X2}{b:X2}"
    End Function

    Private Sub CopyToClipboard() Handles CopyButton.Click
        Try
            Clipboard.SetText(HexCode.Text)
            copyMessage.Visible = True
            Threading.Thread.Sleep(2000)
            copyMessage.Visible = False
        Catch ex As Exception
            MessageBox.Show("Failed to copy to clipboard: " & ex.Message)
        End Try
    End Sub

    Private Sub HexCode_KeyDown(sender As Object, e As KeyEventArgs) Handles HexCode.KeyDown
        If e.KeyCode = Keys.Enter Then
            CopyToClipboard()
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    <STAThread>
    Friend Shared Sub Main()
        Application.SetHighDpiMode(HighDpiMode.SystemAware)
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form1)
    End Sub
End Class