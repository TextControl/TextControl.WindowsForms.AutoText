Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents TextControl1 As TXTextControl.TextControl
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TextControl1 = New TXTextControl.TextControl
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'TextControl1
        '
        Me.TextControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextControl1.Font = New System.Drawing.Font("Arial", 10.0!)
        Me.TextControl1.Location = New System.Drawing.Point(0, 0)
        Me.TextControl1.Name = "TextControl1"
        Me.TextControl1.Size = New System.Drawing.Size(520, 382)
        Me.TextControl1.TabIndex = 0
        Me.TextControl1.Text = "TextControl1"
        '
        'ComboBox1
        '
        Me.ComboBox1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))
        Me.ComboBox1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple
        Me.ComboBox1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Location = New System.Drawing.Point(48, 232)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(184, 96)
        Me.ComboBox1.Sorted = True
        Me.ComboBox1.TabIndex = 3
        Me.ComboBox1.Visible = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 382)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TextControl1)
        Me.Name = "Form1"
        Me.Text = "AutoText Sample"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim wordList As New ArrayList

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Read the word list from file into an array
        Dim oFile As System.IO.File
        Dim reader As System.IO.StreamReader = oFile.OpenText("wordlist.txt")

        Dim line As String = reader.ReadLine()

        While Not line Is Nothing
            wordList.Add(line)
            line = reader.ReadLine()
        End While

        reader.Close()
        TextControl1.Load("default.rtf", TXTextControl.StreamType.RichTextFormat)
        TextControl1.Selection.Start = TextControl1.Text.Length - 4
    End Sub

    Private Sub TextControl1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextControl1.KeyUp
        ' Fill the Combo with the word suggestions
        If e.Control = True And e.KeyCode = Keys.Space Then
            Dim newLocation As System.Drawing.Point
            newLocation.Y = TextControl1.InputPosition.Location.Y / 15 + TextControl1.Font.Size + 10 - (TextControl1.ScrollLocation.Y / 15)
            newLocation.X = TextControl1.InputPosition.Location.X / 15

            ComboBox1.Location = newLocation

            TextControl1.Selection.Start -= 1
            TextControl1.Selection.Length = 1
            TextControl1.Selection.Text = ""

            Dim startsWith As String = getLastWord()

            For Each word As String In wordList
                If word.StartsWith(startsWith) = True Then
                    If Not ComboBox1.Items.Contains(word) Then
                        ComboBox1.Items.Add(word)
                    End If
                End If
            Next

            If ComboBox1.Items.Count <> 0 Then
                ComboBox1.Visible = True
                ComboBox1.Text = startsWith
                ComboBox1.Focus()
            Else
                escapeBox()
            End If
        End If
    End Sub

    Private Function getLastWord() As String
        ' Get the last word in text
        Dim currentPosition As Integer = TextControl1.InputPosition.TextPosition
        Dim testString As String

        Do While testString <> " " And testString <> vbCrLf
            TextControl1.Selection.Start -= 1
            TextControl1.Selection.Length = 1
            testString = TextControl1.Selection.Text
            ' Exit loop, if word is at the beginning of text
            If TextControl1.Selection.Start = 0 Then
                TextControl1.Selection.Start = 0
                TextControl1.Selection.Length = currentPosition
                Return TextControl1.Selection.Text
            End If
        Loop

        TextControl1.Selection.Start += 1
        TextControl1.Selection.Length = currentPosition - TextControl1.Selection.Start

        Return TextControl1.Selection.Text
    End Function

    Private Sub applyWord()
        ' Apply the selected word to the text
        If ComboBox1.SelectedItem <> "" Then
            TextControl1.Selection.Text = ComboBox1.SelectedItem + " "
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            ComboBox1.Visible = False
        End If
    End Sub

    Private Sub escapeBox()
        TextControl1.Selection.Start += TextControl1.Selection.Length
        TextControl1.Selection.Length = 0
        TextControl1.Focus()
    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        Select Case e.KeyChar
            Case Chr(27)
                ComboBox1.Items.Clear()
                ComboBox1.Text = ""
                ComboBox1.Visible = False
                escapeBox()
            Case Chr(13)
                applyWord()
        End Select
    End Sub

    Private Sub TextControl1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextControl1.GotFocus
        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        ComboBox1.Visible = False
        escapeBox()
    End Sub
End Class
