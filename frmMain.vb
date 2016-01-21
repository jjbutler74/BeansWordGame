Imports System.IO

Public Class frmMain
    Inherits System.Windows.Forms.Form

    ' Stuff to play wav files
    Private Declare Auto Function PlaySound Lib "winmm.dll" (ByVal lpszSoundName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
    Private Const SND_FILENAME As Integer = &H20000

    Dim intMistakes As Integer ' Number of Mistakes
    Dim flgStatus As String ' Main Game Status
    Dim intNoActionLoop As Integer ' Number of Timer loops where the user takes no action

    ' Win Ranges
    Dim intWin1 As Integer
    Dim intWin2 As Integer
    Dim intWin3 As Integer
    Dim intWin4 As Integer

    Dim intNumberWordsPerGame As Integer ' Number of Words per Game
    Dim intQuitGameTime As Integer ' Number of minutes to wait before quiting
    Dim strLetterCase As String ' Letter Case
    Dim strLetterAck As String ' Correct Letter Acknowledgment 
    Dim flgForceFit As String ' Flag to force picture to fit in Picture Boxes
    Dim flgGiveHints As String ' Flag to give letter hints after an amount of time
    Dim flgShowLetters As String ' Flag to show letters at first
    Dim flgSayLetters As String ' Flag to say letters at first

    ' Counters
    Dim intCnt As Integer
    Dim intCnt2 As Integer

    Dim intCurrentWord As Integer ' (1 to 99)
    Dim intCurrentWordLength As Integer ' (1 to 10)
    Dim strCurrentWord As String
    Dim intNumberOfWords As Integer
    Dim intCurrentLetter As Integer

    Dim arrMainList As New ArrayList() ' Holds words

    ' Log File Vars
    Dim strLogFile As String
    Dim objWriter As System.IO.StreamWriter

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
    Friend WithEvents gbMistakes As System.Windows.Forms.GroupBox
    Friend WithEvents lblMistakesHeading As System.Windows.Forms.Label
    Friend WithEvents lblMistakes As System.Windows.Forms.Label
    Friend WithEvents pbPic As System.Windows.Forms.PictureBox
    Friend WithEvents timeMain As System.Windows.Forms.Timer
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents btnBye As System.Windows.Forms.Button
    Friend WithEvents btnNewGame As System.Windows.Forms.Button
    Friend WithEvents pbPrize As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents pbFace As System.Windows.Forms.PictureBox
    Private WithEvents lblHeadingNumberOfWords As System.Windows.Forms.Label
    Private WithEvents lblHeading2 As System.Windows.Forms.Label
    Private WithEvents lblHeadingCurrentWord As System.Windows.Forms.Label
    Private WithEvents lblHeading1 As System.Windows.Forms.Label
    Private WithEvents lblLetter1 As System.Windows.Forms.Label
    Private WithEvents lblLetter2 As System.Windows.Forms.Label
    Private WithEvents lblLetter3 As System.Windows.Forms.Label
    Private WithEvents lblLetter4 As System.Windows.Forms.Label
    Private WithEvents lblLetter5 As System.Windows.Forms.Label
    Private WithEvents lblLetter6 As System.Windows.Forms.Label
    Private WithEvents lblLetter7 As System.Windows.Forms.Label
    Private WithEvents lblLetter8 As System.Windows.Forms.Label
    Private WithEvents lblLetter9 As System.Windows.Forms.Label
    Private WithEvents lblLetter10 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnBye = New System.Windows.Forms.Button
        Me.lblHeadingNumberOfWords = New System.Windows.Forms.Label
        Me.lblHeading2 = New System.Windows.Forms.Label
        Me.lblHeadingCurrentWord = New System.Windows.Forms.Label
        Me.lblHeading1 = New System.Windows.Forms.Label
        Me.gbMistakes = New System.Windows.Forms.GroupBox
        Me.lblMistakes = New System.Windows.Forms.Label
        Me.lblMistakesHeading = New System.Windows.Forms.Label
        Me.pbPic = New System.Windows.Forms.PictureBox
        Me.timeMain = New System.Windows.Forms.Timer(Me.components)
        Me.lblTitle = New System.Windows.Forms.Label
        Me.btnNewGame = New System.Windows.Forms.Button
        Me.pbPrize = New System.Windows.Forms.PictureBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.pbFace = New System.Windows.Forms.PictureBox
        Me.lblLetter1 = New System.Windows.Forms.Label
        Me.lblLetter2 = New System.Windows.Forms.Label
        Me.lblLetter3 = New System.Windows.Forms.Label
        Me.lblLetter4 = New System.Windows.Forms.Label
        Me.lblLetter5 = New System.Windows.Forms.Label
        Me.lblLetter6 = New System.Windows.Forms.Label
        Me.lblLetter7 = New System.Windows.Forms.Label
        Me.lblLetter8 = New System.Windows.Forms.Label
        Me.lblLetter9 = New System.Windows.Forms.Label
        Me.lblLetter10 = New System.Windows.Forms.Label
        Me.gbMistakes.SuspendLayout()
        CType(Me.pbPic, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbPrize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.pbFace, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnBye
        '
        Me.btnBye.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnBye.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBye.Location = New System.Drawing.Point(533, 632)
        Me.btnBye.Name = "btnBye"
        Me.btnBye.Size = New System.Drawing.Size(181, 40)
        Me.btnBye.TabIndex = 0
        Me.btnBye.TabStop = False
        Me.btnBye.Text = "Bye Bye"
        Me.btnBye.UseVisualStyleBackColor = False
        '
        'lblHeadingNumberOfWords
        '
        Me.lblHeadingNumberOfWords.BackColor = System.Drawing.Color.Transparent
        Me.lblHeadingNumberOfWords.Font = New System.Drawing.Font("Comic Sans MS", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeadingNumberOfWords.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblHeadingNumberOfWords.Location = New System.Drawing.Point(604, 34)
        Me.lblHeadingNumberOfWords.Margin = New System.Windows.Forms.Padding(0)
        Me.lblHeadingNumberOfWords.Name = "lblHeadingNumberOfWords"
        Me.lblHeadingNumberOfWords.Size = New System.Drawing.Size(73, 48)
        Me.lblHeadingNumberOfWords.TabIndex = 37
        Me.lblHeadingNumberOfWords.Text = "99"
        Me.lblHeadingNumberOfWords.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblHeadingNumberOfWords.Visible = False
        '
        'lblHeading2
        '
        Me.lblHeading2.BackColor = System.Drawing.Color.Transparent
        Me.lblHeading2.Font = New System.Drawing.Font("Comic Sans MS", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading2.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblHeading2.Location = New System.Drawing.Point(536, 34)
        Me.lblHeading2.Margin = New System.Windows.Forms.Padding(0)
        Me.lblHeading2.Name = "lblHeading2"
        Me.lblHeading2.Size = New System.Drawing.Size(68, 48)
        Me.lblHeading2.TabIndex = 36
        Me.lblHeading2.Text = "of"
        Me.lblHeading2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblHeading2.Visible = False
        '
        'lblHeadingCurrentWord
        '
        Me.lblHeadingCurrentWord.BackColor = System.Drawing.Color.Transparent
        Me.lblHeadingCurrentWord.Font = New System.Drawing.Font("Comic Sans MS", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeadingCurrentWord.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblHeadingCurrentWord.Location = New System.Drawing.Point(462, 34)
        Me.lblHeadingCurrentWord.Margin = New System.Windows.Forms.Padding(0)
        Me.lblHeadingCurrentWord.Name = "lblHeadingCurrentWord"
        Me.lblHeadingCurrentWord.Size = New System.Drawing.Size(74, 48)
        Me.lblHeadingCurrentWord.TabIndex = 35
        Me.lblHeadingCurrentWord.Text = "99"
        Me.lblHeadingCurrentWord.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblHeadingCurrentWord.Visible = False
        '
        'lblHeading1
        '
        Me.lblHeading1.BackColor = System.Drawing.Color.Transparent
        Me.lblHeading1.Font = New System.Drawing.Font("Comic Sans MS", 30.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHeading1.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblHeading1.Location = New System.Drawing.Point(329, 34)
        Me.lblHeading1.Margin = New System.Windows.Forms.Padding(0)
        Me.lblHeading1.Name = "lblHeading1"
        Me.lblHeading1.Size = New System.Drawing.Size(133, 48)
        Me.lblHeading1.TabIndex = 34
        Me.lblHeading1.Text = "Word"
        Me.lblHeading1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblHeading1.Visible = False
        '
        'gbMistakes
        '
        Me.gbMistakes.Controls.Add(Me.lblMistakes)
        Me.gbMistakes.Controls.Add(Me.lblMistakesHeading)
        Me.gbMistakes.Location = New System.Drawing.Point(886, 8)
        Me.gbMistakes.Name = "gbMistakes"
        Me.gbMistakes.Size = New System.Drawing.Size(88, 80)
        Me.gbMistakes.TabIndex = 61
        Me.gbMistakes.TabStop = False
        '
        'lblMistakes
        '
        Me.lblMistakes.Font = New System.Drawing.Font("Arial", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMistakes.ForeColor = System.Drawing.Color.LightSlateGray
        Me.lblMistakes.Location = New System.Drawing.Point(8, 36)
        Me.lblMistakes.Name = "lblMistakes"
        Me.lblMistakes.Size = New System.Drawing.Size(75, 40)
        Me.lblMistakes.TabIndex = 1
        Me.lblMistakes.Text = "0"
        Me.lblMistakes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblMistakesHeading
        '
        Me.lblMistakesHeading.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMistakesHeading.ForeColor = System.Drawing.Color.LightSlateGray
        Me.lblMistakesHeading.Location = New System.Drawing.Point(8, 8)
        Me.lblMistakesHeading.Name = "lblMistakesHeading"
        Me.lblMistakesHeading.Size = New System.Drawing.Size(75, 23)
        Me.lblMistakesHeading.TabIndex = 0
        Me.lblMistakesHeading.Text = "Mistakes"
        Me.lblMistakesHeading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbPic
        '
        Me.pbPic.BackColor = System.Drawing.Color.White
        Me.pbPic.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbPic.Location = New System.Drawing.Point(253, 96)
        Me.pbPic.Name = "pbPic"
        Me.pbPic.Size = New System.Drawing.Size(500, 400)
        Me.pbPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbPic.TabIndex = 62
        Me.pbPic.TabStop = False
        Me.pbPic.Visible = False
        '
        'timeMain
        '
        Me.timeMain.Interval = 3000
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 96.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.Magenta
        Me.lblTitle.Location = New System.Drawing.Point(24, 88)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(1, 1)
        Me.lblTitle.TabIndex = 68
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnNewGame
        '
        Me.btnNewGame.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnNewGame.Enabled = False
        Me.btnNewGame.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNewGame.Location = New System.Drawing.Point(293, 632)
        Me.btnNewGame.Name = "btnNewGame"
        Me.btnNewGame.Size = New System.Drawing.Size(181, 40)
        Me.btnNewGame.TabIndex = 69
        Me.btnNewGame.TabStop = False
        Me.btnNewGame.Text = "New Game"
        Me.btnNewGame.UseVisualStyleBackColor = False
        '
        'pbPrize
        '
        Me.pbPrize.Location = New System.Drawing.Point(32, 96)
        Me.pbPrize.Name = "pbPrize"
        Me.pbPrize.Size = New System.Drawing.Size(1, 1)
        Me.pbPrize.TabIndex = 70
        Me.pbPrize.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.pbFace)
        Me.GroupBox1.Location = New System.Drawing.Point(32, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(88, 80)
        Me.GroupBox1.TabIndex = 71
        Me.GroupBox1.TabStop = False
        '
        'pbFace
        '
        Me.pbFace.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbFace.Location = New System.Drawing.Point(8, 10)
        Me.pbFace.Name = "pbFace"
        Me.pbFace.Size = New System.Drawing.Size(72, 64)
        Me.pbFace.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbFace.TabIndex = 0
        Me.pbFace.TabStop = False
        '
        'lblLetter1
        '
        Me.lblLetter1.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter1.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter1.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter1.Location = New System.Drawing.Point(31, 520)
        Me.lblLetter1.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter1.Name = "lblLetter1"
        Me.lblLetter1.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter1.TabIndex = 72
        Me.lblLetter1.Text = "W"
        Me.lblLetter1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter1.Visible = False
        '
        'lblLetter2
        '
        Me.lblLetter2.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter2.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter2.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter2.Location = New System.Drawing.Point(127, 520)
        Me.lblLetter2.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter2.Name = "lblLetter2"
        Me.lblLetter2.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter2.TabIndex = 76
        Me.lblLetter2.Text = "W"
        Me.lblLetter2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter2.Visible = False
        '
        'lblLetter3
        '
        Me.lblLetter3.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter3.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter3.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter3.Location = New System.Drawing.Point(223, 520)
        Me.lblLetter3.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter3.Name = "lblLetter3"
        Me.lblLetter3.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter3.TabIndex = 77
        Me.lblLetter3.Text = "W"
        Me.lblLetter3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter3.Visible = False
        '
        'lblLetter4
        '
        Me.lblLetter4.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter4.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter4.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter4.Location = New System.Drawing.Point(319, 520)
        Me.lblLetter4.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter4.Name = "lblLetter4"
        Me.lblLetter4.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter4.TabIndex = 78
        Me.lblLetter4.Text = "W"
        Me.lblLetter4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter4.Visible = False
        '
        'lblLetter5
        '
        Me.lblLetter5.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter5.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter5.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter5.Location = New System.Drawing.Point(415, 520)
        Me.lblLetter5.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter5.Name = "lblLetter5"
        Me.lblLetter5.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter5.TabIndex = 79
        Me.lblLetter5.Text = "W"
        Me.lblLetter5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter5.Visible = False
        '
        'lblLetter6
        '
        Me.lblLetter6.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter6.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter6.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter6.Location = New System.Drawing.Point(511, 520)
        Me.lblLetter6.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter6.Name = "lblLetter6"
        Me.lblLetter6.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter6.TabIndex = 80
        Me.lblLetter6.Text = "W"
        Me.lblLetter6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter6.Visible = False
        '
        'lblLetter7
        '
        Me.lblLetter7.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter7.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter7.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter7.Location = New System.Drawing.Point(607, 520)
        Me.lblLetter7.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter7.Name = "lblLetter7"
        Me.lblLetter7.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter7.TabIndex = 81
        Me.lblLetter7.Text = "W"
        Me.lblLetter7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter7.Visible = False
        '
        'lblLetter8
        '
        Me.lblLetter8.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter8.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter8.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter8.Location = New System.Drawing.Point(703, 520)
        Me.lblLetter8.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter8.Name = "lblLetter8"
        Me.lblLetter8.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter8.TabIndex = 82
        Me.lblLetter8.Text = "W"
        Me.lblLetter8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter8.Visible = False
        '
        'lblLetter9
        '
        Me.lblLetter9.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter9.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter9.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter9.Location = New System.Drawing.Point(799, 520)
        Me.lblLetter9.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter9.Name = "lblLetter9"
        Me.lblLetter9.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter9.TabIndex = 83
        Me.lblLetter9.Text = "W"
        Me.lblLetter9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter9.Visible = False
        '
        'lblLetter10
        '
        Me.lblLetter10.BackColor = System.Drawing.Color.Transparent
        Me.lblLetter10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLetter10.Font = New System.Drawing.Font("Comic Sans MS", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetter10.ForeColor = System.Drawing.Color.Lavender
        Me.lblLetter10.Location = New System.Drawing.Point(895, 520)
        Me.lblLetter10.Margin = New System.Windows.Forms.Padding(0)
        Me.lblLetter10.Name = "lblLetter10"
        Me.lblLetter10.Size = New System.Drawing.Size(87, 87)
        Me.lblLetter10.TabIndex = 84
        Me.lblLetter10.Text = "W"
        Me.lblLetter10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLetter10.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(1006, 684)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblHeadingNumberOfWords)
        Me.Controls.Add(Me.lblHeadingCurrentWord)
        Me.Controls.Add(Me.lblLetter10)
        Me.Controls.Add(Me.lblLetter9)
        Me.Controls.Add(Me.lblLetter8)
        Me.Controls.Add(Me.lblLetter7)
        Me.Controls.Add(Me.lblLetter6)
        Me.Controls.Add(Me.lblLetter5)
        Me.Controls.Add(Me.lblLetter4)
        Me.Controls.Add(Me.lblLetter3)
        Me.Controls.Add(Me.lblLetter2)
        Me.Controls.Add(Me.lblLetter1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gbMistakes)
        Me.Controls.Add(Me.pbPrize)
        Me.Controls.Add(Me.btnNewGame)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.pbPic)
        Me.Controls.Add(Me.lblHeading1)
        Me.Controls.Add(Me.btnBye)
        Me.Controls.Add(Me.lblHeading2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Beans Word Game"
        Me.gbMistakes.ResumeLayout(False)
        CType(Me.pbPic, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbPrize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.pbFace, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

    Private Sub frmMain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        intNoActionLoop = 0
        If flgStatus <> "WaitForUser" Then Exit Sub
        flgStatus = "Processing"

        Dim lblLetters As Control() = ControlArrayUtils.getMixedControlArray(Me, "lblLetter")
        Dim strLetterPressed As String
        Dim strCurrentLetter As String

        strLetterPressed = e.KeyChar
        strCurrentLetter = Mid(strCurrentWord, intCurrentLetter, 1)

        If LCase(strLetterPressed) = LCase(strCurrentLetter) Then
            ' Correct 
            Call LogStatus("Correct Letter Pressed: " & strLetterPressed)
            lblLetters(intCurrentLetter).Text = strCurrentLetter
            lblLetters(intCurrentLetter).BackColor = Color.LightGreen
            pbFace.Image = Image.FromFile(App_Path() & "happy.gif")

            If strLetterAck = "Sounds" Then
                PlayWav(App_Path() & "correctBeep" & CInt(Int((5 * Rnd()) + 1)) & ".wav", False)
                Delay(0.75)
            Else
                PlayWav(App_Path() & "correct" & CInt(Int((5 * Rnd()) + 1)) & ".wav", False)
                Delay(1.25)
            End If

            lblLetters(intCurrentLetter).BackColor = Color.Transparent

            If intCurrentLetter = intCurrentWordLength Then
                ' Word is Done Right!
                Delay(0.5)
                PlayWav(App_Path() & "good" & CInt(Int((3 * Rnd()) + 1)) & ".wav", False)
                Delay(1)
                For intCnt2 = 1 To 3
                    For intCnt = 1 To intCurrentWordLength
                        lblLetters(intCnt).ForeColor = Color.Transparent
                        lblLetters(intCnt).BackColor = Color.Blue
                        Delay(0)
                    Next intCnt
                    For intCnt = 1 To intCurrentWordLength
                        lblLetters(intCnt).ForeColor = Color.Blue
                        lblLetters(intCnt).BackColor = Color.Transparent
                        Delay(0)
                    Next intCnt
                Next intCnt2

                ' Erase Word Boxes
                For intCnt = intCurrentWordLength To 1 Step -1
                    lblLetters(intCnt).Visible = False
                    Delay(0.25)
                Next intCnt

                pbPic.Image = Nothing
                pbPic.Visible = False

                Delay(0.5)

                ' See if game is over
                If intCurrentWord = intNumberOfWords Then
                    Call GameDone()
                Else
                    ' New Word
                    intCurrentWord = intCurrentWord + 1
                    intNoActionLoop = 0
                    flgStatus = "NewWord"
                    Call NewWord()
                End If
            Else
                ' Next Letter
                intCurrentLetter = intCurrentLetter + 1
                lblLetters(intCurrentLetter).BackColor = Color.LightBlue
                pbFace.Image = Image.FromFile(App_Path() & "wait.gif")
                flgStatus = "WaitForUser"
            End If
        Else
            ' Wrong
            Call LogStatus("Wrong Letter Pressed: " & strLetterPressed)
            If strLetterCase = "Upper Case" Then
                lblLetters(intCurrentLetter).Text = UCase(strLetterPressed)
            ElseIf strLetterCase = "Lower Case" Then
                lblLetters(intCurrentLetter).Text = LCase(strLetterPressed)
            Else
                lblLetters(intCurrentLetter).Text = strLetterPressed
            End If
            lblLetters(intCurrentLetter).BackColor = Color.LightCoral
            pbFace.Image = Image.FromFile(App_Path() & "sad.gif")
            intMistakes = intMistakes + 1
            lblMistakes.Text = intMistakes
            PlayWav(App_Path() & "wrong" & CInt(Int((3 * Rnd()) + 1)) & ".wav", False)
            Delay(3)
            lblLetters(intCurrentLetter).Text = ""
            lblLetters(intCurrentLetter).BackColor = Color.LightBlue
            pbFace.Image = Image.FromFile(App_Path() & "wait.gif")
            flgStatus = "WaitForUser"
        End If
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Make sure only one copy of program is running
        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then End

        Call LoadRegValues()
        Call OpenLogFile()
        Call LoadPicsInArray()
        Call InitGame()
        Call SetFromSettings() ' Potential screen show
        Call LoadArrayValuesToScreen() ' Potential screen show
        Call TitleDisplay() ' Purpose screen show

        timeMain.Enabled = True
        Call NewWord() ' Go!
    End Sub

    Private Sub LoadRegValues()
        ' Get Win Range Values from Reg
        intWin1 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Excellent", 0)
        intWin2 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Great", 3)
        intWin3 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Good", 6)
        intWin4 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "OK", 9)
        ' Go to defualts if reg settings are out of whack
        If intWin1 >= intWin2 Or intWin2 >= intWin3 Or intWin3 >= intWin4 Then
            intWin1 = 0
            intWin2 = 3
            intWin3 = 6
            intWin4 = 9
        End If

        ' Number of Words per Game
        intNumberWordsPerGame = GetSetting("Beans Word Game", "Settings", "Number of Words per Game", 10)
        If Not IsNumeric(intNumberWordsPerGame) Then intNumberWordsPerGame = 10
        If intNumberWordsPerGame < 1 Or intQuitGameTime > 99 Then intNumberWordsPerGame = 10

        ' Get Quit Game Time from Reg
        intQuitGameTime = GetSetting("Beans Word Game", "Settings", "Quit Game Time", 5)
        If intQuitGameTime < 3 Or intQuitGameTime > 30 Then intQuitGameTime = 5

        ' Get Case
        strLetterCase = GetSetting("Beans Word Game", "Settings", "Letter Case", "Mixed Case")
        If strLetterCase <> "Upper Case" And strLetterCase <> "Lower Case" Then strLetterCase = "Mixed Case"

        ' Get Correct Letter Acknowledgment 
        strLetterAck = GetSetting("Beans Word Game", "Settings", "Correct Letter Acknowledgment", "Words")
        If strLetterAck <> "Sounds" Then strLetterAck = "Words"

        ' Get Picture Force Fit Value from Reg
        flgForceFit = GetSetting("Beans Word Game", "Settings", "Force Fit", "N")
        If flgForceFit <> "Y" Then flgForceFit = "N"

        ' Get Give Hints Value from Reg
        flgGiveHints = GetSetting("Beans Word Game", "Settings", "Give Hints", "Y")
        If flgGiveHints <> "N" Then flgGiveHints = "Y"

        ' Get Show Letters Value from Reg
        flgShowLetters = GetSetting("Beans Word Game", "Settings", "Show Letters", "Y")
        If flgShowLetters <> "N" Then flgShowLetters = "Y"

        ' Get Say Letters Value from Reg
        flgSayLetters = GetSetting("Beans Word Game", "Settings", "Say Letters", "Y")
        If flgSayLetters <> "N" Then flgSayLetters = "Y"
    End Sub

    Private Sub OpenLogFile()
        strLogFile = App_Path() & "logs\GameLog" & Mid(DatePart(DateInterval.Year, Now), 3) & Format(DatePart(DateInterval.Month, Now), "00") & Format(DatePart(DateInterval.Day, Now), "00") & "-" & Format(DatePart(DateInterval.Hour, Now), "00") & Format(DatePart(DateInterval.Minute, Now), "00") & Format(DatePart(DateInterval.Second, Now), "00") & ".txt"
        If (Not System.IO.Directory.Exists(App_Path() & "logs\")) Then
            System.IO.Directory.CreateDirectory(App_Path() & "logs\")
        End If

        objWriter = New System.IO.StreamWriter(strLogFile)
        Call LogStatus("Start Beans Word Game")
    End Sub

    Private Sub LoadPicsInArray()
        ' Clear Main list (handy for playing second game)
        arrMainList.Clear()
        Dim diJpeg As New IO.DirectoryInfo("pics")
        Dim aryFiJpeg As IO.FileInfo()
        Dim fiJpeg As IO.FileInfo

        ' Load all jpegs in to temp table
        Try
            aryFiJpeg = diJpeg.GetFiles("*.jpg")
        Catch
            MsgBox("The folder For pictures (pics) was not found")
            End
        End Try

        Dim arrTempJpeg As New ArrayList()
        Dim intRndNumber As Integer
        Dim intMaxPics As Integer
        Dim intMaxChar As Integer = 10

        ' Load Max Pics from Reg (1-99)
        intMaxPics = intNumberWordsPerGame

        ' Move File Array into Temp Array
        intNumberOfWords = 0
        For Each fiJpeg In aryFiJpeg
            If Len(fiJpeg.Name) <= (intMaxChar + 4) Then
                intNumberOfWords = intNumberOfWords + 1
                If strLetterCase = "Upper Case" Then
                    arrTempJpeg.Add(UCase(fiJpeg.Name))
                ElseIf strLetterCase = "Lower Case" Then
                    arrTempJpeg.Add(LCase(fiJpeg.Name))
                Else
                    arrTempJpeg.Add(fiJpeg.Name)
                End If
            End If
        Next

        ' Move Temp Array into Main List - Random Style!
        For intCnt = 1 To arrTempJpeg.Count
GetAnotherPic:
            Randomize()
            intRndNumber = CInt(Int(arrTempJpeg.Count * Rnd()))
            If Len(arrTempJpeg(intRndNumber)) < 1 Then GoTo GetAnotherPic
            arrMainList.Add(arrTempJpeg(intRndNumber))
            arrTempJpeg(intRndNumber) = ""
        Next intCnt

        ' Done with Temp Array
        arrTempJpeg.Clear()

        If intNumberOfWords < 1 Then
            MsgBox("No Words were found (jpeg files in the pics folder)")
            End
        End If

        If intNumberOfWords > intMaxPics Then intNumberOfWords = intMaxPics
    End Sub

    Private Sub InitGame()
        intMistakes = 0
        lblMistakes.Text = intMistakes
        intCurrentWord = 1
        intNoActionLoop = 0
        flgStatus = "NewWord"
    End Sub

    Private Sub SetFromSettings()
        ' Set Picture Fit
        If flgForceFit = "Y" Then
            pbPic.SizeMode = PictureBoxSizeMode.StretchImage
        Else
            pbPic.SizeMode = PictureBoxSizeMode.CenterImage
        End If
    End Sub

    Private Sub LoadArrayValuesToScreen()
        ' Load Screen values for current word (1) and total words

        lblHeadingCurrentWord.Text = intCurrentWord
        lblHeadingNumberOfWords.Text = intNumberOfWords

        ' Make screen values visible
        lblHeadingCurrentWord.Visible = True
        lblHeading1.Visible = True
        lblHeading2.Visible = True
        lblHeadingNumberOfWords.Visible = True
    End Sub

    Private Sub TitleDisplay()
        Dim lblLetters As Control() = ControlArrayUtils.getMixedControlArray(Me, "lblLetter")

        ' Set Title Box size
        lblTitle.Width = 952
        lblTitle.Height = 425

        ' Show Form
        Me.Visible = True

        ' Start Game Happy
        pbFace.Image = Image.FromFile(App_Path() & "happy.gif")

        Delay(1)

        ' Big Intro Title
        lblTitle.Text = "Beans" & Chr(10) & Chr(13) & "Word Game"
        For intCnt = 2 To 100 Step 2
            lblTitle.Font = New Font("Comic Sans MS", intCnt, FontStyle.Regular)
            Delay(0)
        Next intCnt

        PlayWav(App_Path() & "intro1.wav", False)

        ' Setup Screen Letters
        For intCnt = 1 To 10
            lblLetters(intCnt).Font = New Font("Comic Sans MS", 48, FontStyle.Regular)
            lblLetters(intCnt).ForeColor = System.Drawing.Color.Red
            lblLetters(intCnt).Text = Chr(64 + intCnt)
            lblLetters(intCnt).Visible = True
            Delay(0)
        Next intCnt

        Delay(0)
        For intCnt = 10 To 1 Step -1
            lblLetters(intCnt).ForeColor = System.Drawing.Color.LightGreen
            Delay(0)
        Next intCnt
        Delay(0)
        For intCnt = 1 To 10
            lblLetters(intCnt).ForeColor = System.Drawing.Color.Orange
            Delay(0)
        Next intCnt
        Delay(0)
        For intCnt = 10 To 1 Step -1
            lblLetters(intCnt).ForeColor = System.Drawing.Color.Yellow
            Delay(0)
        Next intCnt
        Delay(0)
        For intCnt = 1 To 10
            lblLetters(intCnt).ForeColor = System.Drawing.Color.Blue
            Delay(0)
        Next intCnt

        For intCnt = 100 To 2 Step -2
            lblTitle.Font = New Font("Comic Sans MS", intCnt, FontStyle.Regular)
            Delay(0)
        Next intCnt

        ' Remove Title Heading Box
        lblTitle.Width = 1
        lblTitle.Height = 1

        ' Reset Letter Boxes
        For intCnt = 1 To 10
            lblLetters(intCnt).Visible = False
            lblLetters(intCnt).Text = ""
            Delay(0)
        Next intCnt

        PlayWav(App_Path() & "intro2.wav", False)
    End Sub

    Private Sub NewWord()
        Dim lblLetters As Control() = ControlArrayUtils.getMixedControlArray(Me, "lblLetter")

        flgStatus = "Processing"

        ' Update Display Number
        lblHeadingCurrentWord.Text = intCurrentWord

        ' Grab Current Word
        intCurrentWordLength = Len(arrMainList.Item(intCurrentWord - 1)) - 4
        strCurrentWord = Mid(arrMainList.Item(intCurrentWord - 1), 1, intCurrentWordLength)

        ' Show Blank Letters Boxes
        For intCnt = 1 To intCurrentWordLength
            ' Letter Box placement (Start Pos + Each Boxes move to the Left
            lblLetters(intCnt).Left = (511 - (intCurrentWordLength * 48)) + ((intCnt - 1) * 96)
            lblLetters(intCnt).Text = ""
            lblLetters(intCnt).Visible = True
            Delay(0.25)
        Next

        Delay(1)

        ' Show Current Word Picture
        Try
            pbPic.Image = Image.FromFile(App_Path() & "pics\" & arrMainList.Item(intCurrentWord - 1))
            pbPic.Visible = True
        Catch
            pbPic.Image = Nothing
            pbPic.Visible = True
        End Try

        Delay(1)

        ' Check if custom WAV file exist
        If System.IO.File.Exists(App_Path() & "wavs\" & strCurrentWord & ".wav") = True Then
            ' If Yes, play custom WAV
            PlayWav(App_Path() & "spell" & CInt(Int((3 * Rnd()) + 1)) & ".wav", True)
            ' No Delay
            PlayWav(App_Path() & "wavs\" & strCurrentWord & ".wav", False)
        Else
            ' If No, Play generic WAV
            PlayWav(App_Path() & "spellNoWav" & CInt(Int((3 * Rnd()) + 1)) & ".wav", True) 'jjb true
        End If

        ' Fill Letter Boxes
        If flgShowLetters = "Y" Or flgSayLetters = "Y" Then
            Delay(1)
            For intCnt = 1 To intCurrentWordLength
                If flgShowLetters = "Y" Then
                    lblLetters(intCnt).Text = Mid(strCurrentWord, intCnt, 1)
                End If
                If flgSayLetters = "Y" Then
                    If Mid(strCurrentWord, intCnt, 1) = " " Then
                        PlayWav(App_Path() & "space.wav", False)
                    ElseIf Mid(strCurrentWord, intCnt, 1) = "-" Then
                        PlayWav(App_Path() & "hyphen.wav", False)
                    Else
                        PlayWav(App_Path() & Mid(strCurrentWord, intCnt, 1) & ".wav", False)
                    End If
                End If
                Delay(1.25)
            Next
            Delay(0.5)
            ' Check if custom WAV file exist, if yes say word again
            If System.IO.File.Exists(App_Path() & "wavs\" & strCurrentWord & ".wav") = True Then
                PlayWav(App_Path() & "wavs\" & strCurrentWord & ".wav", True) ' Repete word
            End If
        End If

        ' Update Log with Picture info
        Call LogStatus("* Current Word: " & strCurrentWord & " *")

        ' Empty Letter Boxes
        If flgShowLetters = "Y" Then
            Delay(1)
            For intCnt = intCurrentWordLength To 1 Step -1
                lblLetters(intCnt).Text = ""
                Delay(0.25)
            Next
        End If

        Delay(0.5)

        ' Go live game
        intCurrentLetter = 1

        ' Set First Box for letter
        pbFace.Image = Image.FromFile(App_Path() & "wait.gif") ' New Word so face will be in Wait mode
        lblLetters(1).BackColor = Color.LightBlue

        flgStatus = "WaitForUser"
    End Sub

    Private Sub timeMain_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timeMain.Tick
        Dim lblLetters As Control() = ControlArrayUtils.getMixedControlArray(Me, "lblLetter")

        If flgStatus = "WaitForUser" Then
            intNoActionLoop = intNoActionLoop + 1

            ' Reset face to Wait at 12 seconds
            If intNoActionLoop = 4 Then
                pbFace.Image = Image.FromFile(App_Path() & "wait.gif")
            End If

            ' Blink every 69 secounds
            If intNoActionLoop Mod 23 = 0 Then
                pbFace.Image = Image.FromFile(App_Path() & "blink.gif") ' Make blink
                Delay(1)
                If flgStatus = "WaitForUser" Then pbFace.Image = Image.FromFile(App_Path() & "wait.gif") ' Unblink
            End If

            ' End Game at 3, 5, 10, 15 or 30 minutes (intQuitGameTime is minutes and loop is 3 seconds)
            If intNoActionLoop = intQuitGameTime * 20 Then
                pbFace.Image = Image.FromFile(App_Path() & "tongue.gif") ' Make tongue
                Delay(1)
                Call LogStatus("No User Action")
                Call CloseGame()
            End If

            ' Give Hints
            If intNoActionLoop Mod 10 = 0 Then
                If intNoActionLoop = 10 Then ' Always give generic hint at 30 sec
                    PlayWav(App_Path() & "hint" & CInt(Int((2 * Rnd()) + 1)) & ".wav", False)
                Else
                    If flgGiveHints = "Y" Then
                        If CInt(Int((2 * Rnd()) + 1)) = 1 Then ' 50/50 chance of specific vs generic
                            PlayWav(App_Path() & "hint" & CInt(Int((2 * Rnd()) + 1)) & ".wav", False)
                        Else
                            PlayWav(App_Path() & "hintBig" & CInt(Int((2 * Rnd()) + 1)) & ".wav", True)
                            If Mid(strCurrentWord, intCurrentLetter, 1) = " " Then
                                PlayWav(App_Path() & "space.wav", True)
                            ElseIf Mid(strCurrentWord, intCurrentLetter, 1) = "-" Then
                                PlayWav(App_Path() & "hyphen.wav", True)
                            Else
                                PlayWav(App_Path() & Mid(strCurrentWord, intCurrentLetter, 1) & ".wav", True)
                            End If
                        End If
                    Else
                        PlayWav(App_Path() & "hint" & CInt(Int((2 * Rnd()) + 1)) & ".wav", False)
                    End If
                End If ' If intNoActionLoop = 10
                Call LogStatus("User Delay")
            End If ' If intNoActionLoop Mod 10 = 0 Then

        End If ' If flgStatus = "WaitForUser" Then
    End Sub

    Private Sub btnNewGame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewGame.Click
        objWriter.Close() ' Close Log File
        pbPrize.BorderStyle = BorderStyle.None
        pbPrize.Width = 1
        pbPrize.Height = 1

        btnNewGame.Enabled = False

        Delay(1)

        Call OpenLogFile()
        Call LoadPicsInArray()
        Call InitGame()
        Call LoadArrayValuesToScreen() ' Potential screen show

        timeMain.Enabled = True
        Me.Focus()
        Call NewWord() ' Go!
    End Sub

    Private Sub btnBye_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBye.Click
        Call CloseGame()
    End Sub

    Private Sub GameDone()
        pbPic.Visible = False
        timeMain.Enabled = False
        Call LogStatus("Game Done - Mistakes " & intMistakes)

        pbPrize.Width = 944
        pbPrize.Height = 528
        pbPrize.BorderStyle = BorderStyle.FixedSingle

        If intMistakes >= 0 And intMistakes <= intWin1 Then
            pbPrize.Image = Image.FromFile(App_Path() & "win1.jpg")
            PlayWav(App_Path() & "win1.wav", False)
        End If
        If intMistakes >= intWin1 + 1 And intMistakes <= intWin2 Then
            pbPrize.Image = Image.FromFile(App_Path() & "win2.jpg")
            PlayWav(App_Path() & "win2.wav", False)
        End If
        If intMistakes >= intWin2 + 1 And intMistakes <= intWin3 Then
            pbPrize.Image = Image.FromFile(App_Path() & "win3.jpg")
            PlayWav(App_Path() & "win3.wav", False)
        End If
        If intMistakes >= intWin3 + 1 And intMistakes <= intWin4 Then
            pbPrize.Image = Image.FromFile(App_Path() & "win4.jpg")
            PlayWav(App_Path() & "win4.wav", False)
        End If
        If intMistakes >= intWin4 + 1 Then ' 
            pbPrize.Image = Image.FromFile(App_Path() & "win5.jpg")
            PlayWav(App_Path() & "win5.wav", False)
        End If

        Delay(1)

        btnNewGame.Enabled = True
    End Sub

    Private Sub CloseGame()
        If flgStatus = "Done" Then Exit Sub
        flgStatus = "Done"
        PlayWav(App_Path() & "bye.wav", False)
        Call LogStatus("Closing Beans Word Game")
        Delay(1)
        objWriter.Close() ' Close Log File
        End
    End Sub

    Public Function App_Path() As String
        Return System.AppDomain.CurrentDomain.BaseDirectory()
    End Function

    Public Sub Delay(ByVal DelayInSeconds As Double)
        Dim ts As TimeSpan
        Dim targetTime As DateTime = DateTime.Now.AddSeconds(DelayInSeconds)
        Do
            ts = targetTime.Subtract(DateTime.Now)
            Application.DoEvents() ' keep app responsive
            System.Threading.Thread.Sleep(50) ' reduce CPU usage
        Loop While ts.TotalSeconds > 0
    End Sub

    Public Function PlayWav(ByVal fileFullPath As String, Optional ByVal flgWait As Boolean = True) As Boolean
        'return true if successful, false if otherwise
        ' +&h1 makes the file play asynchronously 
        Dim iRet As Integer = 0
        Try
            If flgWait = True Then
                iRet = PlaySound(fileFullPath, 0, SND_FILENAME)
            Else
                ' flgWait = False, then the program does not wait for the wav file to finsih playing
                iRet = PlaySound(fileFullPath, 0, SND_FILENAME + &H1)
            End If
        Catch
        End Try
        Return iRet
    End Function

    Private Sub LogStatus(ByVal psStatus As String)
        objWriter.WriteLine(Now & " - " & psStatus)
    End Sub
End Class

