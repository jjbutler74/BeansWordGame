Public Class frmSettings
    Inherits System.Windows.Forms.Form

    ' Win Ranges
    Dim intWin1 As Integer
    Dim intWin2 As Integer
    Dim intWin3 As Integer
    Dim intWin4 As Integer

    ' Number of Words per Game
    Dim intNumberWords As Integer

    ' Quit Game Time
    Dim intQuitGameTime As Integer

    ' Quit Game Time
    Dim strLetterCase As String

    ' Correct Letter Acknowledgment 
    Dim strLetterAck As String

    ' Force Fit Flag
    Dim flgForceFit As String

    ' Hints
    Dim flgGiveHints As String
    Dim flgShowLetters As String
    Friend WithEvents cmbLetterAck As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Dim flgSayLetters As String

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
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmbQuitGameTime As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtNumberWords As System.Windows.Forms.TextBox
    Friend WithEvents txtExcellentLow As System.Windows.Forms.TextBox
    Friend WithEvents txtExcellentHigh As System.Windows.Forms.TextBox
    Friend WithEvents txtPoorLow As System.Windows.Forms.TextBox
    Friend WithEvents txtGreatHigh As System.Windows.Forms.TextBox
    Friend WithEvents txtGreatLow As System.Windows.Forms.TextBox
    Friend WithEvents txtOKHigh As System.Windows.Forms.TextBox
    Friend WithEvents txtOKLow As System.Windows.Forms.TextBox
    Friend WithEvents txtGoodHigh As System.Windows.Forms.TextBox
    Friend WithEvents txtGoodLow As System.Windows.Forms.TextBox
    Friend WithEvents chkForceFit As System.Windows.Forms.CheckBox
    Friend WithEvents chkGiveHints As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLetters As System.Windows.Forms.CheckBox
    Friend WithEvents chkSayLetters As System.Windows.Forms.CheckBox
    Friend WithEvents cmbLetterCase As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.txtExcellentLow = New System.Windows.Forms.TextBox
        Me.txtExcellentHigh = New System.Windows.Forms.TextBox
        Me.txtPoorLow = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtGreatHigh = New System.Windows.Forms.TextBox
        Me.txtGreatLow = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtOKHigh = New System.Windows.Forms.TextBox
        Me.txtOKLow = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtGoodHigh = New System.Windows.Forms.TextBox
        Me.txtGoodLow = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.chkForceFit = New System.Windows.Forms.CheckBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.chkGiveHints = New System.Windows.Forms.CheckBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.chkShowLetters = New System.Windows.Forms.CheckBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.chkSayLetters = New System.Windows.Forms.CheckBox
        Me.cmbQuitGameTime = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.txtNumberWords = New System.Windows.Forms.TextBox
        Me.cmbLetterCase = New System.Windows.Forms.ComboBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.cmbLetterAck = New System.Windows.Forms.ComboBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnExit.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(238, 487)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(100, 33)
        Me.btnExit.TabIndex = 14
        Me.btnExit.Text = "&Cancel"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnSave.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(88, 487)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(100, 33)
        Me.btnSave.TabIndex = 13
        Me.btnSave.Text = "&Save Settings"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'txtExcellentLow
        '
        Me.txtExcellentLow.BackColor = System.Drawing.Color.Gainsboro
        Me.txtExcellentLow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExcellentLow.Location = New System.Drawing.Point(202, 51)
        Me.txtExcellentLow.Name = "txtExcellentLow"
        Me.txtExcellentLow.ReadOnly = True
        Me.txtExcellentLow.Size = New System.Drawing.Size(40, 22)
        Me.txtExcellentLow.TabIndex = 3
        Me.txtExcellentLow.TabStop = False
        '
        'txtExcellentHigh
        '
        Me.txtExcellentHigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExcellentHigh.Location = New System.Drawing.Point(290, 51)
        Me.txtExcellentHigh.MaxLength = 3
        Me.txtExcellentHigh.Name = "txtExcellentHigh"
        Me.txtExcellentHigh.Size = New System.Drawing.Size(40, 22)
        Me.txtExcellentHigh.TabIndex = 1
        '
        'txtPoorLow
        '
        Me.txtPoorLow.BackColor = System.Drawing.Color.Gainsboro
        Me.txtPoorLow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPoorLow.Location = New System.Drawing.Point(202, 179)
        Me.txtPoorLow.Name = "txtPoorLow"
        Me.txtPoorLow.ReadOnly = True
        Me.txtPoorLow.Size = New System.Drawing.Size(40, 22)
        Me.txtPoorLow.TabIndex = 5
        Me.txtPoorLow.TabStop = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(258, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 16)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "to"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(258, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(26, 16)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "to"
        '
        'txtGreatHigh
        '
        Me.txtGreatHigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGreatHigh.Location = New System.Drawing.Point(290, 83)
        Me.txtGreatHigh.MaxLength = 3
        Me.txtGreatHigh.Name = "txtGreatHigh"
        Me.txtGreatHigh.Size = New System.Drawing.Size(40, 22)
        Me.txtGreatHigh.TabIndex = 2
        '
        'txtGreatLow
        '
        Me.txtGreatLow.BackColor = System.Drawing.Color.Gainsboro
        Me.txtGreatLow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGreatLow.Location = New System.Drawing.Point(202, 83)
        Me.txtGreatLow.Name = "txtGreatLow"
        Me.txtGreatLow.ReadOnly = True
        Me.txtGreatLow.Size = New System.Drawing.Size(40, 22)
        Me.txtGreatLow.TabIndex = 7
        Me.txtGreatLow.TabStop = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(258, 147)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 16)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "to"
        '
        'txtOKHigh
        '
        Me.txtOKHigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOKHigh.Location = New System.Drawing.Point(290, 147)
        Me.txtOKHigh.MaxLength = 3
        Me.txtOKHigh.Name = "txtOKHigh"
        Me.txtOKHigh.Size = New System.Drawing.Size(40, 22)
        Me.txtOKHigh.TabIndex = 4
        '
        'txtOKLow
        '
        Me.txtOKLow.BackColor = System.Drawing.Color.Gainsboro
        Me.txtOKLow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOKLow.Location = New System.Drawing.Point(202, 147)
        Me.txtOKLow.Name = "txtOKLow"
        Me.txtOKLow.ReadOnly = True
        Me.txtOKLow.Size = New System.Drawing.Size(40, 22)
        Me.txtOKLow.TabIndex = 13
        Me.txtOKLow.TabStop = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(258, 115)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(26, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "to"
        '
        'txtGoodHigh
        '
        Me.txtGoodHigh.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGoodHigh.Location = New System.Drawing.Point(290, 115)
        Me.txtGoodHigh.MaxLength = 3
        Me.txtGoodHigh.Name = "txtGoodHigh"
        Me.txtGoodHigh.Size = New System.Drawing.Size(40, 22)
        Me.txtGoodHigh.TabIndex = 3
        '
        'txtGoodLow
        '
        Me.txtGoodLow.BackColor = System.Drawing.Color.Gainsboro
        Me.txtGoodLow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGoodLow.Location = New System.Drawing.Point(202, 115)
        Me.txtGoodLow.Name = "txtGoodLow"
        Me.txtGoodLow.ReadOnly = True
        Me.txtGoodLow.Size = New System.Drawing.Size(40, 22)
        Me.txtGoodLow.TabIndex = 10
        Me.txtGoodLow.TabStop = False
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(258, 179)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(16, 16)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "+"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(68, 54)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 16)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Excellent Range "
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(68, 86)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 16)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Great Range "
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(68, 118)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 16)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Good Range "
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(68, 150)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(104, 16)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "OK Range "
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(68, 182)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(104, 16)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "Poor Range "
        '
        'chkForceFit
        '
        Me.chkForceFit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkForceFit.Location = New System.Drawing.Point(290, 348)
        Me.chkForceFit.Name = "chkForceFit"
        Me.chkForceFit.Size = New System.Drawing.Size(24, 24)
        Me.chkForceFit.TabIndex = 9
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(68, 351)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(168, 16)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "Force Pictures to Fit Boxes"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(202, 19)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(136, 16)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Number of Mistakes"
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(68, 223)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(216, 16)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Number of Words per Game"
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(68, 383)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(168, 16)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Give Hints"
        '
        'chkGiveHints
        '
        Me.chkGiveHints.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGiveHints.Location = New System.Drawing.Point(290, 380)
        Me.chkGiveHints.Name = "chkGiveHints"
        Me.chkGiveHints.Size = New System.Drawing.Size(24, 24)
        Me.chkGiveHints.TabIndex = 10
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(68, 415)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(168, 16)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Show Letters at First"
        '
        'chkShowLetters
        '
        Me.chkShowLetters.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowLetters.Location = New System.Drawing.Point(290, 412)
        Me.chkShowLetters.Name = "chkShowLetters"
        Me.chkShowLetters.Size = New System.Drawing.Size(24, 24)
        Me.chkShowLetters.TabIndex = 11
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(68, 447)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(168, 16)
        Me.Label16.TabIndex = 32
        Me.Label16.Text = "Say Letters at First"
        '
        'chkSayLetters
        '
        Me.chkSayLetters.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSayLetters.Location = New System.Drawing.Point(290, 444)
        Me.chkSayLetters.Name = "chkSayLetters"
        Me.chkSayLetters.Size = New System.Drawing.Size(24, 24)
        Me.chkSayLetters.TabIndex = 12
        '
        'cmbQuitGameTime
        '
        Me.cmbQuitGameTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmbQuitGameTime.Items.AddRange(New Object() {"3", "5", "10", "15", "30"})
        Me.cmbQuitGameTime.Location = New System.Drawing.Point(290, 252)
        Me.cmbQuitGameTime.Name = "cmbQuitGameTime"
        Me.cmbQuitGameTime.Size = New System.Drawing.Size(100, 24)
        Me.cmbQuitGameTime.TabIndex = 6
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(68, 255)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(216, 16)
        Me.Label17.TabIndex = 33
        Me.Label17.Text = "Minutes of Inactivity before Quitting"
        '
        'txtNumberWords
        '
        Me.txtNumberWords.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumberWords.Location = New System.Drawing.Point(290, 220)
        Me.txtNumberWords.MaxLength = 2
        Me.txtNumberWords.Name = "txtNumberWords"
        Me.txtNumberWords.Size = New System.Drawing.Size(40, 22)
        Me.txtNumberWords.TabIndex = 5
        '
        'cmbLetterCase
        '
        Me.cmbLetterCase.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmbLetterCase.Items.AddRange(New Object() {"Mixed Case", "Upper Case", "Lower Case"})
        Me.cmbLetterCase.Location = New System.Drawing.Point(290, 284)
        Me.cmbLetterCase.Name = "cmbLetterCase"
        Me.cmbLetterCase.Size = New System.Drawing.Size(100, 24)
        Me.cmbLetterCase.TabIndex = 7
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(68, 287)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(216, 16)
        Me.Label18.TabIndex = 35
        Me.Label18.Text = "Letter Case"
        '
        'cmbLetterAck
        '
        Me.cmbLetterAck.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.cmbLetterAck.Items.AddRange(New Object() {"Words", "Sounds"})
        Me.cmbLetterAck.Location = New System.Drawing.Point(290, 316)
        Me.cmbLetterAck.Name = "cmbLetterAck"
        Me.cmbLetterAck.Size = New System.Drawing.Size(100, 24)
        Me.cmbLetterAck.TabIndex = 8
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(68, 319)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(216, 16)
        Me.Label19.TabIndex = 37
        Me.Label19.Text = "Correct Letter Acknowledgment"
        '
        'frmSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(427, 541)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmbLetterAck)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.cmbLetterCase)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.txtNumberWords)
        Me.Controls.Add(Me.cmbQuitGameTime)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.chkSayLetters)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.chkShowLetters)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.chkGiveHints)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.chkForceFit)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtOKHigh)
        Me.Controls.Add(Me.txtOKLow)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtGoodHigh)
        Me.Controls.Add(Me.txtGoodLow)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtGreatHigh)
        Me.Controls.Add(Me.txtGreatLow)
        Me.Controls.Add(Me.txtPoorLow)
        Me.Controls.Add(Me.txtExcellentHigh)
        Me.Controls.Add(Me.txtExcellentLow)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Beans Word Game Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Get Values from Reg
        intWin1 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Excellent", 0)
        intWin2 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Great", 3)
        intWin3 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "Good", 6)
        intWin4 = GetSetting("Beans Word Game", "Mistake Ranges (High)", "OK", 9)

        intNumberWords = GetSetting("Beans Word Game", "Settings", "Number of Words per Game", 10)
        intQuitGameTime = GetSetting("Beans Word Game", "Settings", "Quit Game Time", 5)
        strLetterCase = GetSetting("Beans Word Game", "Settings", "Letter Case", "Mixed Case")
        strLetterAck = GetSetting("Beans Word Game", "Settings", "Correct Letter Acknowledgment", "Words")

        flgForceFit = GetSetting("Beans Word Game", "Settings", "Force Fit", "N")
        flgGiveHints = GetSetting("Beans Word Game", "Settings", "Give Hints", "Y")
        flgShowLetters = GetSetting("Beans Word Game", "Settings", "Show Letters", "Y")
        flgSayLetters = GetSetting("Beans Word Game", "Settings", "Say Letters", "Y")

        ' Go to defualts if reg settings are out of whack
        If intWin1 >= intWin2 Or intWin2 >= intWin3 Or intWin3 >= intWin4 Then
            intWin1 = 0
            intWin2 = 3
            intWin3 = 6
            intWin4 = 9
        End If

        If intQuitGameTime <> 3 And intQuitGameTime <> 5 And intQuitGameTime <> 10 And intQuitGameTime <> 15 And intQuitGameTime <> 30 Then intQuitGameTime = 5
        If Not IsNumeric(intNumberWords) Then intNumberWords = 10
        If intNumberWords < 1 Or intNumberWords > 99 Then intNumberWords = 10
        If strLetterCase <> "Upper Case" And strLetterCase <> "Lower Case" Then strLetterCase = "Mixed Case"
        If strLetterAck <> "Sounds" Then strLetterAck = "Words"

        If flgForceFit <> "Y" Then flgForceFit = "N"
        If flgGiveHints <> "N" Then flgGiveHints = "Y"
        If flgShowLetters <> "N" Then flgShowLetters = "Y"
        If flgSayLetters <> "N" Then flgSayLetters = "Y"

        ' Put Reg Values on the Form
        txtExcellentLow.Text = 0
        txtExcellentHigh.Text = intWin1
        txtGreatLow.Text = intWin1 + 1
        txtGreatHigh.Text = intWin2
        txtGoodLow.Text = intWin2 + 1
        txtGoodHigh.Text = intWin3
        txtOKLow.Text = intWin3 + 1
        txtOKHigh.Text = intWin4
        txtPoorLow.Text = intWin4 + 1

        txtNumberWords.Text = intNumberWords
        cmbQuitGameTime.Text = intQuitGameTime
        cmbLetterCase.Text = strLetterCase
        cmbLetterAck.Text = strLetterAck

        If flgForceFit = "Y" Then
            chkForceFit.Checked = True
        Else
            chkForceFit.Checked = False
        End If

        If flgGiveHints = "Y" Then
            chkGiveHints.Checked = True
        Else
            chkGiveHints.Checked = False
        End If

        If flgShowLetters = "Y" Then
            chkShowLetters.Checked = True
        Else
            chkShowLetters.Checked = False
        End If

        If flgSayLetters = "Y" Then
            chkSayLetters.Checked = True
        Else
            chkSayLetters.Checked = False
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        ' Validation Time
        If Not IsNumeric(txtExcellentHigh.Text) Then
            MsgBox("The High value in the Excellent Range must be Numeric", MsgBoxStyle.Exclamation)
            txtExcellentHigh.Focus()
            Exit Sub
        End If
        If Not IsNumeric(txtGreatHigh.Text) Then
            MsgBox("The High value in the Great Range must be Numeric", MsgBoxStyle.Exclamation)
            txtGreatHigh.Focus()
            Exit Sub
        End If
        If Not IsNumeric(txtGoodHigh.Text) Then
            MsgBox("The High value in the Good Range must be Numeric", MsgBoxStyle.Exclamation)
            txtGoodHigh.Focus()
            Exit Sub
        End If
        If Not IsNumeric(txtOKHigh.Text) Then
            MsgBox("The High value in the OK Range must be Numeric", MsgBoxStyle.Exclamation)
            txtOKHigh.Focus()
            Exit Sub
        End If

        If CInt(txtExcellentHigh.Text) < 0 Then
            MsgBox("You can't have a negative number here silly", MsgBoxStyle.Exclamation)
            txtExcellentHigh.Focus()
            Exit Sub
        End If

        If CInt(txtOKHigh.Text) > 999 Then
            MsgBox("Don't you think this number is just a little bit too BIG", MsgBoxStyle.Exclamation)
            txtOKHigh.Focus()
            Exit Sub
        End If

        If CInt(txtExcellentHigh.Text) >= CInt(txtGreatHigh.Text) Then
            MsgBox("The High value in the Great Range must be bigger than the High value in the Excellent Range", MsgBoxStyle.Exclamation)
            txtGreatHigh.Focus()
            Exit Sub
        End If
        If CInt(txtGreatHigh.Text) >= CInt(txtGoodHigh.Text) Then
            MsgBox("The High value in the Good Range must be bigger than the High value in the Great Range", MsgBoxStyle.Exclamation)
            txtGoodHigh.Focus()
            Exit Sub
        End If
        If CInt(txtGoodHigh.Text) >= CInt(txtOKHigh.Text) Then
            MsgBox("The High value in the OK Range must be bigger than the High value in the Good Range", MsgBoxStyle.Exclamation)
            txtOKHigh.Focus()
            Exit Sub
        End If
        If Not IsNumeric(txtNumberWords.Text) Then
            MsgBox("Number of Words per Game must be Numeric", MsgBoxStyle.Exclamation)
            txtNumberWords.Focus()
            Exit Sub
        End If
        If txtNumberWords.Text < 1 Or txtNumberWords.Text > 99 Then
            MsgBox("Number of Words per Game must be between 1 and 99", MsgBoxStyle.Exclamation)
            txtNumberWords.Focus()
            Exit Sub
        End If

        ' Put Form Values in to Vars
        intWin1 = txtExcellentHigh.Text
        intWin2 = txtGreatHigh.Text
        intWin3 = txtGoodHigh.Text
        intWin4 = txtOKHigh.Text

        intNumberWords = txtNumberWords.Text
        intQuitGameTime = cmbQuitGameTime.Text
        strLetterCase = cmbLetterCase.Text
        strLetterAck = cmbLetterAck.Text

        If chkForceFit.Checked = True Then
            flgForceFit = "Y"
        Else
            flgForceFit = "N"
        End If

        If chkGiveHints.Checked = True Then
            flgGiveHints = "Y"
        Else
            flgGiveHints = "N"
        End If

        If chkShowLetters.Checked = True Then
            flgShowLetters = "Y"
        Else
            flgShowLetters = "N"
        End If

        If chkSayLetters.Checked = True Then
            flgSayLetters = "Y"
        Else
            flgSayLetters = "N"
        End If

        ' Update Un-editable Form Values
        txtGreatLow.Text = txtExcellentHigh.Text + 1
        txtGoodLow.Text = txtGreatHigh.Text + 1
        txtOKLow.Text = txtGoodHigh.Text + 1
        txtPoorLow.Text = txtOKHigh.Text + 1

        Delay(1)

        ' Save Reg Values
        SaveSetting("Beans Word Game", "Mistake Ranges (High)", "Excellent", intWin1)
        SaveSetting("Beans Word Game", "Mistake Ranges (High)", "Great", intWin2)
        SaveSetting("Beans Word Game", "Mistake Ranges (High)", "Good", intWin3)
        SaveSetting("Beans Word Game", "Mistake Ranges (High)", "OK", intWin4)

        SaveSetting("Beans Word Game", "Settings", "Number of Words per Game", intNumberWords)
        SaveSetting("Beans Word Game", "Settings", "Quit Game Time", intQuitGameTime)
        SaveSetting("Beans Word Game", "Settings", "Letter Case", strLetterCase)
        SaveSetting("Beans Word Game", "Settings", "Correct Letter Acknowledgment", strLetterAck)

        SaveSetting("Beans Word Game", "Settings", "Force Fit", flgForceFit)
        SaveSetting("Beans Word Game", "Settings", "Give Hints", flgGiveHints)
        SaveSetting("Beans Word Game", "Settings", "Show Letters", flgShowLetters)
        SaveSetting("Beans Word Game", "Settings", "Say Letters", flgSayLetters)

        End
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    Public Sub Delay(ByVal DelayInSeconds As Double)
        Dim ts As TimeSpan
        Dim targetTime As DateTime = DateTime.Now.AddSeconds(DelayInSeconds)
        Do
            ts = targetTime.Subtract(DateTime.Now)
            Application.DoEvents() ' keep app responsive
            System.Threading.Thread.Sleep(50) ' reduce CPU usage
        Loop While ts.TotalSeconds > 0
    End Sub
End Class
