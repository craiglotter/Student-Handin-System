Imports Microsoft.Win32
Imports System.IO


Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerComplete(ByVal Result As String)
    Public Delegate Sub WorkerErrorEncountered(ByVal ex As Exception)


    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private currently_working As Boolean = False

    Private InitialRootDirectory As String = ""
    Protected Friend WithEvents Current_Context As System.Windows.Forms.Label
    Private RootDirectory As String = ""

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen, Optional ByVal currentuser As String = "Unknown User", Optional ByVal currentcontext As String = "Students.com.main.uct")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Current_User.Text = currentuser.ToUpper
        Current_Context.Text = currentcontext
        splash_loader = splash
        Worker1 = New Worker
        ' AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        'AddHandler Worker1.WorkerFolderCount, AddressOf WorkerFolderCountHandler
        ' AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        ' AddHandler Worker1.WorkerMessageExtracted, AddressOf WorkerMessageExtractedHandler
        'AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
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
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Protected Friend WithEvents Current_User As System.Windows.Forms.Label
    Protected Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lstDepartment As MTGCComboBox
    Friend WithEvents lstDepartmentCover As System.Windows.Forms.Label
    Friend WithEvents statsDepartment As System.Windows.Forms.Label
    Friend WithEvents statsCourses As System.Windows.Forms.Label
    Friend WithEvents lstCoursesCover As System.Windows.Forms.Label
    Friend WithEvents lstCourses As MTGCComboBox
    Friend WithEvents RegisteredCourses As System.Windows.Forms.ListBox
    Friend WithEvents statsCoursesTotal As System.Windows.Forms.Label
    Friend WithEvents statsAssignmentsTotal As System.Windows.Forms.Label
    Friend WithEvents statsAssignments As System.Windows.Forms.Label
    Friend WithEvents lstAssignmentsCover As System.Windows.Forms.Label
    Friend WithEvents lstAssignments As MTGCComboBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnProceed As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CommunicationLabel1 As System.Windows.Forms.Label
    Friend WithEvents CommunicationLabel2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents statsDepartmentTotal As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblProceed As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Screen))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Current_User = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.btnProceed = New System.Windows.Forms.Button
        Me.RegisteredCourses = New System.Windows.Forms.ListBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.lblProceed = New System.Windows.Forms.Label
        Me.CommunicationLabel2 = New System.Windows.Forms.Label
        Me.CommunicationLabel1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lstAssignmentsCover = New System.Windows.Forms.Label
        Me.lstCoursesCover = New System.Windows.Forms.Label
        Me.statsDepartmentTotal = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lstAssignments = New MTGCComboBox
        Me.lstDepartmentCover = New System.Windows.Forms.Label
        Me.lstDepartment = New MTGCComboBox
        Me.statsDepartment = New System.Windows.Forms.Label
        Me.lstCourses = New MTGCComboBox
        Me.statsCourses = New System.Windows.Forms.Label
        Me.statsCoursesTotal = New System.Windows.Forms.Label
        Me.statsAssignments = New System.Windows.Forms.Label
        Me.statsAssignmentsTotal = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Current_Context = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Current_User
        '
        Me.Current_User.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Current_User.Location = New System.Drawing.Point(80, 8)
        Me.Current_User.Name = "Current_User"
        Me.Current_User.Size = New System.Drawing.Size(192, 16)
        Me.Current_User.TabIndex = 69
        Me.Current_User.Text = "Unknown User"
        Me.ToolTip1.SetToolTip(Me.Current_User, "User currently logged in")
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.DimGray
        Me.Label8.Location = New System.Drawing.Point(456, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 16)
        Me.Label8.TabIndex = 71
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.DimGray
        Me.Label9.Location = New System.Drawing.Point(584, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 16)
        Me.Label9.TabIndex = 72
        Me.Label9.Text = "BUILD 20060823.4"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label9, "Screen Build Number")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Red
        Me.Button2.Enabled = False
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.ForeColor = System.Drawing.Color.Red
        Me.Button2.Location = New System.Drawing.Point(672, 560)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(16, 16)
        Me.Button2.TabIndex = 118
        Me.Button2.Text = "Test"
        Me.ToolTip1.SetToolTip(Me.Button2, "Test")
        Me.Button2.UseVisualStyleBackColor = False
        Me.Button2.Visible = False
        '
        'btnProceed
        '
        Me.btnProceed.BackColor = System.Drawing.Color.Orange
        Me.btnProceed.Enabled = False
        Me.btnProceed.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnProceed.ForeColor = System.Drawing.Color.Black
        Me.btnProceed.Location = New System.Drawing.Point(24, 16)
        Me.btnProceed.Name = "btnProceed"
        Me.btnProceed.Size = New System.Drawing.Size(120, 23)
        Me.btnProceed.TabIndex = 119
        Me.btnProceed.Text = "Proceed"
        Me.ToolTip1.SetToolTip(Me.btnProceed, "Launch the hand-in process")
        Me.btnProceed.UseVisualStyleBackColor = False
        '
        'RegisteredCourses
        '
        Me.RegisteredCourses.BackColor = System.Drawing.Color.SteelBlue
        Me.RegisteredCourses.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RegisteredCourses.ForeColor = System.Drawing.Color.White
        Me.RegisteredCourses.Location = New System.Drawing.Point(568, 152)
        Me.RegisteredCourses.Name = "RegisteredCourses"
        Me.RegisteredCourses.Size = New System.Drawing.Size(112, 286)
        Me.RegisteredCourses.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.RegisteredCourses, "List of courses you are currently registered for")
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.CommunicationLabel2)
        Me.Panel1.Controls.Add(Me.RegisteredCourses)
        Me.Panel1.Controls.Add(Me.CommunicationLabel1)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(0, 32)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(968, 520)
        Me.Panel1.TabIndex = 68
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblProceed)
        Me.GroupBox3.Controls.Add(Me.btnProceed)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox3.ForeColor = System.Drawing.Color.PowderBlue
        Me.GroupBox3.Location = New System.Drawing.Point(16, 456)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(672, 48)
        Me.GroupBox3.TabIndex = 123
        Me.GroupBox3.TabStop = False
        '
        'lblProceed
        '
        Me.lblProceed.ForeColor = System.Drawing.Color.LightBlue
        Me.lblProceed.Location = New System.Drawing.Point(160, 16)
        Me.lblProceed.Name = "lblProceed"
        Me.lblProceed.Size = New System.Drawing.Size(504, 24)
        Me.lblProceed.TabIndex = 120
        Me.lblProceed.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CommunicationLabel2
        '
        Me.CommunicationLabel2.ForeColor = System.Drawing.Color.AliceBlue
        Me.CommunicationLabel2.Location = New System.Drawing.Point(16, 80)
        Me.CommunicationLabel2.Name = "CommunicationLabel2"
        Me.CommunicationLabel2.Size = New System.Drawing.Size(680, 40)
        Me.CommunicationLabel2.TabIndex = 120
        Me.CommunicationLabel2.Text = resources.GetString("CommunicationLabel2.Text")
        '
        'CommunicationLabel1
        '
        Me.CommunicationLabel1.ForeColor = System.Drawing.Color.AliceBlue
        Me.CommunicationLabel1.Location = New System.Drawing.Point(16, 16)
        Me.CommunicationLabel1.Name = "CommunicationLabel1"
        Me.CommunicationLabel1.Size = New System.Drawing.Size(672, 64)
        Me.CommunicationLabel1.TabIndex = 6
        Me.CommunicationLabel1.Text = resources.GetString("CommunicationLabel1.Text")
        '
        'GroupBox1
        '
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.ForeColor = System.Drawing.Color.PowderBlue
        Me.GroupBox1.Location = New System.Drawing.Point(552, 128)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(136, 320)
        Me.GroupBox1.TabIndex = 121
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Your Courses"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lstAssignmentsCover)
        Me.GroupBox2.Controls.Add(Me.lstCoursesCover)
        Me.GroupBox2.Controls.Add(Me.statsDepartmentTotal)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.lstAssignments)
        Me.GroupBox2.Controls.Add(Me.lstDepartmentCover)
        Me.GroupBox2.Controls.Add(Me.lstDepartment)
        Me.GroupBox2.Controls.Add(Me.statsDepartment)
        Me.GroupBox2.Controls.Add(Me.lstCourses)
        Me.GroupBox2.Controls.Add(Me.statsCourses)
        Me.GroupBox2.Controls.Add(Me.statsCoursesTotal)
        Me.GroupBox2.Controls.Add(Me.statsAssignments)
        Me.GroupBox2.Controls.Add(Me.statsAssignmentsTotal)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.ForeColor = System.Drawing.Color.PowderBlue
        Me.GroupBox2.Location = New System.Drawing.Point(16, 128)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(520, 320)
        Me.GroupBox2.TabIndex = 122
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Available Hand-in Folders"
        '
        'lstAssignmentsCover
        '
        Me.lstAssignmentsCover.BackColor = System.Drawing.Color.AliceBlue
        Me.lstAssignmentsCover.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstAssignmentsCover.Location = New System.Drawing.Point(33, 241)
        Me.lstAssignmentsCover.Name = "lstAssignmentsCover"
        Me.lstAssignmentsCover.Size = New System.Drawing.Size(256, 19)
        Me.lstAssignmentsCover.TabIndex = 15
        '
        'lstCoursesCover
        '
        Me.lstCoursesCover.BackColor = System.Drawing.Color.AliceBlue
        Me.lstCoursesCover.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstCoursesCover.Location = New System.Drawing.Point(33, 145)
        Me.lstCoursesCover.Name = "lstCoursesCover"
        Me.lstCoursesCover.Size = New System.Drawing.Size(256, 19)
        Me.lstCoursesCover.TabIndex = 10
        '
        'statsDepartmentTotal
        '
        Me.statsDepartmentTotal.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsDepartmentTotal.Location = New System.Drawing.Point(56, 96)
        Me.statsDepartmentTotal.Name = "statsDepartmentTotal"
        Me.statsDepartmentTotal.Size = New System.Drawing.Size(432, 16)
        Me.statsDepartmentTotal.TabIndex = 18
        '
        'Label4
        '
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(24, 216)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(216, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "3. Available Assignments:"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(24, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 16)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "2. Available Courses:"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "1. Available Departments:"
        '
        'lstAssignments
        '
        Me.lstAssignments.BackColor = System.Drawing.Color.AliceBlue
        Me.lstAssignments.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstAssignments.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstAssignments.ColumnNum = 1
        Me.lstAssignments.ColumnWidth = "121"
        Me.lstAssignments.DisplayMember = "Text"
        Me.lstAssignments.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstAssignments.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstAssignments.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstAssignments.DropDownForeColor = System.Drawing.Color.Black
        Me.lstAssignments.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstAssignments.DropDownWidth = 141
        Me.lstAssignments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstAssignments.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstAssignments.GridLineColor = System.Drawing.Color.Transparent
        Me.lstAssignments.GridLineHorizontal = False
        Me.lstAssignments.GridLineVertical = False
        Me.lstAssignments.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstAssignments.HighlightBorderOnMouseEvents = True
        Me.lstAssignments.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstAssignments.Location = New System.Drawing.Point(32, 240)
        Me.lstAssignments.ManagingFastMouseMoving = True
        Me.lstAssignments.ManagingFastMouseMovingInterval = 30
        Me.lstAssignments.Name = "lstAssignments"
        Me.lstAssignments.NormalBorderColor = System.Drawing.Color.Black
        Me.lstAssignments.Size = New System.Drawing.Size(280, 21)
        Me.lstAssignments.TabIndex = 14
        Me.lstAssignments.TabStop = False
        '
        'lstDepartmentCover
        '
        Me.lstDepartmentCover.BackColor = System.Drawing.Color.AliceBlue
        Me.lstDepartmentCover.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstDepartmentCover.Location = New System.Drawing.Point(33, 49)
        Me.lstDepartmentCover.Name = "lstDepartmentCover"
        Me.lstDepartmentCover.Size = New System.Drawing.Size(256, 19)
        Me.lstDepartmentCover.TabIndex = 7
        '
        'lstDepartment
        '
        Me.lstDepartment.BackColor = System.Drawing.Color.AliceBlue
        Me.lstDepartment.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstDepartment.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstDepartment.ColumnNum = 1
        Me.lstDepartment.ColumnWidth = "121"
        Me.lstDepartment.DisplayMember = "Text"
        Me.lstDepartment.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstDepartment.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstDepartment.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstDepartment.DropDownForeColor = System.Drawing.Color.Black
        Me.lstDepartment.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstDepartment.DropDownWidth = 141
        Me.lstDepartment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstDepartment.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstDepartment.GridLineColor = System.Drawing.Color.Transparent
        Me.lstDepartment.GridLineHorizontal = False
        Me.lstDepartment.GridLineVertical = False
        Me.lstDepartment.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstDepartment.HighlightBorderOnMouseEvents = True
        Me.lstDepartment.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstDepartment.Location = New System.Drawing.Point(32, 48)
        Me.lstDepartment.ManagingFastMouseMoving = True
        Me.lstDepartment.ManagingFastMouseMovingInterval = 30
        Me.lstDepartment.Name = "lstDepartment"
        Me.lstDepartment.NormalBorderColor = System.Drawing.Color.Black
        Me.lstDepartment.Size = New System.Drawing.Size(280, 21)
        Me.lstDepartment.TabIndex = 5
        Me.lstDepartment.TabStop = False
        '
        'statsDepartment
        '
        Me.statsDepartment.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsDepartment.Location = New System.Drawing.Point(56, 80)
        Me.statsDepartment.Name = "statsDepartment"
        Me.statsDepartment.Size = New System.Drawing.Size(432, 16)
        Me.statsDepartment.TabIndex = 8
        '
        'lstCourses
        '
        Me.lstCourses.BackColor = System.Drawing.Color.AliceBlue
        Me.lstCourses.BorderStyle = MTGCComboBox.TipiBordi.FlatXP
        Me.lstCourses.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.lstCourses.ColumnNum = 1
        Me.lstCourses.ColumnWidth = "121"
        Me.lstCourses.DisplayMember = "Text"
        Me.lstCourses.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstCourses.DropDownArrowBackColor = System.Drawing.Color.FromArgb(CType(CType(136, Byte), Integer), CType(CType(169, Byte), Integer), CType(CType(223, Byte), Integer))
        Me.lstCourses.DropDownBackColor = System.Drawing.Color.FromArgb(CType(CType(193, Byte), Integer), CType(CType(210, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.lstCourses.DropDownForeColor = System.Drawing.Color.Black
        Me.lstCourses.DropDownStyle = MTGCComboBox.CustomDropDownStyle.DropDown
        Me.lstCourses.DropDownWidth = 141
        Me.lstCourses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstCourses.ForeColor = System.Drawing.Color.SteelBlue
        Me.lstCourses.GridLineColor = System.Drawing.Color.Transparent
        Me.lstCourses.GridLineHorizontal = False
        Me.lstCourses.GridLineVertical = False
        Me.lstCourses.HighlightBorderColor = System.Drawing.Color.Black
        Me.lstCourses.HighlightBorderOnMouseEvents = True
        Me.lstCourses.LoadingType = MTGCComboBox.CaricamentoCombo.ComboBoxItem
        Me.lstCourses.Location = New System.Drawing.Point(32, 144)
        Me.lstCourses.ManagingFastMouseMoving = True
        Me.lstCourses.ManagingFastMouseMovingInterval = 30
        Me.lstCourses.Name = "lstCourses"
        Me.lstCourses.NormalBorderColor = System.Drawing.Color.Black
        Me.lstCourses.Size = New System.Drawing.Size(280, 21)
        Me.lstCourses.TabIndex = 9
        Me.lstCourses.TabStop = False
        '
        'statsCourses
        '
        Me.statsCourses.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsCourses.Location = New System.Drawing.Point(56, 176)
        Me.statsCourses.Name = "statsCourses"
        Me.statsCourses.Size = New System.Drawing.Size(432, 16)
        Me.statsCourses.TabIndex = 11
        '
        'statsCoursesTotal
        '
        Me.statsCoursesTotal.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsCoursesTotal.Location = New System.Drawing.Point(56, 192)
        Me.statsCoursesTotal.Name = "statsCoursesTotal"
        Me.statsCoursesTotal.Size = New System.Drawing.Size(432, 16)
        Me.statsCoursesTotal.TabIndex = 13
        '
        'statsAssignments
        '
        Me.statsAssignments.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsAssignments.Location = New System.Drawing.Point(48, 272)
        Me.statsAssignments.Name = "statsAssignments"
        Me.statsAssignments.Size = New System.Drawing.Size(440, 16)
        Me.statsAssignments.TabIndex = 16
        '
        'statsAssignmentsTotal
        '
        Me.statsAssignmentsTotal.ForeColor = System.Drawing.Color.LightSteelBlue
        Me.statsAssignmentsTotal.Location = New System.Drawing.Point(48, 288)
        Me.statsAssignmentsTotal.Name = "statsAssignmentsTotal"
        Me.statsAssignmentsTotal.Size = New System.Drawing.Size(440, 16)
        Me.statsAssignmentsTotal.TabIndex = 17
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 70
        Me.Label2.Text = "Current User:"
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Current_Context
        '
        Me.Current_Context.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Current_Context.Location = New System.Drawing.Point(325, 9)
        Me.Current_Context.Name = "Current_Context"
        Me.Current_Context.Size = New System.Drawing.Size(192, 16)
        Me.Current_Context.TabIndex = 119
        Me.ToolTip1.SetToolTip(Me.Current_Context, "User currently logged in")
        Me.Current_Context.Visible = False
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(706, 584)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Current_Context)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Current_User)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button2)
        Me.ForeColor = System.Drawing.Color.SteelBlue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(712, 616)
        Me.MinimumSize = New System.Drawing.Size(712, 616)
        Me.Name = "Main_Screen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Student Handin System 1.0"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                If identifier_msg = "" Then
                    Dim Display_Message1 As New Display_Message("The following application error has occurred: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString() & vbCrLf & "The application will attempt to recover from this shortly.")
                    Display_Message1.ShowDialog()
                Else
                    Dim Display_Message1 As New Display_Message("The following application error has occurred: " & vbCrLf & ex.Message.ToString() & vbCrLf & "The application will attempt to recover from this shortly.")
                    Display_Message1.ShowDialog()
                End If
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                If identifier_msg = "" Then
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & ex.ToString)
                Else
                    filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                End If

                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Student Handin System's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            application_exit = False
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            Timer2.Start()
            If (Not Worker1.ReturnRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory") = "") And (Worker1.ReturnRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory") Is Nothing = False) And (Not Worker1.ReturnRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory").StartsWith("Fail") = True) Then
                RootDirectory = Worker1.ReturnRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory")
            Else
                Worker1.CreateSubRegKey("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory", "")
                Worker1.SetRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory", "\\Comlab\Vol2\handin")
                RootDirectory = Worker1.ReturnRegKeyValue("HKEY_LOCAL_MACHINE", "SOFTWARE\Student Handin System", "RootDirectory")
            End If
            If RootDirectory.StartsWith("Fail") = True Then
                Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located and as such this application will now shut down.")
                Display_Message1.ShowDialog()
                application_exit = True
            End If
            InitialRootDirectory = RootDirectory
            If RootDirectory.StartsWith("\\") Then
                RootDirectory = Worker1.MapDrive(RootDirectory)
            End If
            If Not (RootDirectory).IndexOf("Fail") = -1 Then
                Dim Display_Message1 As New Display_Message("An usable Handin directory cannot be located and as such this application will now shut down.")
                Display_Message1.ShowDialog()
                application_exit = True
            End If
            If Current_User.Text.StartsWith("Unknown") = False Then
                Dim cresult As String = Worker1.UserGroups(Current_User.Text, Current_Context.Text)
                If cresult.Length > 0 Then
                    'MsgBox(cresult)
                    'RegisteredCourses.Text = cresult.Replace("cn=", "")
                    RegisteredCourses.Items.Clear()
                    Dim str As String() = cresult.Split(vbCrLf)
                    Dim playaround As String = ""
                    For Each read As String In str

                        If read.Trim.ToLower.StartsWith("result:") = False Then
                            playaround = read.ToLower().Replace("cn=", "").Split(",ou=")(0).Trim
                            If playaround.EndsWith("_g") Then
                                playaround = playaround.Remove(playaround.Length - 2, 2)
                            End If
                            playaround = playaround.ToUpper

                            If playaround.Length = 7 Or playaround.Length = 8 Then
                                If IsNumeric(playaround.Chars(3)) = True And IsNumeric(playaround.Chars(5)) = True And IsNumeric(playaround.Chars(playaround.Length - 1)) = False Then

                                    RegisteredCourses.Items.Add(playaround)


                                End If
                            End If


                        End If

                    Next
                End If


            Else
                Dim Display_Message1 As New Display_Message("Unknown users cannot be handled and as such this application will now shut down.")
                Display_Message1.ShowDialog()
                application_exit = True
            End If
            Dim dinfo As DirectoryInfo = New DirectoryInfo(RootDirectory)
            Dim yearfound As Boolean = False
            Dim dinfo2 As DirectoryInfo
            For Each dinfo2 In dinfo.GetDirectories
                If dinfo2.Name = Format(Now(), "yyyy") Then
                    yearfound = True
                    Exit For
                End If
            Next
            If yearfound = False Then
                CommunicationLabel1.Text = "Sorry, but no valid project hand-in folders have been created for " & Format(Now(), "yyyy") & ". "
                CommunicationLabel2.Text = ""
                lstDepartment.Enabled = False
                statsDepartment.Text = ""
                lstCourses.Enabled = False
                statsCourses.Text = ""
                statsCoursesTotal.Text = ""
                lstAssignments.Enabled = False
                statsAssignments.Text = ""
                statsAssignmentsTotal.Text = ""
            Else
                lstDepartment.Enabled = True
                dinfo = New DirectoryInfo((RootDirectory & "\").Replace("\\", "\") & Format(Now(), "yyyy"))
                lstDepartmentCover.Text = ""
                lstDepartment.Items.Clear()
                lstDepartment.Text = ""
                For Each dinfo2 In dinfo.GetDirectories
                    lstDepartment.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                Next
            End If
            If lstDepartment.Items.Count > 0 Then
                lstDepartment.Enabled = True
                lstDepartment.SelectedIndex = 0
            Else
                lstDepartment.Enabled = False
                statsDepartment.Text = ""
                lstCourses.Enabled = False
                statsCourses.Text = ""
                statsCoursesTotal.Text = ""
                lstAssignments.Enabled = False
                statsAssignments.Text = ""
                statsAssignmentsTotal.Text = ""
            End If
            If lstDepartment.Items.Count = 1 Then
                statsDepartment.Text = lstDepartment.Items.Count & " Department available to choose from."
            Else
                statsDepartment.Text = lstDepartment.Items.Count & " Departments available to choose from."
            End If
            dinfo = Nothing
            dinfo2 = Nothing

            dataloaded = True
            splash_loader.Visible = False

        Catch ex As Exception
            Error_Handler(ex, "Main_Screen_Load")
        End Try
    End Sub
    Private Sub WorkerErrorEncounteredHandler(ByVal ex As Exception, ByVal message As String)
        Try
            Error_Handler(ex, message)
        Catch exc As Exception
            Error_Handler(exc, "WorkerErrorEncounteredHandler")
        End Try
    End Sub
    Private Sub WorkerCompleteHandler(ByVal Result As String, ByVal Thread As Integer)
        Try
        Catch ex As Exception
            Error_Handler(ex, "WorkerCompleteHandler")
        End Try
    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            If lstDepartment.Items.Count > 0 And lstCourses.Items.Count > 0 And lstAssignments.Items.Count > 0 Then
                btnProceed.BackColor = Color.LimeGreen
                btnProceed.Enabled = True
            Else
                btnProceed.BackColor = Color.Orange
                btnProceed.Enabled = False
            End If
            If application_exit = True Then
                Me.Close()
            End If
        Catch ex As Exception
            Error_Handler(ex, "System Timer")
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Me.WindowState = FormWindowState.Minimized
            Me.Visible = False
            shutting_down = True
            Timer2.Stop()
            If RootDirectory.IndexOf("Fail") = -1 Then


                If InitialRootDirectory.StartsWith("\\") Then
                    RootDirectory = Worker1.UnMapDrive(RootDirectory)
                End If
            End If
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex, "Shutting Down Application")
        Finally
            Application.Exit()
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex, "Main_Screen_closed")
        End Try
    End Sub

    Private Sub lstDepartment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstDepartment.SelectedIndexChanged
        Try
            lstDepartmentCover.Text = lstDepartment.SelectedItem.text
            lblProceed.Text = lstDepartmentCover.Text
            LoadCourseFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartmentCover.Text).Replace("\\", "\"))
        Catch ex As Exception
            Error_Handler(ex, "lstDepartment_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub lstCourses_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstCourses.SelectedIndexChanged
        Try
            lstCoursesCover.Text = lstCourses.SelectedItem.text
            lblProceed.Text = lstDepartmentCover.Text & "\" & lstCoursesCover.Text
            LoadAssignmentFolders((RootDirectory & "\" & Format(Now, "yyyy") & "\" & lstDepartmentCover.Text & "\" & lstCoursesCover.Text).Replace("\\", "\"))
        Catch ex As Exception
            Error_Handler(ex, "lstCourses_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub LoadCourseFolders(ByVal targetdir As String)
        Try
            Dim dinfo2 As DirectoryInfo
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)
            Dim missedcourses As Integer = 0
            If dinfo.Exists = True Then
                lstCourses.Items.Clear()
                lstCoursesCover.Text = ""
                lstCourses.Text = ""
                lstAssignments.Items.Clear()
                lstAssignmentsCover.Text = ""
                lstAssignments.Text = ""
                For Each dinfo2 In dinfo.GetDirectories
                    If RegisteredCourses.Items.IndexOf(dinfo2.Name) > -1 Then
                        lstCourses.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                    Else
                        missedcourses = missedcourses + 1
                    End If


                Next

            End If

            If lstCourses.Items.Count > 0 Then
                lstCourses.Enabled = True
                lstCourses.SelectedIndex = 0
            
            If lstCourses.Items.Count = 1 Then
                statsCourses.Text = lstCourses.Items.Count & " Course available to choose from under " & lstDepartmentCover.Text & "."
            Else
                statsCourses.Text = lstCourses.Items.Count & " Courses available to choose from under " & lstDepartmentCover.Text & "."
            End If
            If missedcourses = 1 Then
                statsCoursesTotal.Text = missedcourses & " Course not eligible to choose from under " & lstDepartmentCover.Text & "."
            Else
                statsCoursesTotal.Text = missedcourses & " Courses not eligible to choose from under " & lstDepartmentCover.Text & "."
                End If
            Else

                lstCourses.Enabled = False
                statsCourses.Text = ""
                statsCoursesTotal.Text = ""
                lstAssignments.Enabled = False
                statsAssignments.Text = ""
                statsAssignmentsTotal.Text = ""
            End If
            dinfo2 = Nothing
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "LoadCourseFolders")
        End Try
    End Sub

   
    Private Sub lstAssignments_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstAssignments.SelectedIndexChanged
        Try
            lstAssignmentsCover.Text = lstAssignments.SelectedItem.text
            lblProceed.Text = lstDepartmentCover.Text & "\" & lstCoursesCover.Text & "\" & lstAssignmentsCover.Text
        Catch ex As Exception
            Error_Handler(ex, "lstAssignments_SelectedIndexChanged")
        End Try
    End Sub

    Private Sub LoadAssignmentFolders(ByVal targetdir As String)
        Try
            Dim dinfo2 As DirectoryInfo
            Dim dinfo As DirectoryInfo = New DirectoryInfo(targetdir)
            Dim missedcourses As Integer = 0
            If dinfo.Exists = True Then
                lstAssignments.Items.Clear()
                lstAssignmentsCover.Text = ""
                lstAssignments.Text = ""
                For Each dinfo2 In dinfo.GetDirectories
                    Dim finfo As FileInfo = New FileInfo((dinfo2.FullName & "\access.ini").Replace("\\", "\"))
                    Dim additem As Boolean = False
                    If finfo.Exists = False Then
                        additem = True
                    Else
                        Dim filereader As StreamReader = New StreamReader(finfo.FullName)
                        Dim lineread As String = ""
                        Dim opentime As String = "open:000000000000"
                        Dim closetime As String = "close:000000000000"
                        While Not filereader.Peek = -1
                            lineread = filereader.ReadLine
                            If lineread.ToLower.StartsWith("open:") Then
                                opentime = lineread.Replace("open:", "")
                            End If
                            If lineread.ToLower.StartsWith("close:") Then
                                closetime = lineread.Replace("close:", "")
                            End If
                        End While
                        filereader.Close()
                        filereader = Nothing
                        Dim opendate, closedate As Date
                        If opentime.Length = 12 And closetime.Length = 12 Then
                            opendate = New Date(CInt(opentime.Substring(0, 4)), CInt(opentime.Substring(4, 2)), CInt(opentime.Substring(6, 2)), CInt(opentime.Substring(8, 2)), CInt(opentime.Substring(10, 2)), 0, 0)
                            closedate = New Date(CInt(closetime.Substring(0, 4)), CInt(closetime.Substring(4, 2)), CInt(closetime.Substring(6, 2)), CInt(closetime.Substring(8, 2)), CInt(closetime.Substring(10, 2)), 0, 0)
                            If Now >= opendate And Now <= closedate Then
                                additem = True
                            End If
                        End If
                    End If
                    If additem = True Then
                        lstAssignments.Items.Add(New MTGCComboBoxItem(dinfo2.Name))
                    Else
                        missedcourses = missedcourses + 1
                    End If
                Next

            End If

            If lstAssignments.Items.Count > 0 Then
                lstAssignments.Enabled = True
                lstAssignments.SelectedIndex = 0
            
            If lstAssignments.Items.Count = 1 Then
                statsAssignments.Text = lstAssignments.Items.Count & " Assignment available to choose from under " & lstCoursesCover.Text & "."
            Else
                statsAssignments.Text = lstAssignments.Items.Count & " Assignments available to choose from under " & lstCoursesCover.Text & "."
            End If
            If missedcourses = 1 Then
                statsAssignmentsTotal.Text = missedcourses & " Assignment not eligible to choose from under " & lstCoursesCover.Text & "."
            Else
                statsAssignmentsTotal.Text = missedcourses & " Assignments not eligible to choose from under " & lstCoursesCover.Text & "."
                End If
            Else
                lstAssignments.Enabled = False
                statsAssignments.Text = ""
                statsAssignmentsTotal.Text = ""
            End If
            dinfo2 = Nothing
            dinfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "LoadCourseFolders")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RegisteredCourses.Items.Add("INF204F")
        RegisteredCourses.Items.Add("INF1002F")
    End Sub

    Private Sub btnProceed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProceed.Click
        Try

        
        lblProceed.Text = lstDepartmentCover.Text & "\" & lstCoursesCover.Text & "\" & lstAssignmentsCover.Text & "\" & Current_User.Text

        Dim finfo As FileInfo = New FileInfo((Application.StartupPath & "\File System Controls.exe").Replace("\\", "\"))
        Dim apptorun As String = ""
        If finfo.Exists = True Then
            apptorun = """" & (Application.StartupPath & "\File System Controls.exe").Replace("\\", "\") & """ """ & lstCoursesCover.Text & "\" & lstAssignmentsCover.Text & """ """ & (RootDirectory & "\").Replace("\\", "\") & Format(Now(), "yyyy") & "\" & lblProceed.Text & """"
            ApplicationLauncher(apptorun)

        End If
        finfo = Nothing
        Me.WindowState = FormWindowState.Minimized
            ' Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "btnProceed_Click")
        End Try
    End Sub

    Private Function ApplicationLauncher(ByVal apptoRun As String) As String
        Dim sresult As String = ""
        Try
            Dim myProcess As Process = New Process

            Dim executable, arguments As String
            Dim str As String()
            executable = ""
            arguments = ""
            If apptoRun.StartsWith("""") = True Then
                Dim endpos As Integer = apptoRun.IndexOf("""", apptoRun.IndexOf("""", 0) + 1)
                executable = apptoRun.Substring(0, endpos + 1)
                If apptoRun.Length >= (endpos + 3) Then
                    arguments = apptoRun.Substring(endpos + 2)
                End If
            Else
                str = apptoRun.Split(" ")
                For i As Integer = 0 To str.Length - 1
                    If i = 0 Then
                        executable = str(i)
                    Else
                        arguments = arguments & str(i) & " "
                    End If
                Next
                arguments = arguments.Remove(arguments.Length - 1, 1)
            End If
            'Activity_Logger("LAUNCH ATTEMPT INITIATED")

            myProcess.StartInfo.FileName = executable.Replace("""", "")
            myProcess.StartInfo.Arguments = arguments
            'Activity_Logger("Executable: " & myProcess.StartInfo.FileName)
            'Activity_Logger("Arguments: " & myProcess.StartInfo.Arguments)
            myProcess.StartInfo.UseShellExecute = True

            myProcess.StartInfo.CreateNoWindow = False
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal

            myProcess.StartInfo.RedirectStandardInput = False
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.RedirectStandardError = False
            myProcess.Start()
            sresult = "Success"
            Return sresult

        Catch ex As Exception
            Error_Handler(ex, "Executing: " & apptoRun)
            sresult = "Fail"
        End Try
        Return sresult
    End Function
End Class
