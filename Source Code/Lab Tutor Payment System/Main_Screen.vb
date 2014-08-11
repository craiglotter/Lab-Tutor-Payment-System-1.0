Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Xml

Public Class Main_Screen
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
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Reset_Column_Widths As System.Windows.Forms.Button
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Student_Number As System.Windows.Forms.ColumnHeader
    Friend WithEvents Last_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents First_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents Staff_Number As System.Windows.Forms.ColumnHeader
    Friend WithEvents Add_Student_Button As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents DateWorked As System.Windows.Forms.ColumnHeader
    Friend WithEvents StudentNumber As System.Windows.Forms.ColumnHeader
    Friend WithEvents StartTime As System.Windows.Forms.ColumnHeader
    Friend WithEvents EndTime As System.Windows.Forms.ColumnHeader
    Friend WithEvents Lab As System.Windows.Forms.ColumnHeader
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ContextMenu3 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog2 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog3 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog3 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Payment_Period_End As System.Windows.Forms.DateTimePicker
    Friend WithEvents Payment_Period_Start As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Department As System.Windows.Forms.TextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Cost_Centre As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Fund As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Hourly_Rate_Cents As System.Windows.Forms.NumericUpDown
    Friend WithEvents Hourly_Rate_Rands As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog4 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog4 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Create_Subfolder As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Progress_Bar1_Panel As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Report_Compiler As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem28 As System.Windows.Forms.MenuItem
    Friend WithEvents Overwrite_Files As System.Windows.Forms.CheckBox
    Friend WithEvents Show_Created_Files As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Add_Student_Button = New System.Windows.Forms.Button
        Me.Reset_Column_Widths = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.ContextMenu3 = New System.Windows.Forms.ContextMenu
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.Student_Number = New System.Windows.Forms.ColumnHeader
        Me.Last_Name = New System.Windows.Forms.ColumnHeader
        Me.First_Name = New System.Windows.Forms.ColumnHeader
        Me.Staff_Number = New System.Windows.Forms.ColumnHeader
        Me.ListView2 = New System.Windows.Forms.ListView
        Me.DateWorked = New System.Windows.Forms.ColumnHeader
        Me.StudentNumber = New System.Windows.Forms.ColumnHeader
        Me.StartTime = New System.Windows.Forms.ColumnHeader
        Me.EndTime = New System.Windows.Forms.ColumnHeader
        Me.Lab = New System.Windows.Forms.ColumnHeader
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.Button5 = New System.Windows.Forms.Button
        Me.Payment_Period_Start = New System.Windows.Forms.DateTimePicker
        Me.Payment_Period_End = New System.Windows.Forms.DateTimePicker
        Me.Department = New System.Windows.Forms.TextBox
        Me.Cost_Centre = New System.Windows.Forms.TextBox
        Me.Fund = New System.Windows.Forms.TextBox
        Me.Hourly_Rate_Cents = New System.Windows.Forms.NumericUpDown
        Me.Hourly_Rate_Rands = New System.Windows.Forms.NumericUpDown
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Report_Compiler = New System.Windows.Forms.TextBox
        Me.Create_Subfolder = New System.Windows.Forms.CheckBox
        Me.Overwrite_Files = New System.Windows.Forms.CheckBox
        Me.Show_Created_Files = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Label14 = New System.Windows.Forms.Label
        Me.Progress_Bar1_Panel = New System.Windows.Forms.Panel
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label13 = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label8 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label7 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem26 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.MenuItem28 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog2 = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog
        Me.OpenFileDialog3 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog3 = New System.Windows.Forms.SaveFileDialog
        Me.SaveFileDialog4 = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog4 = New System.Windows.Forms.OpenFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        CType(Me.Hourly_Rate_Cents, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Hourly_Rate_Rands, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Progress_Bar1_Panel.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Edit"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Delete"
        '
        'Add_Student_Button
        '
        Me.Add_Student_Button.BackColor = System.Drawing.Color.NavajoWhite
        Me.Add_Student_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Add_Student_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Add_Student_Button.Location = New System.Drawing.Point(504, 264)
        Me.Add_Student_Button.Name = "Add_Student_Button"
        Me.Add_Student_Button.Size = New System.Drawing.Size(80, 18)
        Me.Add_Student_Button.TabIndex = 10
        Me.Add_Student_Button.Text = "Add Student"
        Me.ToolTip1.SetToolTip(Me.Add_Student_Button, "Add a Student to the Student List")
        '
        'Reset_Column_Widths
        '
        Me.Reset_Column_Widths.BackColor = System.Drawing.Color.NavajoWhite
        Me.Reset_Column_Widths.Location = New System.Drawing.Point(496, 72)
        Me.Reset_Column_Widths.Name = "Reset_Column_Widths"
        Me.Reset_Column_Widths.Size = New System.Drawing.Size(8, 8)
        Me.Reset_Column_Widths.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.Reset_Column_Widths, "Reset Student List Column Widths")
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button1.Location = New System.Drawing.Point(496, 96)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(8, 8)
        Me.Button1.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.Button1, "Clear List")
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(512, 264)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(80, 18)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "Add Event"
        Me.ToolTip1.SetToolTip(Me.Button3, "Add a Timesheet Event")
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(508, 128)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(8, 8)
        Me.Button4.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.Button4, "Reset Event List Column Widths")
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(508, 152)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(8, 8)
        Me.Button2.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.Button2, "Clear List")
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.SeaShell
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.ContextMenu = Me.ContextMenu3
        Me.ListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.ForeColor = System.Drawing.Color.Black
        Me.ListBox1.Location = New System.Drawing.Point(16, 50)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(144, 28)
        Me.ListBox1.TabIndex = 23
        Me.ToolTip1.SetToolTip(Me.ListBox1, "The Computer Labs used in the Timesheet Events")
        '
        'ContextMenu3
        '
        Me.ContextMenu3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem13, Me.MenuItem14})
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 0
        Me.MenuItem13.Text = "Edit"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 1
        Me.MenuItem14.Text = "Delete"
        '
        'ListView1
        '
        Me.ListView1.BackColor = System.Drawing.Color.SeaShell
        Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Student_Number, Me.Last_Name, Me.First_Name, Me.Staff_Number})
        Me.ListView1.ContextMenu = Me.ContextMenu1
        Me.ListView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.Location = New System.Drawing.Point(13, 64)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(480, 224)
        Me.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView1.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.ListView1, "The Students available for use in Timesheet Events.")
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'Student_Number
        '
        Me.Student_Number.Text = "Student Number"
        Me.Student_Number.Width = 91
        '
        'Last_Name
        '
        Me.Last_Name.Text = "Last Name"
        Me.Last_Name.Width = 130
        '
        'First_Name
        '
        Me.First_Name.Text = "First Name"
        Me.First_Name.Width = 130
        '
        'Staff_Number
        '
        Me.Staff_Number.Text = "Staff Number"
        Me.Staff_Number.Width = 105
        '
        'ListView2
        '
        Me.ListView2.BackColor = System.Drawing.Color.SeaShell
        Me.ListView2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListView2.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.DateWorked, Me.StudentNumber, Me.StartTime, Me.EndTime, Me.Lab})
        Me.ListView2.ContextMenu = Me.ContextMenu2
        Me.ListView2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView2.FullRowSelect = True
        Me.ListView2.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView2.Location = New System.Drawing.Point(16, 120)
        Me.ListView2.MultiSelect = False
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(488, 168)
        Me.ListView2.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView2.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.ListView2, "Timesheet Events")
        Me.ListView2.View = System.Windows.Forms.View.Details
        '
        'DateWorked
        '
        Me.DateWorked.Text = "Date"
        Me.DateWorked.Width = 75
        '
        'StudentNumber
        '
        Me.StudentNumber.Text = "Student Number"
        Me.StudentNumber.Width = 97
        '
        'StartTime
        '
        Me.StartTime.Text = "Start Time"
        Me.StartTime.Width = 90
        '
        'EndTime
        '
        Me.EndTime.Text = "End Time"
        Me.EndTime.Width = 90
        '
        'Lab
        '
        Me.Lab.Text = "Computer Lab"
        Me.Lab.Width = 112
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem11, Me.MenuItem12})
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 0
        Me.MenuItem11.Text = "Edit"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 1
        Me.MenuItem12.Text = "Delete"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(168, 60)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(56, 18)
        Me.Button5.TabIndex = 26
        Me.Button5.Text = "Add Lab"
        Me.ToolTip1.SetToolTip(Me.Button5, "Add a Computer Lab")
        '
        'Payment_Period_Start
        '
        Me.Payment_Period_Start.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Payment_Period_Start.CalendarForeColor = System.Drawing.Color.Black
        Me.Payment_Period_Start.CalendarMonthBackground = System.Drawing.Color.SeaShell
        Me.Payment_Period_Start.CalendarTitleBackColor = System.Drawing.Color.NavajoWhite
        Me.Payment_Period_Start.CalendarTitleForeColor = System.Drawing.Color.Black
        Me.Payment_Period_Start.CalendarTrailingForeColor = System.Drawing.Color.LightGray
        Me.Payment_Period_Start.CustomFormat = "yy/mm/dd"
        Me.Payment_Period_Start.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Payment_Period_Start.Location = New System.Drawing.Point(6, 32)
        Me.Payment_Period_Start.Name = "Payment_Period_Start"
        Me.Payment_Period_Start.Size = New System.Drawing.Size(136, 20)
        Me.Payment_Period_Start.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.Payment_Period_Start, "Payment Period Start")
        '
        'Payment_Period_End
        '
        Me.Payment_Period_End.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Payment_Period_End.CalendarForeColor = System.Drawing.Color.Black
        Me.Payment_Period_End.CalendarMonthBackground = System.Drawing.Color.SeaShell
        Me.Payment_Period_End.CalendarTitleBackColor = System.Drawing.Color.NavajoWhite
        Me.Payment_Period_End.CalendarTitleForeColor = System.Drawing.Color.Black
        Me.Payment_Period_End.CalendarTrailingForeColor = System.Drawing.Color.LightGray
        Me.Payment_Period_End.CustomFormat = "yy/mm/dd"
        Me.Payment_Period_End.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Payment_Period_End.Location = New System.Drawing.Point(150, 32)
        Me.Payment_Period_End.Name = "Payment_Period_End"
        Me.Payment_Period_End.Size = New System.Drawing.Size(136, 20)
        Me.Payment_Period_End.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.Payment_Period_End, "Payment Period End")
        '
        'Department
        '
        Me.Department.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Department.Location = New System.Drawing.Point(8, 32)
        Me.Department.Name = "Department"
        Me.Department.Size = New System.Drawing.Size(128, 20)
        Me.Department.TabIndex = 6
        Me.Department.Text = ""
        Me.ToolTip1.SetToolTip(Me.Department, "Department Title")
        '
        'Cost_Centre
        '
        Me.Cost_Centre.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cost_Centre.Location = New System.Drawing.Point(8, 32)
        Me.Cost_Centre.Name = "Cost_Centre"
        Me.Cost_Centre.Size = New System.Drawing.Size(128, 20)
        Me.Cost_Centre.TabIndex = 6
        Me.Cost_Centre.Text = ""
        Me.ToolTip1.SetToolTip(Me.Cost_Centre, "Cost Centre ID")
        '
        'Fund
        '
        Me.Fund.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Fund.Location = New System.Drawing.Point(8, 32)
        Me.Fund.Name = "Fund"
        Me.Fund.Size = New System.Drawing.Size(128, 20)
        Me.Fund.TabIndex = 6
        Me.Fund.Text = ""
        Me.ToolTip1.SetToolTip(Me.Fund, "Fund ID")
        '
        'Hourly_Rate_Cents
        '
        Me.Hourly_Rate_Cents.BackColor = System.Drawing.Color.White
        Me.Hourly_Rate_Cents.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Hourly_Rate_Cents.ForeColor = System.Drawing.Color.Black
        Me.Hourly_Rate_Cents.Location = New System.Drawing.Point(72, 32)
        Me.Hourly_Rate_Cents.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.Hourly_Rate_Cents.Name = "Hourly_Rate_Cents"
        Me.Hourly_Rate_Cents.Size = New System.Drawing.Size(40, 20)
        Me.Hourly_Rate_Cents.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.Hourly_Rate_Cents, "Cents")
        '
        'Hourly_Rate_Rands
        '
        Me.Hourly_Rate_Rands.BackColor = System.Drawing.Color.White
        Me.Hourly_Rate_Rands.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Hourly_Rate_Rands.ForeColor = System.Drawing.Color.Black
        Me.Hourly_Rate_Rands.Location = New System.Drawing.Point(24, 32)
        Me.Hourly_Rate_Rands.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.Hourly_Rate_Rands.Name = "Hourly_Rate_Rands"
        Me.Hourly_Rate_Rands.Size = New System.Drawing.Size(40, 20)
        Me.Hourly_Rate_Rands.TabIndex = 11
        Me.Hourly_Rate_Rands.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.Hourly_Rate_Rands, "Rands")
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(320, 136)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(208, 18)
        Me.Button6.TabIndex = 22
        Me.Button6.Text = "Generate Student Timesheets"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Button6, "Generate Student Timesheets")
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.Location = New System.Drawing.Point(320, 160)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(208, 18)
        Me.Button7.TabIndex = 23
        Me.Button7.Text = "Generate HR106 Multiple Payments Sheets"
        Me.Button7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Button7, "Generate HR106 Multiple Payments Sheets")
        '
        'Report_Compiler
        '
        Me.Report_Compiler.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Report_Compiler.Location = New System.Drawing.Point(8, 32)
        Me.Report_Compiler.Name = "Report_Compiler"
        Me.Report_Compiler.Size = New System.Drawing.Size(128, 20)
        Me.Report_Compiler.TabIndex = 6
        Me.Report_Compiler.Text = ""
        Me.ToolTip1.SetToolTip(Me.Report_Compiler, "The person responsible for compiling the payment reports")
        '
        'Create_Subfolder
        '
        Me.Create_Subfolder.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Create_Subfolder.Checked = True
        Me.Create_Subfolder.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Create_Subfolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Create_Subfolder.Location = New System.Drawing.Point(320, 208)
        Me.Create_Subfolder.Name = "Create_Subfolder"
        Me.Create_Subfolder.Size = New System.Drawing.Size(216, 32)
        Me.Create_Subfolder.TabIndex = 24
        Me.Create_Subfolder.Text = "Put reports in a subfolder using today's date."
        Me.Create_Subfolder.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ToolTip1.SetToolTip(Me.Create_Subfolder, "Put reports in a subfolder using today's date.")
        '
        'Overwrite_Files
        '
        Me.Overwrite_Files.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Overwrite_Files.Checked = True
        Me.Overwrite_Files.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Overwrite_Files.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Overwrite_Files.Location = New System.Drawing.Point(320, 240)
        Me.Overwrite_Files.Name = "Overwrite_Files"
        Me.Overwrite_Files.Size = New System.Drawing.Size(216, 16)
        Me.Overwrite_Files.TabIndex = 27
        Me.Overwrite_Files.Text = "Overwrite existing files"
        Me.Overwrite_Files.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ToolTip1.SetToolTip(Me.Overwrite_Files, "Overwrite existing files")
        '
        'Show_Created_Files
        '
        Me.Show_Created_Files.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Show_Created_Files.Checked = True
        Me.Show_Created_Files.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Show_Created_Files.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Show_Created_Files.Location = New System.Drawing.Point(320, 264)
        Me.Show_Created_Files.Name = "Show_Created_Files"
        Me.Show_Created_Files.Size = New System.Drawing.Size(216, 16)
        Me.Show_Created_Files.TabIndex = 28
        Me.Show_Created_Files.Text = "Show created files"
        Me.Show_Created_Files.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ToolTip1.SetToolTip(Me.Show_Created_Files, "Show created files")
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.Padding = New System.Drawing.Point(0, 0)
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(680, 352)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage1.BackgroundImage = CType(resources.GetObject("TabPage1.BackgroundImage"), System.Drawing.Image)
        Me.TabPage1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.ForeColor = System.Drawing.Color.Black
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(672, 323)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Student List"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Add_Student_Button)
        Me.GroupBox1.Controls.Add(Me.Reset_Column_Widths)
        Me.GroupBox1.Controls.Add(Me.ListView1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(642, 302)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Student List"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.IndianRed
        Me.Label5.Location = New System.Drawing.Point(13, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(496, 32)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "The Student List controls which students are available for use on the timesheet a" & _
        "nd the generated reports. To add a student, click on the 'Add Student' button be" & _
        "low."
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage2.BackgroundImage = CType(resources.GetObject("TabPage2.BackgroundImage"), System.Drawing.Image)
        Me.TabPage2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.ForeColor = System.Drawing.Color.Black
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(672, 323)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Timesheet Events"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.Button5)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.ListBox1)
        Me.GroupBox2.Controls.Add(Me.ListView2)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(15, 10)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(642, 302)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Timesheet Events"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.IndianRed
        Me.Label2.Location = New System.Drawing.Point(16, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(336, 16)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "The Computer Labs available for use in Timesheet events."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.IndianRed
        Me.Label1.Location = New System.Drawing.Point(16, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(544, 16)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Reports are generated from Timesheet events. To add an event, click on the 'Add E" & _
        "vent' button below."
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TabPage3.BackgroundImage = CType(resources.GetObject("TabPage3.BackgroundImage"), System.Drawing.Image)
        Me.TabPage3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TabPage3.Controls.Add(Me.GroupBox3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(672, 323)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Report Data"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.Show_Created_Files)
        Me.GroupBox3.Controls.Add(Me.Overwrite_Files)
        Me.GroupBox3.Controls.Add(Me.Panel6)
        Me.GroupBox3.Controls.Add(Me.Progress_Bar1_Panel)
        Me.GroupBox3.Controls.Add(Me.Create_Subfolder)
        Me.GroupBox3.Controls.Add(Me.Button7)
        Me.GroupBox3.Controls.Add(Me.Button6)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.Panel5)
        Me.GroupBox3.Controls.Add(Me.Panel4)
        Me.GroupBox3.Controls.Add(Me.Panel3)
        Me.GroupBox3.Controls.Add(Me.Panel2)
        Me.GroupBox3.Controls.Add(Me.Panel1)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(15, 10)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(642, 302)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Report Data"
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.SeaShell
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.Report_Compiler)
        Me.Panel6.Controls.Add(Me.Label14)
        Me.Panel6.ForeColor = System.Drawing.Color.Black
        Me.Panel6.Location = New System.Drawing.Point(320, 56)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(144, 64)
        Me.Panel6.TabIndex = 12
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(6, 6)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(104, 16)
        Me.Label14.TabIndex = 5
        Me.Label14.Text = "Report Compiler"
        '
        'Progress_Bar1_Panel
        '
        Me.Progress_Bar1_Panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Progress_Bar1_Panel.Controls.Add(Me.ProgressBar1)
        Me.Progress_Bar1_Panel.Location = New System.Drawing.Point(320, 184)
        Me.Progress_Bar1_Panel.Name = "Progress_Bar1_Panel"
        Me.Progress_Bar1_Panel.Size = New System.Drawing.Size(208, 16)
        Me.Progress_Bar1_Panel.TabIndex = 26
        Me.Progress_Bar1_Panel.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Top
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(206, 16)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 25
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.IndianRed
        Me.Label13.Location = New System.Drawing.Point(16, 32)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(367, 17)
        Me.Label13.TabIndex = 21
        Me.Label13.Text = "The Report Data is essential for the generated report's completeness"
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.Color.SeaShell
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.Label12)
        Me.Panel5.Controls.Add(Me.Label11)
        Me.Panel5.Controls.Add(Me.Hourly_Rate_Cents)
        Me.Panel5.Controls.Add(Me.Hourly_Rate_Rands)
        Me.Panel5.Controls.Add(Me.Label10)
        Me.Panel5.ForeColor = System.Drawing.Color.Black
        Me.Panel5.Location = New System.Drawing.Point(168, 216)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(144, 64)
        Me.Panel5.TabIndex = 11
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 34)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(16, 16)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "R"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(64, 34)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(8, 16)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = ","
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(6, 6)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(104, 16)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Hourly Rate"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.SeaShell
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Fund)
        Me.Panel4.Controls.Add(Me.Label8)
        Me.Panel4.ForeColor = System.Drawing.Color.Black
        Me.Panel4.Location = New System.Drawing.Point(16, 216)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(144, 64)
        Me.Panel4.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(6, 6)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(104, 16)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Fund"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.SeaShell
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.Cost_Centre)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.ForeColor = System.Drawing.Color.Black
        Me.Panel3.Location = New System.Drawing.Point(168, 144)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(144, 64)
        Me.Panel3.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(6, 6)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 16)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "Cost Centre"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SeaShell
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.Department)
        Me.Panel2.Controls.Add(Me.Label9)
        Me.Panel2.ForeColor = System.Drawing.Color.Black
        Me.Panel2.Location = New System.Drawing.Point(16, 144)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(144, 64)
        Me.Panel2.TabIndex = 6
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(6, 6)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(104, 16)
        Me.Label9.TabIndex = 5
        Me.Label9.Text = "Department"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SeaShell
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Payment_Period_End)
        Me.Panel1.Controls.Add(Me.Payment_Period_Start)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.ForeColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(16, 56)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(296, 80)
        Me.Panel1.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(150, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(74, 16)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Period End"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Period Start"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(6, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Payment Period"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem4, Me.MenuItem5, Me.MenuItem6})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem18, Me.MenuItem19, Me.MenuItem20, Me.MenuItem24, Me.MenuItem8, Me.MenuItem25, Me.MenuItem26, Me.MenuItem7})
        Me.MenuItem3.Text = "File"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 0
        Me.MenuItem18.Text = "Save Job Control File"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 1
        Me.MenuItem19.Text = "Load Job Control File"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 2
        Me.MenuItem20.Text = "-"
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 3
        Me.MenuItem24.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem17, Me.MenuItem15, Me.MenuItem16, Me.MenuItem23, Me.MenuItem22, Me.MenuItem21})
        Me.MenuItem24.Text = "Save/Load Individual Files"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "Save Student List"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Load Student List"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 2
        Me.MenuItem17.Text = "-"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 3
        Me.MenuItem15.Text = "Save Timesheet Events List"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 4
        Me.MenuItem16.Text = "Load Timesheet Events List"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 5
        Me.MenuItem23.Text = "-"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 6
        Me.MenuItem22.Text = "Save Report Data"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 7
        Me.MenuItem21.Text = "Load Report Data"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 4
        Me.MenuItem8.Text = "-"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 5
        Me.MenuItem25.Text = "Clear Program Data"
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 6
        Me.MenuItem26.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 7
        Me.MenuItem7.Text = "Exit"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem27, Me.MenuItem28})
        Me.MenuItem4.Text = "Reports"
        '
        'MenuItem27
        '
        Me.MenuItem27.Index = 0
        Me.MenuItem27.Text = "Generate Student Timesheets"
        '
        'MenuItem28
        '
        Me.MenuItem28.Index = 1
        Me.MenuItem28.Text = "Generate HR106 Multiple Payments Sheets"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.Text = "Help"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "About"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "slt"
        Me.SaveFileDialog1.FileName = "students1.slt"
        Me.SaveFileDialog1.Filter = "Student List Files|*.slt|All files|*.*"
        Me.SaveFileDialog1.Title = "Save Student List"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "Student List Files|*.slt|All files|*.*"
        Me.OpenFileDialog1.Title = "Load Student List"
        '
        'SaveFileDialog2
        '
        Me.SaveFileDialog2.DefaultExt = "slt"
        Me.SaveFileDialog2.FileName = "timesheet1.tlt"
        Me.SaveFileDialog2.Filter = "Timesheet Events Files|*.tlt|All files|*.*"
        Me.SaveFileDialog2.Title = "Save Timesheet Events List"
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.Filter = "Timesheet Events Files|*.tlt|All files|*.*"
        Me.OpenFileDialog2.Title = "Load Timesheet Events List"
        '
        'OpenFileDialog3
        '
        Me.OpenFileDialog3.Filter = "Job Control Files|*.jlt|All files|*.*"
        Me.OpenFileDialog3.Title = "Load Job Control File"
        '
        'SaveFileDialog3
        '
        Me.SaveFileDialog3.DefaultExt = "slt"
        Me.SaveFileDialog3.FileName = "jobcontrol1.jlt"
        Me.SaveFileDialog3.Filter = "Job Control Files|*.jlt|All files|*.*"
        Me.SaveFileDialog3.Title = "Save Job Control File"
        '
        'SaveFileDialog4
        '
        Me.SaveFileDialog4.DefaultExt = "rlt"
        Me.SaveFileDialog4.FileName = "report1.rlt"
        Me.SaveFileDialog4.Filter = "Report Data Files|*.rlt|All files|*.*"
        Me.SaveFileDialog4.Title = "Save Report Data"
        '
        'OpenFileDialog4
        '
        Me.OpenFileDialog4.Filter = "Report Data Files|*.rlt|All files|*.*"
        Me.OpenFileDialog4.Title = "Load Report Data"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the folder in which the generated forms will be created."
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(680, 371)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(688, 405)
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(688, 405)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Lab Tutor Payment System"
        CType(Me.Hourly_Rate_Cents, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Hourly_Rate_Rands, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Progress_Bar1_Panel.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Structure Student_Details
        Dim Surname As String
        Dim Staff_Number As String
        Dim First_Name As String
        Dim Student_Number As String
    End Structure

    Structure Event_Details
        Dim Event_Date As Date
        Dim Event_Date_String As String
        Dim Student_Number As String
        Dim Start_Time As String
        Dim End_Time As String
        Dim Computer_Lab As String
    End Structure

    '  Dim Student_List As ArrayList

    Private Sub Error_Handler(ByVal error_message As String)
        Try
            MsgBox("Sorry, but it would appear that Lab Tutor Payment System has encountered an error situation. The following error has been encountered: " & vbCrLf & error_message, MsgBoxStyle.Exclamation, "Lab Tutor Payment System")
        Catch except As Exception
        End Try
    End Sub

    Private Sub Message(ByVal input_message As String)
        Try
            MsgBox(input_message, MsgBoxStyle.Information, "Lab Tutor Payment System")
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub






    Private Sub Reset_Column_Widths_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset_Column_Widths.Click
        Try
            ListView1.Columns.Item(0).Width = 91
            ListView1.Columns.Item(1).Width = 130
            ListView1.Columns.Item(2).Width = 130
            ListView1.Columns.Item(3).Width = 105
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try

    End Sub

    Private Sub Add_Student_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Add_Student_Button.Click
        Add_Student_Function()
    End Sub

    Private Sub Add_Student_Function()
        Try
            Dim add_student_dialog As New Student_Details_Display
            Dim result As DialogResult
            result = add_student_dialog.ShowDialog()
            If result = DialogResult.OK Then
                Dim add_student As New Student_Details
                add_student.First_Name = add_student_dialog.First_Name.Text.Trim
                add_student.Surname = add_student_dialog.Surname.Text.Trim
                add_student.Staff_Number = add_student_dialog.Staff_Number.Text.Trim
                add_student.Student_Number = add_student_dialog.Student_Number.Text.Trim

                Dim continue_flag As Boolean
                continue_flag = True
                If Not add_student.Student_Number.Trim().Equals("") Then
                    Dim i As Long
                    For i = 0 To ListView1.Items.Count - 1

                        If UCase(add_student.Student_Number) = UCase(ListView1.Items(i).Text) Then
                            Message(UCase(add_student.Student_Number) & " cannot be added to the Student list as he/she is already present in the list.")
                            continue_flag = False
                        End If
                    Next
                    If continue_flag = True Then
                        Dim additem As New ListViewItem
                        additem.Text = UCase(add_student.Student_Number.Trim())
                        additem.SubItems.Add(add_student.Surname.Trim())
                        additem.SubItems.Add(add_student.First_Name.Trim())
                        additem.SubItems.Add(UCase(add_student.Staff_Number).Trim())
                        ListView1.Items.Add(additem)
                    End If
                End If
            End If
            add_student_dialog.Dispose()

        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub


    Private Sub Main_Screen_Keypress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try
            If TabControl1.SelectedIndex = 0 Then
                If e.KeyChar = Chr(13) Then
                    Add_Student_Function()
                End If
            End If

            If TabControl1.SelectedIndex = 1 Then
                If e.KeyChar = Chr(13) Then
                    Add_Event_Function()
                End If
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Edit_Student_List()
    End Sub
    Private Sub Edit_Student_List()
        Try
            If ListView1.SelectedIndices.Count = 1 Then
                Dim selectedindice As Long
                selectedindice = ListView1.SelectedIndices.Item(0)
                Dim add_student As New Student_Details

                add_student.First_Name = ListView1.Items(selectedindice).SubItems(2).Text
                add_student.Surname = ListView1.Items(selectedindice).SubItems(1).Text
                add_student.Staff_Number = ListView1.Items(selectedindice).SubItems(3).Text
                add_student.Student_Number = ListView1.Items(selectedindice).Text


                Dim add_student_dialog As New Student_Details_Display(add_student.Surname, add_student.Staff_Number, add_student.First_Name, add_student.Student_Number)
                Dim result As DialogResult
                result = add_student_dialog.ShowDialog()
                Dim new_student_number As Boolean
                new_student_number = True
                If result = DialogResult.OK Then
                    If add_student.Student_Number = UCase(add_student_dialog.Student_Number.Text) Then
                        new_student_number = False
                    End If

                    add_student.First_Name = (add_student_dialog.First_Name.Text)
                    add_student.Surname = (add_student_dialog.Surname.Text)
                    add_student.Staff_Number = (add_student_dialog.Staff_Number.Text)
                    add_student.Student_Number = UCase(add_student_dialog.Student_Number.Text)

                    Dim continue_flag As Boolean
                    continue_flag = True
                    If Not add_student.Student_Number.Equals("") Then
                        If new_student_number = True Then
                            Dim i As Long
                            For i = 0 To ListView1.Items.Count - 1

                                If UCase(add_student.Student_Number) = UCase(ListView1.Items(i).Text) Then
                                    Message(UCase(add_student.Student_Number) & " cannot be added to the Student list as he/she is already present in the list.")
                                    continue_flag = False
                                End If
                            Next
                        End If
                        If continue_flag = True Then
                            ListView1.Items(selectedindice).SubItems(2).Text = (add_student_dialog.First_Name.Text.Trim())
                            ListView1.Items(selectedindice).SubItems(1).Text = (add_student_dialog.Surname.Text.Trim())
                            ListView1.Items(selectedindice).SubItems(3).Text = (add_student_dialog.Staff_Number.Text.Trim())
                            ListView1.Items(selectedindice).Text = UCase(add_student_dialog.Student_Number.Text.Trim())
                        End If
                    End If
                End If


                add_student_dialog.Dispose()
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Me.Close()
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Save_Student_List()
    End Sub

    Private Sub Save_Student_List(Optional ByVal filename As String = "")

        Try
            Dim result As DialogResult
            If filename = "" Then
                result = SaveFileDialog1.ShowDialog()
            Else
                SaveFileDialog1.FileName = filename
                result = DialogResult.OK
            End If
            If result = DialogResult.OK Then

                Dim objStreamWriter = New System.IO.StreamWriter(SaveFileDialog1.FileName, False, System.Text.Encoding.UTF8)
                Dim selectedindice As Long
                objStreamWriter.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                objStreamWriter.WriteLine("<Student_List>")

                For selectedindice = 0 To ListView1.Items.Count - 1
                    objStreamWriter.WriteLine("<Student>")
                    objStreamWriter.WriteLine("<Student_Number>" & ListView1.Items(selectedindice).Text & "</Student_Number>")
                    objStreamWriter.WriteLine("<Surname>" & ListView1.Items(selectedindice).SubItems(1).Text & "</Surname>")
                    objStreamWriter.WriteLine("<First_Name>" & ListView1.Items(selectedindice).SubItems(2).Text & "</First_Name>")
                    objStreamWriter.WriteLine("<Staff_Number>" & ListView1.Items(selectedindice).SubItems(3).Text & "</Staff_Number>")
                    objStreamWriter.WriteLine("</Student>")
                Next
                objStreamWriter.WriteLine("</Student_List>")
                objStreamWriter.close()
                If filename = "" Then
                    Message("Student list successfully saved.")
                End If

            End If
            SaveFileDialog1.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Load_Student_List(Optional ByVal filename As String = "")
        Try

            Dim result As DialogResult

            If filename = "" Then
                result = OpenFileDialog1.ShowDialog()
            Else
                OpenFileDialog1.FileName = filename
                result = DialogResult.OK
            End If

            If result = DialogResult.OK Then
                ListView1.Items.Clear()
                Dim reader As XmlReader = Nothing
                reader = New XmlTextReader(OpenFileDialog1.FileName)
                While reader.Read()

                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            If reader.Name = "Student" Then

                                Dim add_student As New Student_Details

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_student.Student_Number = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_student.Surname = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_student.First_Name = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_student.Staff_Number = reader.Value



                                Dim continue_flag As Boolean
                                continue_flag = True
                                If Not add_student.Student_Number.Trim().Equals("") Then
                                    Dim i As Long
                                    For i = 0 To ListView1.Items.Count - 1

                                        If UCase(add_student.Student_Number) = UCase(ListView1.Items(i).Text) Then
                                            Message(UCase(add_student.Student_Number) & " cannot be added to the Student list as he/she is already present in the list.")
                                            continue_flag = False
                                        End If
                                    Next
                                    If continue_flag = True Then
                                        Dim additem As New ListViewItem
                                        additem.Text = UCase(add_student.Student_Number).Trim()
                                        additem.SubItems.Add(add_student.Surname.Trim())
                                        additem.SubItems.Add(add_student.First_Name.Trim())
                                        additem.SubItems.Add(UCase(add_student.Staff_Number).Trim())
                                        ListView1.Items.Add(additem)
                                    End If

                                End If
                            End If

                    End Select
                End While

                reader.Close()
                reader = Nothing
                OpenFileDialog1.Dispose()
                If filename = "" Then
                    Message("Student List successfully loaded.")
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try


    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Load_Student_List()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Remove_From_Student_List()
    End Sub
    Private Sub Remove_From_Student_List()
        Try
            If ListView1.SelectedIndices.Count = 1 Then
                Dim selectedindice As Long
                selectedindice = ListView1.SelectedIndices.Item(0)
                ListView1.Items(selectedindice).Remove()
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            ListView1.Items.Clear()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Try
            Dim result As DialogResult
            Dim about_form As New About_Screen
            result = about_form.ShowDialog
            about_form.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Add_Computer_Lab()
    End Sub

    Private Sub Add_Computer_Lab()
        Try
            Dim result As DialogResult
            Dim add_screen As New Add_Computer_Lab
            result = add_screen.ShowDialog()
            If result = DialogResult.OK Then
                If Not add_screen.Computer_Lab.Text.Trim = "" Then
                    Dim add_lab As Boolean
                    add_lab = True
                    Dim i As Long
                    For i = 0 To ListBox1.Items.Count - 1
                        If ListBox1.Items(i) = add_screen.Computer_Lab.Text.Trim Then
                            add_lab = False
                        End If
                    Next
                    If add_lab = True Then
                        ListBox1.Items.Add(add_screen.Computer_Lab.Text.Trim)
                    End If
                End If
            End If
            add_screen.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click

        Remove_From_Computer_Lab_List()
    End Sub
    Private Sub Remove_From_Computer_Lab_List()
        Try

            If ListBox1.SelectedIndex > -1 Then
                Dim selectedindice As Long
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Edit_Computer_Lab_List()
    End Sub
    Private Sub Edit_Computer_Lab_List()
        Try
            If ListBox1.SelectedIndex > -1 Then
                Dim selectedindice As Long
                selectedindice = ListBox1.SelectedIndex

                Dim add_computer_lab_dialog As New Add_Computer_Lab(ListBox1.Items.Item(selectedindice))
                Dim result As DialogResult
                result = add_computer_lab_dialog.ShowDialog()
                Dim new_lab As Boolean
                new_lab = True
                If result = DialogResult.OK Then
                    If Convert.ToString(ListBox1.Items.Item(selectedindice)) = add_computer_lab_dialog.Computer_Lab.Text.Trim() Then
                        new_lab = False
                    End If

                    Dim continue_flag As Boolean
                    continue_flag = True
                    If Not add_computer_lab_dialog.Computer_Lab.Text.Trim().Equals("") Then
                        If new_lab = True Then
                            Dim i As Long
                            For i = 0 To ListBox1.Items.Count - 1

                                If add_computer_lab_dialog.Computer_Lab.Text.Trim() = Convert.ToString(ListBox1.Items.Item(i)) Then
                                    Message(add_computer_lab_dialog.Computer_Lab.Text.Trim() & " cannot be added to the Computer Lab list as it is already present in the list.")
                                    continue_flag = False
                                End If
                            Next
                        End If
                        If continue_flag = True Then
                            ListBox1.Items.Item(selectedindice) = add_computer_lab_dialog.Computer_Lab.Text.Trim()
                        End If
                    End If
                End If

                add_computer_lab_dialog.Dispose()
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            ListView2.Items.Clear()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Add_Event_Function()
    End Sub

    Private Sub Add_Event_Function()
        Try
            Dim add_event_dialog As Event_Details_Display
            Dim result As DialogResult
            add_event_dialog = New Event_Details_Display(ListBox1, ListView1)
            result = add_event_dialog.ShowDialog
            If result = DialogResult.OK Then
                Dim addevent As New Event_Details
                addevent.Event_Date = add_event_dialog.Event_Date.Value
                addevent.Student_Number = add_event_dialog.Student_Number.Text
                Dim starttimestring As String
                If add_event_dialog.Start_Time1.Value < 10 Then
                    starttimestring = "0" & add_event_dialog.Start_Time1.Value & ":"
                Else
                    starttimestring = add_event_dialog.Start_Time1.Value & ":"
                End If
                If add_event_dialog.Start_Time2.Value < 10 Then
                    starttimestring = starttimestring & "0" & add_event_dialog.Start_Time2.Value
                Else
                    starttimestring = starttimestring & add_event_dialog.Start_Time2.Value
                End If
                Dim endtimestring As String
                If add_event_dialog.End_Time1.Value < 10 Then
                    endtimestring = "0" & add_event_dialog.End_Time1.Value & ":"
                Else
                    endtimestring = add_event_dialog.End_Time1.Value & ":"
                End If
                If add_event_dialog.End_Time2.Value < 10 Then
                    endtimestring = endtimestring & "0" & add_event_dialog.End_Time2.Value
                Else
                    endtimestring = endtimestring & add_event_dialog.End_Time2.Value
                End If
                addevent.Start_Time = starttimestring
                addevent.End_Time = endtimestring

                addevent.Computer_Lab = add_event_dialog.Computer_Lab.Text

                Dim toadd As New ListViewItem
                toadd.Text = addevent.Event_Date.Date
                toadd.SubItems.Add(addevent.Student_Number)
                toadd.SubItems.Add(addevent.Start_Time)
                toadd.SubItems.Add(addevent.End_Time)
                toadd.SubItems.Add(addevent.Computer_Lab)
                ListView2.Items.Add(toadd)
            End If
            add_event_dialog.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        Save_Timesheet_Events_List()
    End Sub

    Private Sub Save_Timesheet_Events_List(Optional ByVal filename As String = "")
        Try
            Dim result As DialogResult
            If filename = "" Then
                result = SaveFileDialog2.ShowDialog()
            Else
                SaveFileDialog2.FileName = filename
                result = DialogResult.OK
            End If

            If result = DialogResult.OK Then

                Dim objStreamWriter = New System.IO.StreamWriter(SaveFileDialog2.FileName, False, System.Text.Encoding.UTF8)
                Dim selectedindice As Long
                objStreamWriter.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                objStreamWriter.WriteLine("<Timesheet_Events_List>")

                For selectedindice = 0 To ListBox1.Items.Count - 1
                    objStreamWriter.WriteLine("<Computer_Labs>")
                    objStreamWriter.WriteLine("<Lab_Name>" & Convert.ToString(ListBox1.Items.Item(selectedindice)) & "</Lab_Name>")
                    objStreamWriter.WriteLine("</Computer_Labs>")
                Next

                For selectedindice = 0 To ListView2.Items.Count - 1
                    objStreamWriter.WriteLine("<Event>")
                    objStreamWriter.WriteLine("<Event_Date>" & ListView2.Items(selectedindice).Text & "</Event_Date>")
                    objStreamWriter.WriteLine("<Student_Number>" & ListView2.Items(selectedindice).SubItems(1).Text & "</Student_Number>")
                    objStreamWriter.WriteLine("<Start_Time>" & ListView2.Items(selectedindice).SubItems(2).Text & "</Start_Time>")
                    objStreamWriter.WriteLine("<End_Time>" & ListView2.Items(selectedindice).SubItems(3).Text & "</End_Time>")
                    objStreamWriter.WriteLine("<Computer_Lab>" & ListView2.Items(selectedindice).SubItems(4).Text & "</Computer_Lab>")
                    objStreamWriter.WriteLine("</Event>")
                Next
                objStreamWriter.WriteLine("</Timesheet_Events_List>")
                objStreamWriter.close()
                If filename = "" Then
                    Message("Timesheet Events list successfully saved.")
                End If

            End If
                SaveFileDialog2.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        Load_Timesheet_Events_List()
    End Sub

    Private Sub Load_Timesheet_Events_List(Optional ByVal filename As String = "")

        Try

            Dim result As DialogResult

            If filename = "" Then
                result = OpenFileDialog2.ShowDialog()
            Else
                OpenFileDialog2.FileName = filename
                result = DialogResult.OK
            End If

            If result = DialogResult.OK Then
                ListView2.Items.Clear()
                ListBox1.Items.Clear()
                Dim reader As XmlReader = Nothing
                reader = New XmlTextReader(OpenFileDialog2.FileName)
                While reader.Read()

                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            If reader.Name = "Event" Then

                                Dim add_event As New Event_Details

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_event.Event_Date_String = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_event.Student_Number = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_event.Start_Time = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_event.End_Time = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                add_event.Computer_Lab = reader.Value

                                Dim continue_flag As Boolean
                                continue_flag = True
                                If Not add_event.Event_Date_String.Trim().Equals("") Then


                                    If continue_flag = True Then
                                        Dim additem As New ListViewItem
                                        additem.Text = add_event.Event_Date_String.Trim()
                                        additem.SubItems.Add(add_event.Student_Number.Trim())
                                        additem.SubItems.Add(add_event.Start_Time.Trim())
                                        additem.SubItems.Add(add_event.End_Time.Trim())
                                        additem.SubItems.Add(add_event.Computer_Lab.Trim())
                                        ListView2.Items.Add(additem)
                                    End If

                                End If
                            End If

                            If reader.Name = "Computer_Labs" Then



                                reader.Read()
                                reader.Read()
                                reader.Read()



                                Dim continue_flag As Boolean
                                continue_flag = True
                                If Not reader.Value.Trim().Equals("") Then
                                    Dim i As Long
                                    For i = 0 To ListBox1.Items.Count - 1

                                        If reader.Value = Convert.ToString(ListBox1.Items.Item(i)) Then
                                            Message(reader.Value.Trim & " cannot be added to the Computer lab list as it is already present in the list.")
                                            continue_flag = False
                                        End If
                                    Next
                                    If continue_flag = True Then

                                        ListBox1.Items.Add(reader.Value)
                                    End If

                                End If
                            End If

                    End Select
                End While

                reader.Close()
                reader = Nothing
                OpenFileDialog2.Dispose()
                If filename = "" Then
                    Message("Timesheet Events List successfully loaded.")
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try


    End Sub


    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Remove_From_Timesheet_Events_List()
    End Sub
    Private Sub Remove_From_Timesheet_Events_List()
        Try
            If ListView2.SelectedIndices.Count = 1 Then
                Dim selectedindice As Long
                selectedindice = ListView2.SelectedIndices.Item(0)
                ListView2.Items(selectedindice).Remove()
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub

    Private Sub Reset_ListView2_Column_Widths(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Try
            ListView2.Columns.Item(0).Width = 75
            ListView2.Columns.Item(1).Width = 97
            ListView2.Columns.Item(2).Width = 90
            ListView2.Columns.Item(3).Width = 90
            ListView2.Columns.Item(4).Width = 112
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try


    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        Save_Job_Control_File()
    End Sub

    Private Sub Save_Job_Control_File(Optional ByVal filename As String = "")
        Try
            Dim result As DialogResult

            If filename = "" Then
                result = SaveFileDialog3.ShowDialog()
            Else
                SaveFileDialog3.FileName = filename
                result = DialogResult.OK
            End If
            If result = DialogResult.OK Then

                Dim objStreamWriter = New System.IO.StreamWriter(SaveFileDialog3.FileName, False, System.Text.Encoding.UTF8)
                objStreamWriter.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                objStreamWriter.WriteLine("<Job_Control>")

                objStreamWriter.WriteLine("<Student_List>" & SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Students.slt" & "</Student_List>")
                objStreamWriter.WriteLine("<Timesheet_Events_List>" & SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Timesheet.tlt" & "</Timesheet_Events_List>")
                objStreamWriter.WriteLine("<Report_Data>" & SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Report.rlt" & "</Report_Data>")

                Save_Student_List(SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Students.slt")
                Save_Timesheet_Events_List(SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Timesheet.tlt")
                Save_Report_Data(SaveFileDialog3.FileName.Remove(SaveFileDialog3.FileName.Length - 4, 4) & "_Report.rlt")

                objStreamWriter.WriteLine("</Job_Control>")
                objStreamWriter.close()
                If filename = "" Then
                    Message("Job Control File successfully saved.")
                End If

            End If
            SaveFileDialog3.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem19.Click
        Load_Job_Control_File()
    End Sub

    Private Sub Load_Job_Control_File(Optional ByVal filename As String = "")
        Try

            Dim result As DialogResult

            If filename = "" Then
                result = OpenFileDialog3.ShowDialog()
            Else
                OpenFileDialog3.FileName = filename
                result = DialogResult.OK
            End If

            If result = DialogResult.OK Then

                Dim reader As XmlReader = Nothing
                reader = New XmlTextReader(OpenFileDialog3.FileName)
                While reader.Read()

                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            If reader.Name = "Student_List" Then
                                reader.Read()
                                Load_Student_List(reader.Value)
                            End If
                            If reader.Name = "Timesheet_Events_List" Then
                                reader.Read()
                                Load_Timesheet_Events_List(reader.Value)
                            End If
                            If reader.Name = "Report_Data" Then
                                reader.Read()
                                Load_Report_Data(reader.Value)
                            End If

                    End Select
                End While

                reader.Close()
                reader = Nothing
                OpenFileDialog3.Dispose()
                If filename = "" Then
                    Message("Job Control File successfully loaded.")
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try


    End Sub

    Public Sub Program_Exit(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try

        
        Dim result As MsgBoxResult
            result = MsgBox("Would you like to save your current project?", MsgBoxStyle.YesNo, "Lab Tutor Payment System")
            If result = MsgBoxResult.Yes Or result = MsgBoxResult.OK Then
                Save_Job_Control_File()
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Edit_Timesheet_Events_List()
    End Sub

    Private Sub Edit_Timesheet_Events_List()
        Try
            If ListView2.SelectedIndices.Count = 1 Then
                Dim selectedindice As Long
                selectedindice = ListView2.SelectedIndices.Item(0)
                Dim add_event As New Event_Details

                add_event.Event_Date = ListView2.Items(selectedindice).Text
                add_event.Student_Number = ListView2.Items(selectedindice).SubItems(1).Text
                add_event.Start_Time = ListView2.Items(selectedindice).SubItems(2).Text
                add_event.End_Time = ListView2.Items(selectedindice).SubItems(3).Text
                add_event.Computer_Lab = ListView2.Items(selectedindice).SubItems(4).Text

                Dim add_event_dialog As New Event_Details_Display(ListBox1, ListView1, add_event.Event_Date, add_event.Student_Number, add_event.Start_Time, add_event.End_Time, add_event.Computer_Lab)
                Dim result As DialogResult
                result = add_event_dialog.ShowDialog()
                Dim new_event_date As Boolean
                new_event_date = True
                If result = DialogResult.OK Then
                    If add_event.Event_Date = UCase(add_event_dialog.Event_Date.Value) Then
                        new_event_date = False
                    End If

                    add_event.Event_Date = add_event_dialog.Event_Date.Value
                    add_event.Student_Number = add_event_dialog.Student_Number.Text
                    Dim starttimestring As String
                    If add_event_dialog.Start_Time1.Value < 10 Then
                        starttimestring = "0" & add_event_dialog.Start_Time1.Value & ":"
                    Else
                        starttimestring = add_event_dialog.Start_Time1.Value & ":"
                    End If
                    If add_event_dialog.Start_Time2.Value < 10 Then
                        starttimestring = starttimestring & "0" & add_event_dialog.Start_Time2.Value
                    Else
                        starttimestring = starttimestring & add_event_dialog.Start_Time2.Value
                    End If
                    Dim endtimestring As String
                    If add_event_dialog.End_Time1.Value < 10 Then
                        endtimestring = "0" & add_event_dialog.End_Time1.Value & ":"
                    Else
                        endtimestring = add_event_dialog.End_Time1.Value & ":"
                    End If
                    If add_event_dialog.End_Time2.Value < 10 Then
                        endtimestring = endtimestring & "0" & add_event_dialog.End_Time2.Value
                    Else
                        endtimestring = endtimestring & add_event_dialog.End_Time2.Value
                    End If
                    add_event.Start_Time = starttimestring
                    add_event.End_Time = endtimestring

                    add_event.Computer_Lab = add_event_dialog.Computer_Lab.Text
                  


                    Dim toadd As New ListViewItem
                    toadd.Text = add_event.Event_Date.Date
                    toadd.SubItems.Add(add_event.Student_Number)
                    toadd.SubItems.Add(add_event.Start_Time)
                    toadd.SubItems.Add(add_event.End_Time)
                    toadd.SubItems.Add(add_event.Computer_Lab)

                    ListView2.Items(selectedindice) = toadd



            End If


            add_event_dialog.Dispose()
            End If
        Catch except As Exception
            Error_Handler(except.Message)
        End Try
    End Sub


    Private Sub Save_Report_Data(Optional ByVal filename As String = "")

        Try
            Dim result As DialogResult
            If filename = "" Then
                result = SaveFileDialog4.ShowDialog()
            Else
                SaveFileDialog4.FileName = filename
                result = DialogResult.OK
            End If
            If result = DialogResult.OK Then

                Dim objStreamWriter = New System.IO.StreamWriter(SaveFileDialog4.FileName, False, System.Text.Encoding.UTF8)
                Dim selectedindice As Long
                objStreamWriter.WriteLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                objStreamWriter.WriteLine("<Report_Data>")

                objStreamWriter.WriteLine("<Payment_Period_Start>" & Payment_Period_Start.Value & "</Payment_Period_Start>")
                objStreamWriter.WriteLine("<Payment_Period_End>" & Payment_Period_End.Value & "</Payment_Period_End>")
                objStreamWriter.WriteLine("<Department>" & Department.Text & "</Department>")
                objStreamWriter.WriteLine("<Cost_Centre>" & Cost_Centre.Text & "</Cost_Centre>")
                objStreamWriter.WriteLine("<Fund>" & Fund.Text & "</Fund>")
                objStreamWriter.WriteLine("<Hourly_Rate_Rands>" & Hourly_Rate_Rands.Value & "</Hourly_Rate_Rands>")
                objStreamWriter.WriteLine("<Hourly_Rate_Cents>" & Hourly_Rate_Cents.Value & "</Hourly_Rate_Cents>")
                objStreamWriter.WriteLine("<Report_Compiler>" & Report_Compiler.Text & "</Report_Compiler>")
                objStreamWriter.WriteLine("<Create_Subfolder>" & Create_Subfolder.Checked.ToString & "</Create_Subfolder>")
                objStreamWriter.WriteLine("<Overwrite_Files>" & Overwrite_Files.Checked.ToString & "</Overwrite_Files>")
                objStreamWriter.WriteLine("<Show_Created_Files>" & Show_Created_Files.Checked.ToString & "</Show_Created_Files>")
                objStreamWriter.WriteLine("</Report_Data>")
                objStreamWriter.close()
                If filename = "" Then
                    Message("Report data successfully saved.")
                End If

            End If
            SaveFileDialog4.Dispose()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Load_Report_Data(Optional ByVal filename As String = "")
        Try

            Dim result As DialogResult

            If filename = "" Then
                result = OpenFileDialog4.ShowDialog()
            Else
                OpenFileDialog4.FileName = filename
                result = DialogResult.OK
            End If

            If result = DialogResult.OK Then
                Payment_Period_Start.Value = Now()
                Payment_Period_End.Value = Now()
                Department.Text = ""
                Cost_Centre.Text = ""
                Fund.Text = ""
                Hourly_Rate_Rands.Value = 0
                Hourly_Rate_Cents.Value = 0
                Dim reader As XmlReader = Nothing
                reader = New XmlTextReader(OpenFileDialog4.FileName)
                While reader.Read()

                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            If reader.Name = "Report_Data" Then

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Payment_Period_Start.Value = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Payment_Period_End.Value = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Department.Text = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Cost_Centre.Text = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Fund.Text = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Hourly_Rate_Rands.Value = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Hourly_Rate_Cents.Value = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Report_Compiler.Text = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Create_Subfolder.Checked = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Overwrite_Files.Checked = reader.Value

                                reader.Read()
                                reader.Read()
                                reader.Read()
                                reader.Read()
                                Show_Created_Files.Checked = reader.Value
                            End If

                    End Select
                End While

                reader.Close()
                reader = Nothing
                OpenFileDialog4.Dispose()
                If filename = "" Then
                    Message("Report data successfully loaded.")
                End If

            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try


    End Sub

    Private Sub Refresh_Form()
        Payment_Period_Start.Value = Now()
        Payment_Period_End.Value = Now()
        Department.Text = ""
        Cost_Centre.Text = ""
        Fund.Text = ""
        Hourly_Rate_Rands.Value = 0
        Hourly_Rate_Cents.Value = 0
        ListView1.Items.Clear()
        ListView2.Items.Clear()
        ListBox1.Items.Clear()
        Create_Subfolder.Checked = True
        Overwrite_Files.Checked = True
        Show_Created_Files.Checked = True
        Report_Compiler.Text = ""
    End Sub


    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        Save_Report_Data()
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem21.Click
        Load_Report_Data()
    End Sub

    Private Sub MenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem25.Click
        Refresh_Form()
    End Sub

    Private Sub Generate_Student_Timesheet(Optional ByVal student_number As String = "")

        Try
            If Validate_Report_Data() = True Then

                
            Dim result As DialogResult
            result = FolderBrowserDialog1.ShowDialog
                If result = DialogResult.OK Then
                    Dim initial_overwrite_status As Boolean
                    initial_overwrite_status = Overwrite_Files.Checked
                    Dim filetowrite As String
                    Dim foldername As String
                    Dim writefiles As Boolean
                    writefiles = True
                    Dim writtencounter As Long
                    writtencounter = 0


                    If Create_Subfolder.Checked = False Then
                        foldername = FolderBrowserDialog1.SelectedPath
                    Else
                        foldername = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString
                    End If


                    Dim i As Long
                    For i = 0 To ListView1.Items.Count - 1
                        Dim objStreamWriter As StreamWriter
                        If Create_Subfolder.Checked = False Then
                            filetowrite = FolderBrowserDialog1.SelectedPath & "\" & ListView1.Items(i).Text & ".doc"
                            foldername = FolderBrowserDialog1.SelectedPath
                        Else

                            Dim dinfo As New System.IO.DirectoryInfo(FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString)
                            If dinfo.Exists() = False Then
                                dinfo.Create()
                            End If
                            filetowrite = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString & "\" & ListView1.Items(i).Text & ".doc"
                            foldername = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString
                        End If
                        If Overwrite_Files.Checked = False And writefiles = True Then

                            Dim finfo As New System.IO.FileInfo(filetowrite)
                            If finfo.Exists() = True Then
                                result = MsgBox("The file to be generated already exists at the currently selected location." & vbCrLf & "Do you wish to overwrite this and all other conflicting files?", MsgBoxStyle.YesNo, "Lab Tutor Payment System")
                            End If
                            If result = DialogResult.Yes Or result = DialogResult.OK Then
                                writefiles = True
                                Overwrite_Files.Checked = True
                            Else
                                writefiles = False
                                Overwrite_Files.Checked = False
                            End If
                        End If
                        If writefiles = True Then


                            objStreamWriter = New System.IO.StreamWriter(filetowrite, False, System.Text.Encoding.UTF8)
                            Dim difference As Double
                            difference = DateDiff(DateInterval.Day, Payment_Period_Start.Value, Payment_Period_End.Value)
                            Progress_Bar1_Panel.Visible = True
                            ProgressBar1.Value = 0
                            ProgressBar1.Maximum = difference
                            '**************************************
                            objStreamWriter.WriteLine("MIME-Version: 1.0")
                            objStreamWriter.WriteLine("Content-Location: file:///C:/1165A114/Timesheet.htm")
                            objStreamWriter.WriteLine("Content-Transfer-Encoding: quoted-printable")
                            objStreamWriter.WriteLine("Content-Type: text/html; charset=""us-ascii""")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<html xmlns:o=3D""urn:schemas-microsoft-com:office:office""")
                            objStreamWriter.WriteLine("xmlns:w=3D""urn:schemas-microsoft-com:office:word""")
                            objStreamWriter.WriteLine("xmlns=3D""http://www.w3.org/TR/REC-html40"">")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<head>")
                            objStreamWriter.WriteLine("<meta http-equiv=3DContent-Type content=3D""text/html; charset=3Dus-ascii"">")
                            objStreamWriter.WriteLine("<meta name=3DProgId content=3DWord.Document>")
                            objStreamWriter.WriteLine("<meta name=3DGenerator content=3D""Microsoft Word 11"">")
                            objStreamWriter.WriteLine("<meta name=3DOriginator content=3D""Microsoft Word 11"">")
                            objStreamWriter.WriteLine("<link rel=3DFile-List href=3D""Timesheet_files/filelist.xml"">")
                            objStreamWriter.WriteLine("<title>TIMESHEET FOR LESLIE COMMERCE IT, UCT</title>")
                            objStreamWriter.WriteLine("<!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <o:DocumentProperties>")
                            objStreamWriter.WriteLine("  <o:Author>Commerce I.T.</o:Author>")
                            objStreamWriter.WriteLine("  <o:LastAuthor>Commerce I.T.</o:LastAuthor>")
                            objStreamWriter.WriteLine("  <o:Revision>2</o:Revision>")
                            objStreamWriter.WriteLine("  <o:TotalTime>2</o:TotalTime>")
                            objStreamWriter.WriteLine("  <o:Created>" & Now() & "</o:Created>")
                            objStreamWriter.WriteLine("  <o:LastSaved>" & Now() & "</o:LastSaved>")
                            objStreamWriter.WriteLine("  <o:Pages>1</o:Pages>")
                            objStreamWriter.WriteLine("  <o:Words>38</o:Words>")
                            objStreamWriter.WriteLine("  <o:Characters>221</o:Characters>")
                            objStreamWriter.WriteLine("  <o:Company>University of Cape Town</o:Company>")
                            objStreamWriter.WriteLine("  <o:Lines>1</o:Lines>")
                            objStreamWriter.WriteLine("  <o:Paragraphs>1</o:Paragraphs>")
                            objStreamWriter.WriteLine("  <o:CharactersWithSpaces>258</o:CharactersWithSpaces>")
                            objStreamWriter.WriteLine("  <o:Version>11.5606</o:Version>")
                            objStreamWriter.WriteLine(" </o:DocumentProperties>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <w:WordDocument>")
                            objStreamWriter.WriteLine("  <w:SpellingState>Clean</w:SpellingState>")
                            objStreamWriter.WriteLine("  <w:GrammarState>Clean</w:GrammarState>")
                            objStreamWriter.WriteLine("  <w:DisplayHorizontalDrawingGridEvery>0</w:DisplayHorizontalDrawingGridEve=")
                            objStreamWriter.WriteLine("ry>")
                            objStreamWriter.WriteLine("  <w:DisplayVerticalDrawingGridEvery>0</w:DisplayVerticalDrawingGridEvery>")
                            objStreamWriter.WriteLine("  <w:UseMarginsForDrawingGridOrigin/>")
                            objStreamWriter.WriteLine("  <w:ValidateAgainstSchemas/>")
                            objStreamWriter.WriteLine("  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>")
                            objStreamWriter.WriteLine("  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>")
                            objStreamWriter.WriteLine("  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>")
                            objStreamWriter.WriteLine("  <w:Compatibility>")
                            objStreamWriter.WriteLine("   <w:FootnoteLayoutLikeWW8/>")
                            objStreamWriter.WriteLine("   <w:ShapeLayoutLikeWW8/>")
                            objStreamWriter.WriteLine("   <w:AlignTablesRowByRow/>")
                            objStreamWriter.WriteLine("   <w:ForgetLastTabAlignment/>")
                            objStreamWriter.WriteLine("   <w:LayoutRawTableWidth/>")
                            objStreamWriter.WriteLine("   <w:LayoutTableRowsApart/>")
                            objStreamWriter.WriteLine("   <w:UseWord97LineBreakingRules/>")
                            objStreamWriter.WriteLine("   <w:SelectEntireFieldWithStartOrEnd/>")
                            objStreamWriter.WriteLine("   <w:UseWord2002TableStyleRules/>")
                            objStreamWriter.WriteLine("  </w:Compatibility>")
                            objStreamWriter.WriteLine("  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>")
                            objStreamWriter.WriteLine(" </w:WordDocument>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <w:LatentStyles DefLockedState=3D""false"" LatentStyleCount=3D""156"">")
                            objStreamWriter.WriteLine(" </w:LatentStyles>")
                            objStreamWriter.WriteLine("</xml><![endif]-->")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine("<!--")
                            objStreamWriter.WriteLine(" /* Style Definitions */")
                            objStreamWriter.WriteLine(" p.MsoNormal, li.MsoNormal, div.MsoNormal")
                            objStreamWriter.WriteLine("	{mso-style-parent:"""";")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("h1")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:1;")
                            objStreamWriter.WriteLine("	font-size:12.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-font-kerning:0pt;")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoTitle, li.MsoTitle, div.MsoTitle")
                            objStreamWriter.WriteLine("	{margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	font-size:18.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-weight:bold;}")
                            objStreamWriter.WriteLine("@page Section1")
                            objStreamWriter.WriteLine("	{size:595.3pt 841.9pt;")
                            objStreamWriter.WriteLine("	margin:72.0pt 90.0pt 72.0pt 90.0pt;")
                            objStreamWriter.WriteLine("	mso-header-margin:36.0pt;")
                            objStreamWriter.WriteLine("	mso-footer-margin:36.0pt;")
                            objStreamWriter.WriteLine("	mso-paper-source:0;}")
                            objStreamWriter.WriteLine("div.Section1")
                            objStreamWriter.WriteLine("	{page:Section1;}")
                            objStreamWriter.WriteLine("-->")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<!--[if gte mso 10]>")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine(" /* Style Definitions */")
                            objStreamWriter.WriteLine(" table.MsoNormalTable")
                            objStreamWriter.WriteLine("	{mso-style-name:""Table Normal"";")
                            objStreamWriter.WriteLine("	mso-tstyle-rowband-size:0;")
                            objStreamWriter.WriteLine("	mso-tstyle-colband-size:0;")
                            objStreamWriter.WriteLine("	mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	mso-style-parent:"""";")
                            objStreamWriter.WriteLine("	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("	mso-para-margin:0cm;")
                            objStreamWriter.WriteLine("	mso-para-margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-ansi-language:#0400;")
                            objStreamWriter.WriteLine("	mso-fareast-language:#0400;")
                            objStreamWriter.WriteLine("	mso-bidi-language:#0400;}")
                            objStreamWriter.WriteLine("table.MsoTableGrid")
                            objStreamWriter.WriteLine("	{mso-style-name:""Table Grid"";")
                            objStreamWriter.WriteLine("	mso-tstyle-rowband-size:0;")
                            objStreamWriter.WriteLine("	mso-tstyle-colband-size:0;")
                            objStreamWriter.WriteLine("	border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("	mso-border-alt:solid windowtext .5pt;")
                            objStreamWriter.WriteLine("	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("	mso-border-insideh:.5pt solid windowtext;")
                            objStreamWriter.WriteLine("	mso-border-insidev:.5pt solid windowtext;")
                            objStreamWriter.WriteLine("	mso-para-margin:0cm;")
                            objStreamWriter.WriteLine("	mso-para-margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-ansi-language:#0400;")
                            objStreamWriter.WriteLine("	mso-fareast-language:#0400;")
                            objStreamWriter.WriteLine("	mso-bidi-language:#0400;}")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<![endif]-->")
                            objStreamWriter.WriteLine("</head>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<body lang=3DEN-GB style=3D'tab-interval:36.0pt'>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<div class=3DSection1>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoTitle>TIMESHEET FOR COMMERCE IT, UCT</p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal align=3Dcenter style=3D'text-align:center'><b><span")
                            objStreamWriter.WriteLine("style=3D'font-size:18.0pt;mso-bidi-font-size:10.0pt'>HELPDESK STAFF<o:p></o=")
                            objStreamWriter.WriteLine(":p></span></b></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal align=3Dcenter style=3D'text-align:center'><b><span")
                            objStreamWriter.WriteLine("style=3D'font-size:18.0pt;mso-bidi-font-size:10.0pt'><o:p>&nbsp;</o:p></spa=")
                            objStreamWriter.WriteLine("n></b></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=3DMsoTableGrid border=3D0 cellspacing=3D0 cellpadding=3D0")
                            objStreamWriter.WriteLine(" style=3D'border-collapse:collapse;mso-yfti-tbllook:480;mso-padding-alt:0cm=")
                            objStreamWriter.WriteLine(" 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:0;mso-yfti-firstrow:yes;height:25.15pt'>")
                            objStreamWriter.WriteLine("  <td width=3D284 valign=3Dtop style=3D'width:213.05pt;padding:0cm 5.4pt 0c=")
                            objStreamWriter.WriteLine("m 5.4pt;")
                            objStreamWriter.WriteLine("  height:25.15pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal><b><u>NAME:&nbsp;&nbsp;&nbsp;" & ListView1.Items(i).SubItems(2).Text & "&nbsp;" & ListView1.Items(i).SubItems(1).Text & "<o:p></o:p></u></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=3D284 valign=3Dtop style=3D'width:213.05pt;padding:0cm 5.4pt 0c=")
                            objStreamWriter.WriteLine("m 5.4pt;")
                            objStreamWriter.WriteLine("  height:25.15pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal><b><u>STUDENT NO:&nbsp;&nbsp;&nbsp;" & ListView1.Items(i).Text & "<o:p></o:p></u></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:1;mso-yfti-lastrow:yes;height:21.2pt'>")
                            objStreamWriter.WriteLine("  <td width=3D284 valign=3Dtop style=3D'width:213.05pt;padding:0cm 5.4pt 0c=")
                            objStreamWriter.WriteLine("m 5.4pt;")
                            objStreamWriter.WriteLine("  height:21.2pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal><b><u>FROM:&nbsp;&nbsp;&nbsp;" & Payment_Period_Start.Value.ToLongDateString & "<o:p></o:p></u></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=3D284 valign=3Dtop style=3D'width:213.05pt;padding:0cm 5.4pt 0c=")
                            objStreamWriter.WriteLine("m 5.4pt;")
                            objStreamWriter.WriteLine("  height:21.2pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal><b><u>TO:&nbsp;&nbsp;&nbsp;" & Payment_Period_End.Value.ToLongDateString & "<o:p></o:p></u></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=3DMsoNormalTable border=3D1 cellspacing=3D0 cellpadding=3D0")
                            objStreamWriter.WriteLine(" style=3D'border-collapse:collapse;border:none;mso-border-alt:solid windowt=")
                            objStreamWriter.WriteLine("ext .5pt;")
                            objStreamWriter.WriteLine(" mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowte=")
                            objStreamWriter.WriteLine("xt;")
                            objStreamWriter.WriteLine(" mso-border-insidev:.5pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:0;mso-yfti-firstrow:yes'>")
                            objStreamWriter.WriteLine("  <td width=3D158 valign=3Dtop style=3D'width:118.8pt;border:solid windowte=")
                            objStreamWriter.WriteLine("xt 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal align=3Dcenter style=3D'text-align:center'><b><span")
                            objStreamWriter.WriteLine("  style=3D'font-size:12.0pt;mso-bidi-font-size:10.0pt'>DATE<o:p></o:p></spa=")
                            objStreamWriter.WriteLine("n></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=3D95 valign=3Dtop style=3D'width:70.9pt;border:solid windowtext=")
                            objStreamWriter.WriteLine(" 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal align=3Dcenter style=3D'text-align:center'><b><span")
                            objStreamWriter.WriteLine("  style=3D'font-size:12.0pt;mso-bidi-font-size:10.0pt'>HOURS<o:p></o:p></sp=")
                            objStreamWriter.WriteLine("an></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=3D315 valign=3Dtop style=3D'width:236.4pt;border:solid windowte=")
                            objStreamWriter.WriteLine("xt 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=3DMsoNormal align=3Dcenter style=3D'text-align:center'><b><span")
                            objStreamWriter.WriteLine("  style=3D'font-size:12.0pt;mso-bidi-font-size:10.0pt'>PROJECT<o:p></o:p></=")
                            objStreamWriter.WriteLine("span></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            Dim counter As Long
                            Dim rownumber As Long
                            rownumber = 1
                            Dim start_date, end_date As Date
                            start_date = Payment_Period_Start.Value
                            end_date = Payment_Period_End.Value
                            start_date = New Date(Year(start_date), Month(start_date), Day(start_date))
                            end_date = New Date(Year(end_date), Month(end_date), Day(end_date))
                            Dim datafound As Boolean
                            datafound = False
                            Dim hours, hours_total As Double
                            Dim hours_string As String
                            hours_total = 0
                            hours_string = ""
                            hours = 0


                            While DateDiff(DateInterval.Day, start_date, end_date) >= 0

                                hours_string = ""
                                datafound = False
                                ProgressBar1.Increment(1)

                                For counter = 0 To ListView2.Items.Count - 1
                                    If ListView2.Items(counter).SubItems(1).Text = ListView1.Items(i).Text Then
                                        hours = 0
                                        Dim add_event As New Event_Details

                                        add_event.Event_Date = ListView2.Items(counter).Text
                                        add_event.Student_Number = ListView2.Items(counter).SubItems(1).Text
                                        add_event.Start_Time = ListView2.Items(counter).SubItems(2).Text
                                        add_event.End_Time = ListView2.Items(counter).SubItems(3).Text
                                        add_event.Computer_Lab = ListView2.Items(counter).SubItems(4).Text

                                        'MsgBox(start_date & " <--> " & add_event.Event_Date & " <-- " & DateDiff(DateInterval.Day, start_date, add_event.Event_Date))
                                        If DateDiff(DateInterval.Day, add_event.Event_Date, start_date) = 0 Then
                                            datafound = True
                                            Dim starttime, endtime As Date
                                            Dim myarray As Array
                                            myarray = add_event.Start_Time.Split(":")
                                            starttime = New Date(2000, 1, 1, myarray(0), myarray(1), 0)
                                            myarray = add_event.End_Time.Split(":")
                                            endtime = New Date(2000, 1, 1, myarray(0), myarray(1), 0)
                                            hours = hours + Math.Round((DateDiff(DateInterval.Minute, starttime, endtime) / 60), 1)
                                            hours_string = hours_string & " + " & hours
                                            hours_total = hours_total + hours


                                        End If
                                    End If
                                Next
                                If datafound = False Then
                                    objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:" & rownumber & "'>")
                                    objStreamWriter.WriteLine("  <td width=3D158 valign=3Dtop style=3D'width:118.8pt;border:solid windowte=")
                                    objStreamWriter.WriteLine("xt 1.0pt;")
                                    objStreamWriter.WriteLine("  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:s=")
                                    objStreamWriter.WriteLine("olid windowtext .5pt;")
                                    objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>" & start_date.ToLongDateString & "</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=3D95 valign=3Dtop style=3D'width:70.9pt;border-top:none;border-=")
                                    objStreamWriter.WriteLine("left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid window=")
                                    objStreamWriter.WriteLine("text .5pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=3D315 valign=3Dtop style=3D'width:236.4pt;border-top:none;borde=")
                                    objStreamWriter.WriteLine("r-left:")
                                    objStreamWriter.WriteLine("  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1=")
                                    objStreamWriter.WriteLine(".0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid window=")
                                    objStreamWriter.WriteLine("text .5pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine(" </tr>")
                                Else
                                    objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:" & rownumber & "'>")
                                    objStreamWriter.WriteLine("  <td width=3D158 valign=3Dtop style=3D'width:118.8pt;border:solid windowte=")
                                    objStreamWriter.WriteLine("xt 1.0pt;")
                                    objStreamWriter.WriteLine("  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:s=")
                                    objStreamWriter.WriteLine("olid windowtext .5pt;")
                                    objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>" & start_date.ToLongDateString & "</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=3D95 valign=3Dtop style=3D'width:70.9pt;border-top:none;border-=")
                                    objStreamWriter.WriteLine("left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid window=")
                                    objStreamWriter.WriteLine("text .5pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>" & hours_string.Remove(0, 3) & "</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=3D315 valign=3Dtop style=3D'width:236.4pt;border-top:none;borde=")
                                    objStreamWriter.WriteLine("r-left:")
                                    objStreamWriter.WriteLine("  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1=")
                                    objStreamWriter.WriteLine(".0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid window=")
                                    objStreamWriter.WriteLine("text .5pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                                    objStreamWriter.WriteLine("  <p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine(" </tr>")
                                End If
                                start_date = start_date.AddDays(1)
                            End While

                            objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:30;page-break-inside:avoid=")
                            objStreamWriter.WriteLine("'>")
                            objStreamWriter.WriteLine("  <td width=3D568 colspan=3D3 valign=3Dtop style=3D'width:426.1pt;border:so=")
                            objStreamWriter.WriteLine("lid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:s=")
                            objStreamWriter.WriteLine("olid windowtext .5pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <h1>&nbsp;</h1>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")

                            objStreamWriter.WriteLine(" <tr style=3D'mso-yfti-irow:31;mso-yfti-lastrow:yes;page-break-inside:avoid=")
                            objStreamWriter.WriteLine("'>")
                            objStreamWriter.WriteLine("  <td width=3D568 colspan=3D3 valign=3Dtop style=3D'width:426.1pt;border:so=")
                            objStreamWriter.WriteLine("lid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:s=")
                            objStreamWriter.WriteLine("olid windowtext .5pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <h1>TOTAL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & hours_total & "</h1>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=3DMsoNormal><b><u>PAYMENT ORDER COMPILED BY:<o:p>&nbsp;&nbsp;&nbsp;" & Report_Compiler.Text & "</o:p></u></b></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("</div>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("</body>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("</html>")

                            '**************************************
                            writtencounter = writtencounter + 1
                            objStreamWriter.Close()
                        End If
                    Next

                    Message("Student Timesheets successfully generated." & vbCrLf & writtencounter & " files were created.")
                    ProgressBar1.Value = 0
                    Progress_Bar1_Panel.Visible = False
                    Overwrite_Files.Checked = initial_overwrite_status
                    If Show_Created_Files.Checked = True Then
                        System.Diagnostics.Process.Start(foldername)
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Department_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Department.TextChanged
        Try
            Department.Text = Department.Text.Trim
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Cost_Centre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cost_Centre.TextChanged
        Try
            Cost_Centre.Text = Cost_Centre.Text.Trim
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Fund_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fund.TextChanged
        Try
            Fund.Text = Fund.Text.Trim
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Hourly_Rate_Rands_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Hourly_Rate_Rands.Leave
        Try
            If Hourly_Rate_Rands.Value > Hourly_Rate_Rands.Maximum Then
                Hourly_Rate_Rands.Value = Hourly_Rate_Rands.Maximum
            End If
            If Hourly_Rate_Rands.Value < Hourly_Rate_Rands.Minimum Then
                Hourly_Rate_Rands.Value = Hourly_Rate_Rands.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Hourly_Rate_Cents_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Hourly_Rate_Cents.Leave
        Try
            If Hourly_Rate_Cents.Value > Hourly_Rate_Cents.Maximum Then
                Hourly_Rate_Cents.Value = Hourly_Rate_Cents.Maximum
            End If
            If Hourly_Rate_Cents.Value < Hourly_Rate_Cents.Minimum Then
                Hourly_Rate_Cents.Value = Hourly_Rate_Cents.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Dim objStreamReader = New System.IO.StreamReader("C:\Documents and Settings\Administrator\Desktop\hr106.doc", True)
            Dim objStreamWriter = New System.IO.StreamWriter("C:\Documents and Settings\Administrator\Desktop\new hr106.doc", False, System.Text.Encoding.UTF8)
            Dim line As String
            'Dim templine As String
            'Dim chop As Integer
            'chop = 1
            Do
                line = objStreamReader.ReadLine()
                If IsNothing(line) = False Then

                    line = line.Replace("""", """""")
                    line = "objStreamWriter.WriteLine(""" & line & """)"
                    objStreamWriter.WriteLine(line)


                End If

            Loop Until line Is Nothing
            MsgBox("Completed")
            objStreamReader.close()
            objStreamWriter.close()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Generate_Student_Timesheet()
    End Sub



    Private Sub Message_Handler(ByVal message As String)
        MsgBox(message, MsgBoxStyle.Information, "Lab Tutor Payment System")
    End Sub



    Private Function Validate_Report_Data() As Boolean
        Try
            Dim message As String
            message = ""
            
            If DateDiff(DateInterval.Day, Payment_Period_Start.Value, Payment_Period_End.Value) < 1 Then
                message = message & "Please select a Payment Period End date that is greater than the Payment Period Start date." & vbCrLf
            End If
            If message.Length > 0 Then
                Message_Handler(message)
                Return False
            Else

                Return True
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
            Return False
        End Try
    End Function

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Payment_Period_Start.Value = Now()
        Payment_Period_End.Value = Now()
    End Sub

    Private Function capitalize_first_letter(ByVal input As String) As String
        Try
            If input.Trim.Length > 0 Then
                input = input.ToLower
                Dim myarray As Array
                myarray = input.Split(" ")
                Dim i As Long
                Dim replace As String
                For i = 0 To myarray.Length - 1
                    If myarray(i).length > 0 Then
                        replace = myarray(i).Substring(0, 1).ToUpper
                        myarray(i) = myarray(i).Remove(0, 1)
                        myarray(i) = replace & myarray(i)
                    End If
                Next
                input = ""
                For i = 0 To myarray.Length - 1
                    input = input & myarray(i) & " "
                Next
                input = input.Trim()
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
        Return input
    End Function

    Private Sub Report_Compiler_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Report_Compiler.Leave
        Report_Compiler.Text = capitalize_first_letter(Report_Compiler.Text)
    End Sub

    Private Sub Generate_Multiple_Payments_Sheets()
        Try
            If Validate_Report_Data() = True Then


                Dim result As DialogResult
                result = FolderBrowserDialog1.ShowDialog
                If result = DialogResult.OK Then

                    Dim initial_overwrite_status As Boolean
                    initial_overwrite_status = Overwrite_Files.Checked
                    Dim filetowrite As String
                    Dim writefiles As Boolean
                    writefiles = True
                    Dim writtencounter As Long
                    writtencounter = 0
                    Dim foldername As String
                    If Create_Subfolder.Checked = False Then

                        foldername = FolderBrowserDialog1.SelectedPath
                    Else

                        Dim dinfo As New System.IO.DirectoryInfo(FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString)
                        If dinfo.Exists() = False Then
                            dinfo.Create()
                        End If

                        foldername = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString
                    End If

                    Dim i As Long


                    Dim objStreamWriter As StreamWriter

                    Dim formcount, runner As Long


                    runner = 1
                    formcount = 1

                    Dim difference As Double
                    Progress_Bar1_Panel.Visible = True
                    ProgressBar1.Value = 0
                    ProgressBar1.Maximum = ListView1.Items.Count

                    Dim listview1start As Long
                    listview1start = 0

                    Dim pagetotalhours, pagetotalrands, pagetotalcents, hr106totalhours, hr106totalrands, hr106totalcents As Double
                    pagetotalhours = 0
                    pagetotalrands = 0
                    pagetotalcents = 0
                    hr106totalhours = 0
                    hr106totalrands = 0
                    hr106totalcents = 0
                    Dim hr106totalpayme, pagetotalpayme As Double
                    hr106totalpayme = 0
                    pagetotalpayme = 0

                    While runner <= formcount

                        pagetotalhours = 0
                        pagetotalrands = 0
                        pagetotalcents = 0
                        pagetotalpayme = 0

                        If Create_Subfolder.Checked = False Then
                            filetowrite = FolderBrowserDialog1.SelectedPath & "\HR106_Form" & runner & ".doc"
                            foldername = FolderBrowserDialog1.SelectedPath
                        Else

                            Dim dinfo As New System.IO.DirectoryInfo(FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString)
                            If dinfo.Exists() = False Then
                                dinfo.Create()
                            End If
                            filetowrite = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString & "\HR106_Form" & runner & ".doc"
                            foldername = FolderBrowserDialog1.SelectedPath & "\" & Now().ToLongDateString
                        End If
                        If Overwrite_Files.Checked = False And writefiles = True Then

                            Dim finfo As New System.IO.FileInfo(filetowrite)
                            If finfo.Exists() = True Then
                                result = MsgBox("The file to be generated already exists at the currently selected location." & vbCrLf & "Do you wish to overwrite this and all other conflicting files?", MsgBoxStyle.YesNo, "Lab Tutor Payment System")
                                If result = DialogResult.Yes Or result = DialogResult.OK Then
                                    writefiles = True
                                    Overwrite_Files.Checked = True
                                Else
                                    writefiles = False
                                    Overwrite_Files.Checked = False
                                End If
                            End If
                        End If
                        If writefiles = True Then

                            Dim finfo2 As New System.IO.FileInfo(foldername & "\Files\image001.gif")
                            If finfo2.Exists = False Then
                                Dim dinfo2 = New System.IO.DirectoryInfo(finfo2.DirectoryName)
                                If dinfo2.Exists() = False Then
                                    dinfo2.Create()
                                End If
                                Dim f As FileInfo
                                f = New FileInfo(Application.ExecutablePath())
                                Dim finfo3 As New System.IO.FileInfo(f.DirectoryName & "\Forms\image001.gif")
                                finfo3 = finfo3.CopyTo(finfo2.FullName)
                                

                            End If
                            objStreamWriter = New System.IO.StreamWriter(filetowrite, False, System.Text.Encoding.UTF8)

                            '*******************************************************

                            objStreamWriter.WriteLine("<html xmlns:v=""urn:schemas-microsoft-com:vml""")
                            objStreamWriter.WriteLine("xmlns:o=""urn:schemas-microsoft-com:office:office""")
                            objStreamWriter.WriteLine("xmlns:w=""urn:schemas-microsoft-com:office:word""")
                            objStreamWriter.WriteLine("xmlns:st1=""urn:schemas-microsoft-com:office:smarttags""")
                            objStreamWriter.WriteLine("xmlns=""http://www.w3.org/TR/REC-html40"">")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<head>")
                            objStreamWriter.WriteLine("<meta http-equiv=Content-Type content=""text/html; charset=windows-1252"">")
                            objStreamWriter.WriteLine("<meta name=ProgId content=Word.Document>")
                            objStreamWriter.WriteLine("<meta name=Generator content=""Microsoft Word 11"">")
                            objStreamWriter.WriteLine("<meta name=Originator content=""Microsoft Word 11"">")
                            '   objStreamWriter.WriteLine("<link rel=File-List href=""Files\filelist.xml"">")
                            'objStreamWriter.WriteLine("<link rel=Edit-Time-Data href=""Files\editdata.mso"">")
                            objStreamWriter.WriteLine("<!--[if !mso]>")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine("v\:* {behavior:url(#default#VML);}")
                            objStreamWriter.WriteLine("o\:* {behavior:url(#default#VML);}")
                            objStreamWriter.WriteLine("w\:* {behavior:url(#default#VML);}")
                            objStreamWriter.WriteLine(".shape {behavior:url(#default#VML);}")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<![endif]-->")
                            objStreamWriter.WriteLine("<title>Multiple Payments</title>")
                            objStreamWriter.WriteLine("<o:SmartTagType namespaceuri=""urn:schemas-microsoft-com:office:smarttags""")
                            objStreamWriter.WriteLine(" name=""place""/>")
                            objStreamWriter.WriteLine("<!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <o:DocumentProperties>")
                            objStreamWriter.WriteLine("  <o:Author>Commerce I.T.</o:Author>")
                            objStreamWriter.WriteLine("  <o:LastAuthor>Commerce I.T.</o:LastAuthor>")
                            objStreamWriter.WriteLine("  <o:Revision>2</o:Revision>")
                            objStreamWriter.WriteLine("  <o:TotalTime>5</o:TotalTime>")
                            objStreamWriter.WriteLine("  <o:LastPrinted>2004-05-29T09:48:00Z</o:LastPrinted>")
                            objStreamWriter.WriteLine("  <o:Created>2004-09-27T12:36:00Z</o:Created>")
                            objStreamWriter.WriteLine("  <o:LastSaved>2004-09-27T12:36:00Z</o:LastSaved>")
                            objStreamWriter.WriteLine("  <o:Pages>1</o:Pages>")
                            objStreamWriter.WriteLine("  <o:Words>184</o:Words>")
                            objStreamWriter.WriteLine("  <o:Characters>1054</o:Characters>")
                            objStreamWriter.WriteLine("  <o:Company>University of Cape Town</o:Company>")
                            objStreamWriter.WriteLine("  <o:Lines>8</o:Lines>")
                            objStreamWriter.WriteLine("  <o:Paragraphs>2</o:Paragraphs>")
                            objStreamWriter.WriteLine("  <o:CharactersWithSpaces>1236</o:CharactersWithSpaces>")
                            objStreamWriter.WriteLine("  <o:Version>11.5606</o:Version>")
                            objStreamWriter.WriteLine(" </o:DocumentProperties>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <w:WordDocument>")
                            objStreamWriter.WriteLine("  <w:DoNotHyphenateCaps/>")
                            objStreamWriter.WriteLine("  <w:DrawingGridHorizontalSpacing>6 pt</w:DrawingGridHorizontalSpacing>")
                            objStreamWriter.WriteLine("  <w:DrawingGridVerticalSpacing>6 pt</w:DrawingGridVerticalSpacing>")
                            objStreamWriter.WriteLine("  <w:DisplayVerticalDrawingGridEvery>0</w:DisplayVerticalDrawingGridEvery>")
                            objStreamWriter.WriteLine("  <w:UseMarginsForDrawingGridOrigin/>")
                            objStreamWriter.WriteLine("  <w:ValidateAgainstSchemas/>")
                            objStreamWriter.WriteLine("  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>")
                            objStreamWriter.WriteLine("  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>")
                            objStreamWriter.WriteLine("  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>")
                            objStreamWriter.WriteLine("  <w:Compatibility>")
                            objStreamWriter.WriteLine("   <w:BalanceSingleByteDoubleByteWidth/>")
                            objStreamWriter.WriteLine("   <w:DoNotLeaveBackslashAlone/>")
                            objStreamWriter.WriteLine("   <w:ULTrailSpace/>")
                            objStreamWriter.WriteLine("   <w:DoNotExpandShiftReturn/>")
                            objStreamWriter.WriteLine("   <w:UsePrinterMetrics/>")
                            objStreamWriter.WriteLine("   <w:WW6BorderRules/>")
                            objStreamWriter.WriteLine("   <w:FootnoteLayoutLikeWW8/>")
                            objStreamWriter.WriteLine("   <w:ShapeLayoutLikeWW8/>")
                            objStreamWriter.WriteLine("   <w:AlignTablesRowByRow/>")
                            objStreamWriter.WriteLine("   <w:ForgetLastTabAlignment/>")
                            objStreamWriter.WriteLine("   <w:AutoSpaceLikeWord95/>")
                            objStreamWriter.WriteLine("   <w:LayoutRawTableWidth/>")
                            objStreamWriter.WriteLine("   <w:LayoutTableRowsApart/>")
                            objStreamWriter.WriteLine("   <w:UseWord97LineBreakingRules/>")
                            objStreamWriter.WriteLine("   <w:SelectEntireFieldWithStartOrEnd/>")
                            objStreamWriter.WriteLine("   <w:UseWord2002TableStyleRules/>")
                            objStreamWriter.WriteLine("  </w:Compatibility>")
                            objStreamWriter.WriteLine("  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>")
                            objStreamWriter.WriteLine(" </w:WordDocument>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <w:LatentStyles DefLockedState=""false"" LatentStyleCount=""156"">")
                            objStreamWriter.WriteLine(" </w:LatentStyles>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if !mso]><object")
                            objStreamWriter.WriteLine(" classid=""clsid:38481807-CA0E-42D2-BF39-B33AF135CC4D"" id=ieooui></object>")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine("st1\:*{behavior:url(#ieooui) }")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<![endif]-->")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine("<!--")
                            objStreamWriter.WriteLine(" /* Font Definitions */")
                            objStreamWriter.WriteLine(" @font-face")
                            objStreamWriter.WriteLine("	{font-family:Helvetica;")
                            objStreamWriter.WriteLine("	panose-1:2 11 6 4 2 2 2 2 2 4;")
                            objStreamWriter.WriteLine("	mso-font-charset:0;")
                            objStreamWriter.WriteLine("	mso-generic-font-family:swiss;")
                            objStreamWriter.WriteLine("	mso-font-pitch:variable;")
                            objStreamWriter.WriteLine("	mso-font-signature:536902279 -2147483648 8 0 511 0;}")
                            objStreamWriter.WriteLine("@font-face")
                            objStreamWriter.WriteLine("	{font-family:Wingdings;")
                            objStreamWriter.WriteLine("	panose-1:5 0 0 0 0 0 0 0 0 0;")
                            objStreamWriter.WriteLine("	mso-font-charset:2;")
                            objStreamWriter.WriteLine("	mso-generic-font-family:auto;")
                            objStreamWriter.WriteLine("	mso-font-pitch:variable;")
                            objStreamWriter.WriteLine("	mso-font-signature:0 268435456 0 0 -2147483648 0;}")
                            objStreamWriter.WriteLine("@font-face")
                            objStreamWriter.WriteLine("	{font-family:Tahoma;")
                            objStreamWriter.WriteLine("	panose-1:2 11 6 4 3 5 4 4 2 4;")
                            objStreamWriter.WriteLine("	mso-font-charset:0;")
                            objStreamWriter.WriteLine("	mso-generic-font-family:swiss;")
                            objStreamWriter.WriteLine("	mso-font-pitch:variable;")
                            objStreamWriter.WriteLine("	mso-font-signature:1627421319 -2147483648 8 0 66047 0;}")
                            objStreamWriter.WriteLine(" /* Style Definitions */")
                            objStreamWriter.WriteLine(" p.MsoNormal, li.MsoNormal, div.MsoNormal")
                            objStreamWriter.WriteLine("	{mso-style-parent:"""";")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:12.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("h1")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin-top:12.0pt;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:3.0pt;")
                            objStreamWriter.WriteLine("	margin-left:0cm;")
                            objStreamWriter.WriteLine("	text-align:justify;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:1;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:14.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-bidi-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-font-kerning:14.0pt;")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("h2")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:2;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:12.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("h3")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:3;")
                            objStreamWriter.WriteLine("	tab-stops:0cm;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Tahoma;")
                            objStreamWriter.WriteLine("	mso-bidi-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("h4")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:4;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:14.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-weight:normal;}")
                            objStreamWriter.WriteLine("h5")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:5;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:8.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Tahoma;")
                            objStreamWriter.WriteLine("	mso-bidi-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("h6")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:6;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:14.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoHeading7, li.MsoHeading7, div.MsoHeading7")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin-top:0cm;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:0cm;")
                            objStreamWriter.WriteLine("	margin-left:-21.3pt;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:7;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:11.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-weight:bold;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("p.MsoHeading8, li.MsoHeading8, div.MsoHeading8")
                            objStreamWriter.WriteLine("	{mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-outline-level:8;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-style:italic;}")
                            objStreamWriter.WriteLine("p.MsoFootnoteText, li.MsoFootnoteText, div.MsoFootnoteText")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoCommentText, li.MsoCommentText, div.MsoCommentText")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoHeader, li.MsoHeader, div.MsoHeader")
                            objStreamWriter.WriteLine("	{margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:justify;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	tab-stops:center 207.65pt right 415.3pt;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoFooter, li.MsoFooter, div.MsoFooter")
                            objStreamWriter.WriteLine("	{margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:justify;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	tab-stops:center 207.65pt right 415.3pt;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("span.MsoFootnoteReference")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	vertical-align:super;}")
                            objStreamWriter.WriteLine("span.MsoCommentReference")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	mso-ansi-font-size:8.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:8.0pt;}")
                            objStreamWriter.WriteLine("p.MsoSubtitle, li.MsoSubtitle, div.MsoSubtitle")
                            objStreamWriter.WriteLine("	{margin-top:0cm;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:3.0pt;")
                            objStreamWriter.WriteLine("	margin-left:0cm;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	font-size:12.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-bidi-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.MsoBodyText2, li.MsoBodyText2, div.MsoBodyText2")
                            objStreamWriter.WriteLine("	{margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	text-align:justify;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("a:link, span.MsoHyperlink")
                            objStreamWriter.WriteLine("	{color:blue;")
                            objStreamWriter.WriteLine("	text-decoration:underline;")
                            objStreamWriter.WriteLine("	text-underline:single;}")
                            objStreamWriter.WriteLine("a:visited, span.MsoHyperlinkFollowed")
                            objStreamWriter.WriteLine("	{color:purple;")
                            objStreamWriter.WriteLine("	text-decoration:underline;")
                            objStreamWriter.WriteLine("	text-underline:single;}")
                            objStreamWriter.WriteLine("p.MsoCommentSubject, li.MsoCommentSubject, div.MsoCommentSubject")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	mso-style-parent:""Comment Text"";")
                            objStreamWriter.WriteLine("	mso-style-next:""Comment Text"";")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-weight:bold;}")
                            objStreamWriter.WriteLine("p.MsoAcetate, li.MsoAcetate, div.MsoAcetate")
                            objStreamWriter.WriteLine("	{mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:8.0pt;")
                            objStreamWriter.WriteLine("	font-family:Tahoma;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.Formtitle, li.Formtitle, div.Formtitle")
                            objStreamWriter.WriteLine("	{mso-style-name:""Form title"";")
                            objStreamWriter.WriteLine("	mso-style-parent:""Heading 1"";")
                            objStreamWriter.WriteLine("	mso-style-next:Normal;")
                            objStreamWriter.WriteLine("	margin-top:0cm;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:18.0pt;")
                            objStreamWriter.WriteLine("	margin-left:0cm;")
                            objStreamWriter.WriteLine("	text-align:center;")
                            objStreamWriter.WriteLine("	mso-pagination:none;")
                            objStreamWriter.WriteLine("	page-break-after:avoid;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:16.0pt;")
                            objStreamWriter.WriteLine("	mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:Arial;")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-bidi-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-font-kerning:14.0pt;")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;")
                            objStreamWriter.WriteLine("	font-weight:bold;")
                            objStreamWriter.WriteLine("	mso-bidi-font-weight:normal;}")
                            objStreamWriter.WriteLine("p.TipNoteText, li.TipNoteText, div.TipNoteText")
                            objStreamWriter.WriteLine("	{mso-style-name:""Tip\/Note Text"";")
                            objStreamWriter.WriteLine("	margin-top:0cm;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:0cm;")
                            objStreamWriter.WriteLine("	margin-left:106.3pt;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	tab-stops:14.2pt;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.TopicTextBulleted, li.TopicTextBulleted, div.TopicTextBulleted")
                            objStreamWriter.WriteLine("	{mso-style-name:""Topic Text Bulleted"";")
                            objStreamWriter.WriteLine("	margin-top:0cm;")
                            objStreamWriter.WriteLine("	margin-right:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:6.0pt;")
                            objStreamWriter.WriteLine("	margin-left:3.0cm;")
                            objStreamWriter.WriteLine("	text-indent:-14.15pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	tab-stops:3.0cm;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine("p.TableText, li.TableText, div.TableText")
                            objStreamWriter.WriteLine("	{mso-style-name:""Table Text"";")
                            objStreamWriter.WriteLine("	margin:0cm;")
                            objStreamWriter.WriteLine("	margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	mso-layout-grid-align:none;")
                            objStreamWriter.WriteLine("	punctuation-wrap:simple;")
                            objStreamWriter.WriteLine("	text-autospace:none;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-fareast-language:EN-US;}")
                            objStreamWriter.WriteLine(" /* Page Definitions */")
                            objStreamWriter.WriteLine(" @page")
                            'objStreamWriter.WriteLine("	{mso-footnote-separator:url(""Files/Header.htm"") fs;")
                            'objStreamWriter.WriteLine("	mso-footnote-continuation-separator:url(""Files/Header.htm"") fcs;")
                            'objStreamWriter.WriteLine("	mso-endnote-separator:url(""Files/Header.htm"") es;")
                            'objStreamWriter.WriteLine("	mso-endnote-continuation-separator:url(""Files/Header.htm"") ecs;}")
                            objStreamWriter.WriteLine("@page Section1")
                            objStreamWriter.WriteLine("	{size:841.7pt 595.45pt;")
                            objStreamWriter.WriteLine("	mso-page-orientation:landscape;")
                            objStreamWriter.WriteLine("	margin:36.0pt 2.0cm 36.0pt 2.0cm;")
                            objStreamWriter.WriteLine("	mso-header-margin:28.05pt;")
                            objStreamWriter.WriteLine("	mso-footer-margin:28.05pt;")
                            'objStreamWriter.WriteLine("	mso-footer:url(""Files/Header.htm"") f1;")
                            objStreamWriter.WriteLine("	mso-paper-source:0;}")
                            objStreamWriter.WriteLine("div.Section1")
                            objStreamWriter.WriteLine("	{page:Section1;}")
                            objStreamWriter.WriteLine(" /* List Definitions */")
                            objStreamWriter.WriteLine(" @list l0")
                            objStreamWriter.WriteLine("	{mso-list-id:126051447;")
                            objStreamWriter.WriteLine("	mso-list-type:hybrid;")
                            objStreamWriter.WriteLine("	mso-list-template-ids:295347656 67698703 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}")
                            objStreamWriter.WriteLine("@list l0:level1")
                            objStreamWriter.WriteLine("	{mso-level-tab-stop:36.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;}")
                            objStreamWriter.WriteLine("@list l1")
                            objStreamWriter.WriteLine("	{mso-list-id:175310959;")
                            objStreamWriter.WriteLine("	mso-list-type:hybrid;")
                            objStreamWriter.WriteLine("	mso-list-template-ids:1401484124 67698689 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}")
                            objStreamWriter.WriteLine("@list l1:level1")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0B7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:36.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:Symbol;}")
                            objStreamWriter.WriteLine("@list l2")
                            objStreamWriter.WriteLine("	{mso-list-id:702249733;")
                            objStreamWriter.WriteLine("	mso-list-type:hybrid;")
                            objStreamWriter.WriteLine("	mso-list-template-ids:-1829493478 67698689 67698691 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}")
                            objStreamWriter.WriteLine("@list l2:level1")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0B7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:36.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:Symbol;}")
                            objStreamWriter.WriteLine("@list l3")
                            objStreamWriter.WriteLine("	{mso-list-id:975988893;")
                            objStreamWriter.WriteLine("	mso-list-type:hybrid;")
                            objStreamWriter.WriteLine("	mso-list-template-ids:1193193544 -1306904382 -639568170 67698693 67698689 67698691 67698693 67698689 67698691 67698693;}")
                            objStreamWriter.WriteLine("@list l3:level1")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0B7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:36.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-21.8pt;")
                            objStreamWriter.WriteLine("	font-family:Symbol;}")
                            objStreamWriter.WriteLine("@list l3:level2")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0B7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:72.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:Symbol;}")
                            objStreamWriter.WriteLine("@list l4")
                            objStreamWriter.WriteLine("	{mso-list-id:2141998080;")
                            objStreamWriter.WriteLine("	mso-list-type:hybrid;")
                            objStreamWriter.WriteLine("	mso-list-template-ids:-651806402 -639568170 470351875 470351877 470351873 470351875 470351877 470351873 470351875 470351877;}")
                            objStreamWriter.WriteLine("@list l4:level1")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0B7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:107.85pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	margin-left:107.85pt;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:Symbol;}")
                            objStreamWriter.WriteLine("@list l4:level2")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:o;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:72.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Courier New"";}")
                            objStreamWriter.WriteLine("@list l4:level3")
                            objStreamWriter.WriteLine("	{mso-level-number-format:bullet;")
                            objStreamWriter.WriteLine("	mso-level-text:\F0A7;")
                            objStreamWriter.WriteLine("	mso-level-tab-stop:108.0pt;")
                            objStreamWriter.WriteLine("	mso-level-number-position:left;")
                            objStreamWriter.WriteLine("	text-indent:-18.0pt;")
                            objStreamWriter.WriteLine("	font-family:Wingdings;}")
                            objStreamWriter.WriteLine("ol")
                            objStreamWriter.WriteLine("	{margin-bottom:0cm;}")
                            objStreamWriter.WriteLine("ul")
                            objStreamWriter.WriteLine("	{margin-bottom:0cm;}")
                            objStreamWriter.WriteLine("-->")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<!--[if gte mso 10]>")
                            objStreamWriter.WriteLine("<style>")
                            objStreamWriter.WriteLine(" /* Style Definitions */")
                            objStreamWriter.WriteLine(" table.MsoNormalTable")
                            objStreamWriter.WriteLine("	{mso-style-name:""Table Normal"";")
                            objStreamWriter.WriteLine("	mso-tstyle-rowband-size:0;")
                            objStreamWriter.WriteLine("	mso-tstyle-colband-size:0;")
                            objStreamWriter.WriteLine("	mso-style-noshow:yes;")
                            objStreamWriter.WriteLine("	mso-style-parent:"""";")
                            objStreamWriter.WriteLine("	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("	mso-para-margin:0cm;")
                            objStreamWriter.WriteLine("	mso-para-margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("	mso-pagination:widow-orphan;")
                            objStreamWriter.WriteLine("	font-size:10.0pt;")
                            objStreamWriter.WriteLine("	font-family:""Times New Roman"";")
                            objStreamWriter.WriteLine("	mso-ansi-language:#0400;")
                            objStreamWriter.WriteLine("	mso-fareast-language:#0400;")
                            objStreamWriter.WriteLine("	mso-bidi-language:#0400;}")
                            objStreamWriter.WriteLine("</style>")
                            objStreamWriter.WriteLine("<![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <o:shapedefaults v:ext=""edit"" spidmax=""2050""/>")
                            objStreamWriter.WriteLine("</xml><![endif]--><!--[if gte mso 9]><xml>")
                            objStreamWriter.WriteLine(" <o:shapelayout v:ext=""edit"">")
                            objStreamWriter.WriteLine("  <o:idmap v:ext=""edit"" data=""1""/>")
                            objStreamWriter.WriteLine(" </o:shapelayout></xml><![endif]-->")
                            objStreamWriter.WriteLine("</head>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<body lang=EN-GB link=blue vlink=purple style='tab-interval:36.0pt'>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<div class=Section1>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0")
                            objStreamWriter.WriteLine(" style='border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;")
                            objStreamWriter.WriteLine(" mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:.5pt solid windowtext;")
                            objStreamWriter.WriteLine(" mso-border-insidev:.5pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;")
                            objStreamWriter.WriteLine("  page-break-inside:avoid'>")
                            objStreamWriter.WriteLine("  <td width=272 valign=top style='width:203.8pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><!--[if gte vml 1]><v:shapetype")
                            objStreamWriter.WriteLine("   id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t""")
                            objStreamWriter.WriteLine("   path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">")
                            objStreamWriter.WriteLine("   <v:stroke joinstyle=""miter""/>")
                            objStreamWriter.WriteLine("   <v:formulas>")
                            objStreamWriter.WriteLine("    <v:f eqn=""if lineDrawn pixelLineWidth 0""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""sum @0 1 0""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""sum 0 0 @1""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @2 1 2""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @3 21600 pixelWidth""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @3 21600 pixelHeight""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""sum @0 0 1""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @6 1 2""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @7 21600 pixelWidth""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""sum @8 21600 0""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""prod @7 21600 pixelHeight""/>")
                            objStreamWriter.WriteLine("    <v:f eqn=""sum @10 21600 0""/>")
                            objStreamWriter.WriteLine("   </v:formulas>")
                            objStreamWriter.WriteLine("   <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>")
                            objStreamWriter.WriteLine("   <o:lock v:ext=""edit"" aspectratio=""t""/>")
                            objStreamWriter.WriteLine("  </v:shapetype><v:shape id=""_x0000_i1025"" type=""#_x0000_t75"" style='width:46.5pt;")
                            objStreamWriter.WriteLine("   height:60pt'>")
                            objStreamWriter.WriteLine("   <v:imagedata src=""Files/image001.gif"" o:title=""UCTlogo-notext""/>")
                            objStreamWriter.WriteLine("  </v:shape><![endif]--><![if !vml]><img width=62 height=80")
                            objStreamWriter.WriteLine("  src=""Files/image001.gif"" v:shapes=""_x0000_i1025""><![endif]><span")
                            objStreamWriter.WriteLine("  style='font-size:14.0pt;mso-bidi-font-size:10.0pt'><o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=448 valign=top style='width:335.85pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='margin-bottom:6.0pt;text-align:center'><b><span")
                            objStreamWriter.WriteLine("  style='font-size:14.0pt;mso-bidi-font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></b></p>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><b><span")
                            objStreamWriter.WriteLine("  style='font-size:14.0pt;mso-bidi-font-size:10.0pt;font-family:Arial'>MULTIPLE")
                            objStreamWriter.WriteLine("  PAYMENTS<o:p></o:p></span></b></p>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:14.0pt;mso-bidi-font-size:10.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=252 valign=top style='width:188.7pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='margin-bottom:6.0pt;text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:14.0pt;mso-bidi-font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:14.0pt;mso-bidi-font-size:10.0pt;font-family:Arial'>HR106<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=Formtitle align=left style='margin-bottom:0cm;margin-bottom:.0001pt;")
                            objStreamWriter.WriteLine("text-align:left'><span style='font-size:8.0pt;mso-bidi-font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial'>NOTES<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l1 level1 lfo3;")
                            objStreamWriter.WriteLine("tab-stops:0cm list 36.0pt'><![if !supportLists]><span style='font-size:8.0pt;")
                            objStreamWriter.WriteLine("font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol'><span")
                            objStreamWriter.WriteLine("style='mso-list:Ignore'><span style='font:7.0pt ""Times New Roman""'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                            objStreamWriter.WriteLine("</span></span></span><![endif]><span style='font-size:8.0pt;font-family:Arial'>Forms")
                            objStreamWriter.WriteLine("must be downloaded from the UCT website: <a")
                            objStreamWriter.WriteLine("href=""http://www.uct.ac.za/depts/sapweb/forms/forms.htm"">http//www.uct.ac.za/depts/sapweb/forms/forms.htm</a><span")
                            objStreamWriter.WriteLine("style='mso-spacerun:yes'> </span><o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l1 level1 lfo3;")
                            objStreamWriter.WriteLine("tab-stops:0cm list 36.0pt'><![if !supportLists]><span style='font-size:8.0pt;")
                            objStreamWriter.WriteLine("font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol'><span")
                            objStreamWriter.WriteLine("style='mso-list:Ignore'><span style='font:7.0pt ""Times New Roman""'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                            objStreamWriter.WriteLine("</span></span></span><![endif]><span style='font-size:8.0pt;font-family:Arial'>This")
                            objStreamWriter.WriteLine("form is used when making variable payments to staff who have already been")
                            objStreamWriter.WriteLine("appointed on a paid-on-claim contract via form HR100.<span")
                            objStreamWriter.WriteLine("style='mso-spacerun:yes'> </span>Such payments are made on payday at month end")
                            objStreamWriter.WriteLine("(ie. on 25<sup>th</sup> of the month).<span style='mso-spacerun:yes'> </span><o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l1 level1 lfo3;")
                            objStreamWriter.WriteLine("tab-stops:0cm list 36.0pt'><![if !supportLists]><span style='font-size:8.0pt;")
                            objStreamWriter.WriteLine("font-family:Symbol;mso-fareast-font-family:Symbol;mso-bidi-font-family:Symbol'><span")
                            objStreamWriter.WriteLine("style='mso-list:Ignore'><span style='font:7.0pt ""Times New Roman""'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                            objStreamWriter.WriteLine("</span></span></span><![endif]><span style='font-size:8.0pt;font-family:Arial'>It")
                            objStreamWriter.WriteLine("may also be used to make a leave payout to paid-on-claim staff.<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0")
                            objStreamWriter.WriteLine(" style='background:#E6E6E6;border-collapse:collapse;border:none;mso-border-alt:")
                            objStreamWriter.WriteLine(" solid windowtext .75pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:")
                            objStreamWriter.WriteLine(" .75pt solid windowtext;mso-border-insidev:.75pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;")
                            objStreamWriter.WriteLine("  page-break-inside:avoid'>")
                            objStreamWriter.WriteLine("  <td width=139 style='width:104.4pt;border:solid windowtext 1.0pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <h2><span style='font-size:9.0pt;font-family:Arial;font-weight:normal;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>Submission Date<o:p></o:p></span></h2>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.9pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <h2><span style='font-size:9.0pt;font-family:Arial'>By 3<sup>rd</sup> of the")
                            objStreamWriter.WriteLine("  month<s><o:p></o:p></s></span></h2>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.9pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>Additional Forms<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.9pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial'>Preceded")
                            objStreamWriter.WriteLine("  by HR100<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.9pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <h2><span style='font-size:9.0pt;font-family:Arial;font-weight:normal;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>Help Document<o:p></o:p></span></h2>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=155 style='width:116.35pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial'>See page 2<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("normal'><span style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></b></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("normal'><span style='font-size:9.0pt;font-family:Arial'>EVENT</span></b><span")
                            objStreamWriter.WriteLine("style='font-size:9.0pt;font-family:Arial'><o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0")
                            objStreamWriter.WriteLine(" style='margin-left:-1.7pt;border-collapse:collapse;border:none;mso-border-alt:")
                            objStreamWriter.WriteLine(" solid windowtext .75pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:")
                            objStreamWriter.WriteLine(" .75pt solid windowtext;mso-border-insidev:.75pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=142 style='width:106.35pt;border:solid windowtext 1.0pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-bottom-alt:solid windowtext .5pt;")
                            objStreamWriter.WriteLine("  background:#E5E5E5;mso-shading:white;mso-pattern:gray-10 auto;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'>Event Type <i>(tick)<o:p></o:p></i></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial'>Multiple")
                            objStreamWriter.WriteLine("  Payments<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>&nbsp;<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold;mso-bidi-font-style:italic'>Department<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=324 colspan=2 style='width:243.15pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoHeading8><span style='font-size:9.0pt;font-style:normal;")
                            objStreamWriter.WriteLine("  mso-bidi-font-style:italic'><o:p>" & Department.Text & "</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;")
                            objStreamWriter.WriteLine("  page-break-inside:avoid;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=142 style='width:106.35pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-bottom-alt:solid windowtext .5pt;background:#E5E5E5;mso-shading:")
                            objStreamWriter.WriteLine("  white;mso-pattern:gray-10 auto;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'>Period of Payment<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>From (<i style='mso-bidi-font-style:normal'>ddmmyyyy</i>)<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial'><o:p>")

                            Dim datestring As String
                            If Payment_Period_Start.Value.Day < 10 Then
                                datestring = "0" & Payment_Period_Start.Value.Day
                            Else
                                datestring = Payment_Period_Start.Value.Day
                            End If
                            If Payment_Period_Start.Value.Month < 10 Then
                                datestring = datestring & "/" & "0" & Payment_Period_Start.Value.Month
                            Else
                                datestring = datestring & "/" & Payment_Period_Start.Value.Month
                            End If
                            datestring = datestring & "/" & Payment_Period_Start.Value.Year

                            objStreamWriter.WriteLine(datestring)
                            objStreamWriter.WriteLine("</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>To (<i style='mso-bidi-font-style:normal'>ddmmyyy</i>)<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'><o:p>")
                            If Payment_Period_End.Value.Day < 10 Then
                                datestring = "0" & Payment_Period_End.Value.Day
                            Else
                                datestring = Payment_Period_End.Value.Day
                            End If
                            If Payment_Period_End.Value.Month < 10 Then
                                datestring = datestring & "/" & "0" & Payment_Period_End.Value.Month
                            Else
                                datestring = datestring & "/" & Payment_Period_End.Value.Month
                            End If
                            datestring = datestring & "/" & Payment_Period_End.Value.Year
                            objStreamWriter.WriteLine(datestring)
                            objStreamWriter.WriteLine("</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoHeading8><span style='font-size:9.0pt'>(Limited to 1 month)<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal align=center style='text-align:center'><b><span")
                            objStreamWriter.WriteLine("style='font-size:9.0pt;font-family:Arial'>PAYMENT DETAILS<o:p></o:p></span></b></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0")
                            objStreamWriter.WriteLine(" style='margin-left:-1.7pt;border-collapse:collapse;border:none;mso-border-alt:")
                            objStreamWriter.WriteLine(" solid windowtext .75pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:")
                            objStreamWriter.WriteLine(" .75pt solid windowtext;mso-border-insidev:.75pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=95 rowspan=3 valign=top style='width:70.9pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E5E5E5;mso-shading:white;")
                            objStreamWriter.WriteLine("  mso-pattern:gray-10 auto;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center;tab-stops:36.0pt'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;font-weight:normal;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>Staff<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial'>Number<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><i")
                            objStreamWriter.WriteLine("  style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;font-family:")
                            objStreamWriter.WriteLine("  Arial'>(not Student Number)<o:p></o:p></span></i></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=216 rowspan=3 valign=top style='width:162.3pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E5E5E5;mso-shading:white;mso-pattern:gray-10 auto;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='margin-top:3.0pt;text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;font-weight:normal'>Surname and")
                            objStreamWriter.WriteLine("  First Name<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 rowspan=3 valign=top style='width:63.4pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='margin-top:3.0pt;text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;font-weight:normal'>Cost Centre<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 rowspan=3 valign=top style='width:63.45pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='margin-top:3.0pt;text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;font-weight:normal'>Fund<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=493 colspan=7 style='width:370.0pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  border-left:none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Multiple")
                            objStreamWriter.WriteLine("  Payments / Leave Payout<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=42 rowspan=2 style='width:31.7pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>MP or LP?<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=42 rowspan=2 style='width:31.7pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>To Fees?<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Hours<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 colspan=2 style='width:121.55pt;border-top:none;border-left:")
                            objStreamWriter.WriteLine("  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Hourly Rate<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 colspan=2 style='width:121.6pt;border-top:none;border-left:")
                            objStreamWriter.WriteLine("  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Total Amount<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.75pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><st1:place w:st=""on""><span")
                            objStreamWriter.WriteLine("   style='font-size:9.0pt;font-family:Arial;font-weight:normal;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("   bold'>Rands</span></st1:place><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Cents<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><st1:place w:st=""on""><span")
                            objStreamWriter.WriteLine("   style='font-size:9.0pt;font-family:Arial;font-weight:normal;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("   bold'>Rands</span></st1:place><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial;font-weight:normal;mso-bidi-font-weight:bold'>Cents<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            Dim ill As Long
                            Dim datafoundcount As Long
                            datafoundcount = 0

                            For ill = listview1start To ListView1.Items.Count - 1
                                Dim add_student As New Student_Details

                                add_student.First_Name = ListView1.Items(ill).SubItems(2).Text
                                add_student.Surname = ListView1.Items(ill).SubItems(1).Text
                                add_student.Staff_Number = ListView1.Items(ill).SubItems(3).Text
                                add_student.Student_Number = ListView1.Items(ill).Text

                                Dim counter1 As Long
                                Dim datafound As Boolean
                                Dim hours, hours_total As Double
                                hours = 0
                                hours_total = 0
                                datafound = False

                                Dim start_date, end_date As Date
                                start_date = Payment_Period_Start.Value
                                end_date = Payment_Period_End.Value
                                start_date = New Date(Year(start_date), Month(start_date), Day(start_date))
                                end_date = New Date(Year(end_date), Month(end_date), Day(end_date))


                                For counter1 = 0 To ListView2.Items.Count - 1
                                    hours = 0
                                    If (ListView2.Items(counter1).SubItems(1).Text = add_student.Student_Number) Then

                                        Dim add_event As New Event_Details

                                        add_event.Event_Date = ListView2.Items(counter1).Text
                                        add_event.Student_Number = ListView2.Items(counter1).SubItems(1).Text
                                        add_event.Start_Time = ListView2.Items(counter1).SubItems(2).Text
                                        add_event.End_Time = ListView2.Items(counter1).SubItems(3).Text
                                        add_event.Computer_Lab = ListView2.Items(counter1).SubItems(4).Text

                                        If (DateDiff(DateInterval.Day, start_date, add_event.Event_Date) >= 0) And (DateDiff(DateInterval.Day, add_event.Event_Date, end_date) >= 0) Then

                                            datafound = True

                                        Dim starttime, endtime As Date
                                        Dim myarray As Array
                                        myarray = add_event.Start_Time.Split(":")
                                        starttime = New Date(2000, 1, 1, myarray(0), myarray(1), 0)
                                        myarray = add_event.End_Time.Split(":")
                                        endtime = New Date(2000, 1, 1, myarray(0), myarray(1), 0)
                                        hours = hours + Math.Round((DateDiff(DateInterval.Minute, starttime, endtime) / 60), 1)

                                        hours_total = hours_total + hours

                                        End If
                                    End If
                                Next

                                If datafound = True Then

                                    datafoundcount = datafoundcount + 1


                                    objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                                    objStreamWriter.WriteLine("  height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <td width=95 style='width:70.9pt;border:solid windowtext 1.0pt;border-top:")
                                    objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                                    objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p>" & add_student.Staff_Number & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=216 style='width:162.3pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>" & add_student.Surname & ", " & add_student.First_Name & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.4pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>" & Cost_Centre.Text & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>" & Fund.Text & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>" & hours_total & "</o:p></span></h3>")
                                    pagetotalhours = pagetotalhours + hours_total
                                    hr106totalhours = hr106totalhours + hours_total
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.75pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>" & Hourly_Rate_Rands.Value & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>" & Hourly_Rate_Cents.Value & "</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>")
                                    Dim payme As Double
                                    payme = Math.Round((((Hourly_Rate_Rands.Value * 100) + Hourly_Rate_Cents.Value) * hours_total) / 100, 2)
                                    pagetotalpayme = pagetotalpayme + payme
                                    hr106totalpayme = hr106totalpayme + payme

                                    Dim splitarray As Array
                                    splitarray = payme.ToString.Split(".")
                                    If splitarray.Length > 1 Then


                                        If splitarray(1).length < 2 Then
                                            If splitarray(1).length > 0 Then
                                                splitarray(1) = splitarray(1) & "0"
                                            Else
                                                splitarray(1) = splitarray(1) & "00"
                                            End If
                                        End If

                                    End If
                                    objStreamWriter.WriteLine(splitarray(0))
                                    objStreamWriter.WriteLine("</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>")
                                    If splitarray.Length > 1 Then
                                        objStreamWriter.WriteLine(splitarray(1))
                                    Else
                                        objStreamWriter.WriteLine("00")
                                    End If
                                    objStreamWriter.WriteLine("</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine(" </tr>")
                                End If
                                ProgressBar1.Increment(1)
                                If datafoundcount = 7 Then
                                    listview1start = ill + 1
                                    formcount = formcount + 1
                                    Exit For
                                End If
                            Next
                            If datafoundcount < 7 Then
                                Dim loopcontrol As Long

                                For loopcontrol = 1 To (6 - datafoundcount)
                                    objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                                    objStreamWriter.WriteLine("  height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <td width=95 style='width:70.9pt;border:solid windowtext 1.0pt;border-top:")
                                    objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                                    objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=216 style='width:162.3pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.4pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.75pt;border-top:none;border-left:none;")
                                    objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                    objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                    objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                    objStreamWriter.WriteLine("  15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                    objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                    objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                    objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></h3>")
                                    objStreamWriter.WriteLine("  </td>")
                                    objStreamWriter.WriteLine(" </tr>")
                                Next

                            End If

                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=95 style='width:70.9pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><i style='mso-bidi-font-style:normal'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=216 style='width:162.3pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'>PAGE TOTAL<o:p></o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 style='width:63.4pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>" & pagetotalhours & "</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.75pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")

                            

                            Dim splitarray2 As Array
                            splitarray2 = pagetotalpayme.ToString.Split(".")
                            If splitarray2.Length > 1 Then


                                If splitarray2(1).length < 2 Then
                                    If splitarray2(1).length > 0 Then
                                        splitarray2(1) = splitarray2(1) & "0"
                                    Else
                                        splitarray2(1) = splitarray2(1) & "00"
                                    End If
                                End If

                            End If

                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>" & splitarray2(0) & "</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                            objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                            objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                            objStreamWriter.WriteLine("  font-family:Arial'><o:p>")
                            If splitarray2.Length > 1 Then
                                objStreamWriter.WriteLine(splitarray2(1))
                            Else
                                objStreamWriter.WriteLine("00")
                            End If

                            objStreamWriter.WriteLine("</o:p></span></i></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            If datafoundcount < 7 Then
                                objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;")
                                objStreamWriter.WriteLine("  page-break-inside:avoid;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <td width=95 style='width:70.9pt;border:solid windowtext 1.0pt;border-top:")
                                objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                                objStreamWriter.WriteLine("  padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><i style='mso-bidi-font-style:normal'><span")
                                objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=216 style='width:162.3pt;border-top:none;border-left:none;")
                                objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                objStreamWriter.WriteLine("  15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'>TOTAL<o:p></o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=85 style='width:63.4pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                objStreamWriter.WriteLine("  15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=42 style='width:31.7pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=85 style='width:63.45pt;border-top:none;border-left:none;")
                                objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                objStreamWriter.WriteLine("  15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>" & hr106totalhours & "</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=81 style='width:60.75pt;border-top:none;border-left:none;")
                                objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                                objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                                objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                                objStreamWriter.WriteLine("  15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>&nbsp;</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")

                                Dim splitarray3 As Array
                                splitarray3 = hr106totalpayme.ToString.Split(".")
                                If splitarray3.Length > 1 Then


                                    If splitarray3(1).length < 2 Then
                                        If splitarray3(1).length > 0 Then
                                            splitarray3(1) = splitarray3(1) & "0"
                                        Else
                                            splitarray3(1) = splitarray3(1) & "00"
                                        End If
                                    End If

                                End If

                                objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>" & splitarray3(0) & "</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine("  <td width=81 style='width:60.8pt;border-top:none;border-left:none;border-bottom:")
                                objStreamWriter.WriteLine("  solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;mso-border-top-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:")
                                objStreamWriter.WriteLine("  solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                                objStreamWriter.WriteLine("  <h3><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;")
                                objStreamWriter.WriteLine("  font-family:Arial'><o:p>")
                                If splitarray3.Length > 1 Then
                                    objStreamWriter.WriteLine(splitarray3(1))
                                Else
                                    objStreamWriter.WriteLine("00")
                                End If
                                objStreamWriter.WriteLine("</o:p></span></i></h3>")
                                objStreamWriter.WriteLine("  </td>")
                                objStreamWriter.WriteLine(" </tr>")
                            End If
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<h3 align=center style='text-align:center;tab-stops:36.0pt'><span")
                            objStreamWriter.WriteLine("style='font-size:9.0pt;font-family:Arial'>AUTHORITY FOR APPOINTMENT AND FUNDS<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0")
                            objStreamWriter.WriteLine(" style='margin-left:-1.7pt;border-collapse:collapse;border:none;mso-border-alt:")
                            objStreamWriter.WriteLine(" solid windowtext .75pt;mso-padding-alt:0cm 5.4pt 0cm 5.4pt;mso-border-insideh:")
                            objStreamWriter.WriteLine(" .75pt solid windowtext;mso-border-insidev:.75pt solid windowtext'>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=311 valign=top style='width:233.2pt;border:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3><span style='font-size:9.0pt;font-family:Arial;font-weight:normal'><o:p>&nbsp;</o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 align=center style='text-align:center;tab-stops:36.0pt'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;font-weight:normal;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>PRINT NAME<o:p></o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:bold'>SIGNATURE<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:bold'>CONTACT")
                            objStreamWriter.WriteLine("  NUMBER<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border:solid windowtext 1.0pt;border-left:")
                            objStreamWriter.WriteLine("  none;mso-border-left-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal align=center style='text-align:center'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:bold'>DATE<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=311 style='width:233.2pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>Administrator / Fund Holder<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=311 style='width:233.2pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoFootnoteText><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  mso-bidi-font-weight:bold'>Head of Department<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;page-break-inside:avoid;")
                            objStreamWriter.WriteLine("  height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=311 style='width:233.2pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>Area Finance Manager<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span")
                            objStreamWriter.WriteLine("  style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></b></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine(" <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes;")
                            objStreamWriter.WriteLine("  page-break-inside:avoid;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <td width=311 style='width:233.2pt;border:solid windowtext 1.0pt;border-top:")
                            objStreamWriter.WriteLine("  none;mso-border-top-alt:solid windowtext .75pt;mso-border-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  background:#E6E6E6;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'>HR Administrator<o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <h3 style='tab-stops:36.0pt'><span style='font-size:9.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("  font-weight:normal;mso-bidi-font-weight:bold'><o:p>&nbsp;</o:p></span></h3>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=169 style='width:126.85pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.55pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine("  <td width=162 style='width:121.6pt;border-top:none;border-left:none;")
                            objStreamWriter.WriteLine("  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("  mso-border-top-alt:solid windowtext .75pt;mso-border-left-alt:solid windowtext .75pt;")
                            objStreamWriter.WriteLine("  mso-border-alt:solid windowtext .75pt;padding:0cm 5.4pt 0cm 5.4pt;height:")
                            objStreamWriter.WriteLine("  15.0pt'>")
                            objStreamWriter.WriteLine("  <p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial;mso-bidi-font-weight:")
                            objStreamWriter.WriteLine("  bold'><o:p>&nbsp;</o:p></span></p>")
                            objStreamWriter.WriteLine("  </td>")
                            objStreamWriter.WriteLine(" </tr>")
                            objStreamWriter.WriteLine("</table>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("<p class=MsoNormal><o:p>&nbsp;</o:p></p>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("</div>")

                            objStreamWriter.WriteLine("<div style='mso-element:footer' id=f1>")
                            objStreamWriter.WriteLine("<div style='mso-element:para-border-div;border:none;border-top:solid windowtext 1.0pt;")
                            objStreamWriter.WriteLine("mso-border-top-alt:solid windowtext .75pt;padding:1.0pt 0cm 0cm 0cm;margin-left:")
                            objStreamWriter.WriteLine("-4.5pt;margin-right:7.1pt'>")
                            objStreamWriter.WriteLine("<p class=MsoFooter style='tab-stops:center 13.0cm right 722.95pt;border:none;")
                            objStreamWriter.WriteLine("mso-border-top-alt:solid windowtext .75pt;padding:0cm;mso-padding-alt:1.0pt 0cm 0cm 0cm'><!--[if supportFields]><span")
                            objStreamWriter.WriteLine("style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;mso-bidi-font-family:")
                            objStreamWriter.WriteLine("""Times New Roman""'><span style='mso-element:field-begin'></span><span")
                            objStreamWriter.WriteLine("style='mso-spacerun:yes'> </span>DATE \@ &quot;dd/MM/yyyy&quot; <span")
                            objStreamWriter.WriteLine("style='mso-element:field-separator'></span></span><![endif]--><span")
                            objStreamWriter.WriteLine("style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;mso-bidi-font-family:")
                            objStreamWriter.WriteLine("""Times New Roman""'><span style='mso-no-proof:yes'>27/09/2004</span></span><!--[if supportFields]><span")
                            objStreamWriter.WriteLine("style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;mso-bidi-font-family:")
                            objStreamWriter.WriteLine("""Times New Roman""'><span style='mso-element:field-end'></span></span><![endif]--><span")
                            objStreamWriter.WriteLine("style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;mso-bidi-font-family:")
                            objStreamWriter.WriteLine("""Times New Roman""'><span style='mso-tab-count:1'> </span></span><span")
                            objStreamWriter.WriteLine("lang=EN-US style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;")
                            objStreamWriter.WriteLine("mso-bidi-font-family:""Times New Roman"";mso-ansi-language:EN-US'>Page <span")
                            objStreamWriter.WriteLine("style='mso-field-code:"" PAGE ""'><span style='mso-no-proof:yes'>1</span></span>")
                            objStreamWriter.WriteLine("of 1</span><span class=MsoPageNumber><span style='font-size:8.0pt;mso-bidi-font-size:")
                            objStreamWriter.WriteLine("10.0pt;font-family:Arial;mso-bidi-font-family:""Times New Roman""'><span")
                            objStreamWriter.WriteLine("style='mso-tab-count:1'>                                                              </span>HR106<span")
                            objStreamWriter.WriteLine("style='mso-tab-count:2'>                                                                                                                                                                                                                                                        </span></span></span><span")
                            objStreamWriter.WriteLine("style='font-size:8.0pt;mso-bidi-font-size:10.0pt;font-family:Arial;mso-bidi-font-family:")
                            objStreamWriter.WriteLine("""Times New Roman""'><o:p></o:p></span></p>")
                            objStreamWriter.WriteLine("</div>")
                            objStreamWriter.WriteLine("</div>")


                            objStreamWriter.WriteLine("</body>")
                            objStreamWriter.WriteLine("")
                            objStreamWriter.WriteLine("</html>")


                            '*******************************************************
                            objStreamWriter.Close()
                            writtencounter = writtencounter + 1
                        End If
                        runner = runner + 1
                    End While

                    Message("HR106 forms successfully generated." & vbCrLf & writtencounter & " files were created.")
                    Overwrite_Files.Checked = initial_overwrite_status
                    ProgressBar1.Value = 0
                    Progress_Bar1_Panel.Visible = False
                    If Show_Created_Files.Checked = True Then
                        System.Diagnostics.Process.Start(foldername)
                    End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Generate_Multiple_Payments_Sheets()
    End Sub

    Private Sub MenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem27.Click
        Generate_Student_Timesheet()
    End Sub

    Private Sub MenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem28.Click
        Generate_Multiple_Payments_Sheets()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            Dim helpscreen As New Help
            
            helpscreen.Show()

        Catch et As Exception
            Error_Handler(et.Message)
        End Try
    End Sub
End Class
