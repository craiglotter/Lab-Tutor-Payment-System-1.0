Imports Lab_Tutor_Payment_System
Public Class Event_Details_Display



    Inherits System.Windows.Forms.Form

    Public start_date, end_date As Date



#Region " Windows Form Designer generated code "

    Public Sub New(ByVal Computer_Lab_Input As ListBox, ByVal Student_Number_Input As ListView)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Try
            Dim i As Long
            If Computer_Lab_Input.Items.Count > 0 Then

                For i = 0 To Computer_Lab_Input.Items.Count - 1
                    Computer_Lab.Items.Add(Convert.ToString(Computer_Lab_Input.Items.Item(i)))
                Next
                Computer_Lab.SelectedIndex = 0
            End If

            If Student_Number_Input.Items.Count > 0 Then


                For i = 0 To Student_Number_Input.Items.Count - 1
                    Student_Number.Items.Add(Student_Number_Input.Items.Item(i).Text)
                Next
                Student_Number.SelectedIndex = 0
            End If
            Event_Date.Value = Now()

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try

        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal Computer_Lab_Input As ListBox, ByVal Student_Number_Input As ListView, ByVal Event_Date_Input As Date, ByVal Student_Number_Input2 As String, ByVal Start_Time_Input As String, ByVal End_Time_Input As String, ByVal Computer_Lab_Input2 As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Try
            Dim i As Long
            If Computer_Lab_Input.Items.Count > 0 Then

                For i = 0 To Computer_Lab_Input.Items.Count - 1
                    Computer_Lab.Items.Add(Convert.ToString(Computer_Lab_Input.Items.Item(i)))
                Next
                Computer_Lab.SelectedIndex = 0
            End If

            If Student_Number_Input.Items.Count > 0 Then


                For i = 0 To Student_Number_Input.Items.Count - 1
                    Student_Number.Items.Add(Student_Number_Input.Items.Item(i).Text)
                Next
                Student_Number.SelectedIndex = 0
            End If
            Event_Date.Value = Now()

            Event_Date.Value = Event_Date_Input
            Student_Number.Items.Insert(0, Student_Number_Input2)
            Dim myarray As Array
            myarray = Start_Time_Input.Split(":")
            Start_Time1.Value = myarray(0)
            Start_Time2.Value = myarray(1)
            myarray = End_Time_Input.Split(":")
            End_Time1.Value = myarray(0)
            End_Time2.Value = myarray(1)
            Computer_Lab.Items.Insert(0, Computer_Lab_Input2)

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try

        'Add any initialization after the InitializeComponent() call

    End Sub


    'Public Sub New(ByVal Surname_Input As String, ByVal Staff_Number_Input As String, ByVal First_Name_Input As String, ByVal Student_Number_Input As String)
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()
    '    Try
    '        Surname.Text = Surname_Input
    '        Staff_Number.Text = Staff_Number_Input
    '        First_Name.Text = First_Name_Input
    '        Student_Number.Text = Student_Number_Input
    '    Catch ex As Exception
    '        Error_Handler(ex.Message)
    '    End Try


    '    'Add any initialization after the InitializeComponent() call

    'End Sub

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Computer_Lab As System.Windows.Forms.ComboBox
    Friend WithEvents Student_Number As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Event_Date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Start_Time1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents End_Time1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Start_Time2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents End_Time2 As System.Windows.Forms.NumericUpDown
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Event_Details_Display))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.End_Time2 = New System.Windows.Forms.NumericUpDown
        Me.Start_Time2 = New System.Windows.Forms.NumericUpDown
        Me.End_Time1 = New System.Windows.Forms.NumericUpDown
        Me.Start_Time1 = New System.Windows.Forms.NumericUpDown
        Me.Event_Date = New System.Windows.Forms.DateTimePicker
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Student_Number = New System.Windows.Forms.ComboBox
        Me.Computer_Lab = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.End_Time2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Start_Time2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.End_Time1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Start_Time1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.End_Time2)
        Me.GroupBox1.Controls.Add(Me.Start_Time2)
        Me.GroupBox1.Controls.Add(Me.End_Time1)
        Me.GroupBox1.Controls.Add(Me.Start_Time1)
        Me.GroupBox1.Controls.Add(Me.Event_Date)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Student_Number)
        Me.GroupBox1.Controls.Add(Me.Computer_Lab)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(400, 272)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Add Timesheet Event"
        '
        'End_Time2
        '
        Me.End_Time2.BackColor = System.Drawing.Color.SeaShell
        Me.End_Time2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.End_Time2.ForeColor = System.Drawing.Color.Black
        Me.End_Time2.Location = New System.Drawing.Point(64, 200)
        Me.End_Time2.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.End_Time2.Name = "End_Time2"
        Me.End_Time2.Size = New System.Drawing.Size(40, 20)
        Me.End_Time2.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.End_Time2, "Minutes, e.g. 30")
        '
        'Start_Time2
        '
        Me.Start_Time2.BackColor = System.Drawing.Color.SeaShell
        Me.Start_Time2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Start_Time2.ForeColor = System.Drawing.Color.Black
        Me.Start_Time2.Location = New System.Drawing.Point(64, 160)
        Me.Start_Time2.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.Start_Time2.Name = "Start_Time2"
        Me.Start_Time2.Size = New System.Drawing.Size(40, 20)
        Me.Start_Time2.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.Start_Time2, "Minutes, e.g. 30")
        '
        'End_Time1
        '
        Me.End_Time1.BackColor = System.Drawing.Color.SeaShell
        Me.End_Time1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.End_Time1.ForeColor = System.Drawing.Color.Black
        Me.End_Time1.Location = New System.Drawing.Point(16, 200)
        Me.End_Time1.Maximum = New Decimal(New Integer() {23, 0, 0, 0})
        Me.End_Time1.Name = "End_Time1"
        Me.End_Time1.Size = New System.Drawing.Size(40, 20)
        Me.End_Time1.TabIndex = 5
        Me.End_Time1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.End_Time1, "Hours, e.g. 13 (24 Hour Clock)")
        Me.End_Time1.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Start_Time1
        '
        Me.Start_Time1.BackColor = System.Drawing.Color.SeaShell
        Me.Start_Time1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Start_Time1.ForeColor = System.Drawing.Color.Black
        Me.Start_Time1.Location = New System.Drawing.Point(16, 160)
        Me.Start_Time1.Maximum = New Decimal(New Integer() {23, 0, 0, 0})
        Me.Start_Time1.Name = "Start_Time1"
        Me.Start_Time1.Size = New System.Drawing.Size(40, 20)
        Me.Start_Time1.TabIndex = 3
        Me.Start_Time1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.ToolTip1.SetToolTip(Me.Start_Time1, "Hours, e.g. 13 (24 Hour Clock)")
        Me.Start_Time1.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left
        '
        'Event_Date
        '
        Me.Event_Date.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Event_Date.CalendarForeColor = System.Drawing.Color.Black
        Me.Event_Date.CalendarMonthBackground = System.Drawing.Color.SeaShell
        Me.Event_Date.CalendarTitleBackColor = System.Drawing.Color.NavajoWhite
        Me.Event_Date.CalendarTitleForeColor = System.Drawing.Color.Black
        Me.Event_Date.CalendarTrailingForeColor = System.Drawing.Color.LightGray
        Me.Event_Date.CustomFormat = "yy/mm/dd"
        Me.Event_Date.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Event_Date.Location = New System.Drawing.Point(16, 80)
        Me.Event_Date.Name = "Event_Date"
        Me.Event_Date.Size = New System.Drawing.Size(136, 20)
        Me.Event_Date.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Event_Date, "Timesheet Event Date")
        Me.Event_Date.Value = New Date(2004, 9, 22, 0, 0, 0, 0)
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(56, 200)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(8, 16)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = ":"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(56, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(8, 16)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = ":"
        '
        'Student_Number
        '
        Me.Student_Number.BackColor = System.Drawing.Color.SeaShell
        Me.Student_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Student_Number.Location = New System.Drawing.Point(16, 120)
        Me.Student_Number.MaxDropDownItems = 20
        Me.Student_Number.MaxLength = 9
        Me.Student_Number.Name = "Student_Number"
        Me.Student_Number.Size = New System.Drawing.Size(136, 21)
        Me.Student_Number.Sorted = True
        Me.Student_Number.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Student_Number, "Student Number")
        '
        'Computer_Lab
        '
        Me.Computer_Lab.BackColor = System.Drawing.Color.SeaShell
        Me.Computer_Lab.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Computer_Lab.Location = New System.Drawing.Point(16, 240)
        Me.Computer_Lab.MaxDropDownItems = 20
        Me.Computer_Lab.Name = "Computer_Lab"
        Me.Computer_Lab.Size = New System.Drawing.Size(136, 21)
        Me.Computer_Lab.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.Computer_Lab, "Computer Lab Tutored")
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 224)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 16)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Computer Lab"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.IndianRed
        Me.Label5.Location = New System.Drawing.Point(16, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 32)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Fill in the form below and click on the 'Accept' button to add an event to the pa" & _
        "yment system."
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(272, 246)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 18)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Accept"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(336, 246)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 18)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "Cancel"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 184)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "End Time"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Start Time"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Student Number"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Date"
        '
        'Event_Details_Display
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(424, 321)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(426, 323)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(426, 323)
        Me.Name = "Event_Details_Display"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Lab Tutor Payment System :: Add Timesheet Event"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.End_Time2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Start_Time2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.End_Time1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Start_Time1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        MsgBox("Sorry, but it would appear that Lab Tutor Payment System has encountered an error situation. The following error has been encountered: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Lab Tutor Payment System")
    End Sub

    Private Sub Message_Handler(ByVal message As String)
        MsgBox(message, MsgBoxStyle.Information, "Lab Tutor Payment System")
    End Sub

    Private Sub Form_Keypress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress, MyBase.KeyPress
        Try
            If e.KeyChar = Chr(13) Then

            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try

    End Sub

    Private Function Validate_Input() As Boolean
        Dim result As Boolean
        result = True
        Try



            Dim message As String
            message = ""

            If Student_Number.Text.Trim().Equals("") Or Student_Number.Text.Trim().Length < 9 Then
                message = message & "- You have entered an invalid student number." & vbCrLf
            End If

            If Computer_Lab.Text.Trim().Equals("") Then
                message = message & "- You have neglected to enter a computer lab." & vbCrLf
            End If

            If IsDate(Event_Date.Value) = False Then
                message = message & "- You have selected an invalid event date." & vbCrLf
            End If

            Dim starttime, endtime As Date
            starttime = New Date(2000, 1, 1, Start_Time1.Value, Start_Time2.Value, 0)
            endtime = New Date(2000, 1, 1, End_Time1.Value, End_Time2.Value, 0)
            If (DateDiff(DateInterval.Minute, starttime, endtime)) < 1 Then
                message = message & "- Your end time cannot be before your start time. Please remember that a 24h clock convention is being used." & vbCrLf
            End If

            start_date = starttime
            end_date = starttime

            If message.Equals("") = False Then
                Message_Handler(message)
                result = False
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
            result = False
        End Try
        Return result
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Validate_Input() = True Then
            MyBase.DialogResult = DialogResult.OK
        End If
    End Sub


    Private Sub Start_Time1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start_Time1.Leave
        Try
            If Start_Time1.Value > Start_Time1.Maximum Then
                Start_Time1.Value = Start_Time1.Maximum
            End If
            If Start_Time1.Value < Start_Time1.Minimum Then
                Start_Time1.Value = Start_Time1.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Start_Time2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start_Time2.Leave
        Try
            If Start_Time2.Value > Start_Time2.Maximum Then
                Start_Time2.Value = Start_Time2.Maximum
            End If
            If Start_Time2.Value < Start_Time2.Minimum Then
                Start_Time2.Value = Start_Time2.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub End_Time1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_Time1.Leave
        Try
            If End_Time1.Value > End_Time1.Maximum Then
                End_Time1.Value = End_Time1.Maximum
            End If
            If End_Time1.Value < End_Time1.Minimum Then
                End_Time1.Value = End_Time1.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub End_Time2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles End_Time2.Leave
        Try
            If End_Time2.Value > End_Time2.Maximum Then
                End_Time2.Value = End_Time2.Maximum
            End If
            If End_Time2.Value < End_Time2.Minimum Then
                End_Time2.Value = End_Time2.Minimum
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    
End Class
