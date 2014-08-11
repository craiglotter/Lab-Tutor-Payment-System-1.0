Public Class Student_Details_Display



    Inherits System.Windows.Forms.Form




#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Try
            Surname.Text = ""
            Staff_Number.Text = ""
            First_Name.Text = ""
            Student_Number.Text = ""
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
        
        'Add any initialization after the InitializeComponent() call

    End Sub

    Public Sub New(ByVal Surname_Input As String, ByVal Staff_Number_Input As String, ByVal First_Name_Input As String, ByVal Student_Number_Input As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Try
            Surname.Text = Surname_Input
            Staff_Number.Text = Staff_Number_Input
            First_Name.Text = First_Name_Input
            Student_Number.Text = Student_Number_Input
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
        

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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Staff_Number As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents First_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Student_Number As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Surname As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Student_Details_Display))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Staff_Number = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.First_Name = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Surname = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Student_Number = New System.Windows.Forms.TextBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Staff_Number)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.First_Name)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Surname)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Student_Number)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(400, 240)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Add Student Details"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.IndianRed
        Me.Label5.Location = New System.Drawing.Point(16, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(240, 32)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Fill in the form below and click on the 'Accept' button to add a student to the p" & _
        "ayment system."
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(264, 208)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 18)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Accept"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.NavajoWhite
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(328, 208)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 18)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Cancel"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 184)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Staff Number"
        '
        'Staff_Number
        '
        Me.Staff_Number.BackColor = System.Drawing.Color.SeaShell
        Me.Staff_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Staff_Number.Location = New System.Drawing.Point(16, 200)
        Me.Staff_Number.Name = "Staff_Number"
        Me.Staff_Number.Size = New System.Drawing.Size(136, 20)
        Me.Staff_Number.TabIndex = 4
        Me.Staff_Number.Text = ""
        Me.ToolTip1.SetToolTip(Me.Staff_Number, "Staff Number")
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 144)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "First Name"
        '
        'First_Name
        '
        Me.First_Name.BackColor = System.Drawing.Color.SeaShell
        Me.First_Name.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.First_Name.Location = New System.Drawing.Point(16, 160)
        Me.First_Name.Name = "First_Name"
        Me.First_Name.Size = New System.Drawing.Size(136, 20)
        Me.First_Name.TabIndex = 3
        Me.First_Name.Text = ""
        Me.ToolTip1.SetToolTip(Me.First_Name, "First Name")
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Last Name"
        '
        'Surname
        '
        Me.Surname.BackColor = System.Drawing.Color.SeaShell
        Me.Surname.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Surname.Location = New System.Drawing.Point(16, 120)
        Me.Surname.Name = "Surname"
        Me.Surname.Size = New System.Drawing.Size(136, 20)
        Me.Surname.TabIndex = 2
        Me.Surname.Text = ""
        Me.ToolTip1.SetToolTip(Me.Surname, "Surname")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Student Number"
        '
        'Student_Number
        '
        Me.Student_Number.BackColor = System.Drawing.Color.SeaShell
        Me.Student_Number.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Student_Number.Location = New System.Drawing.Point(16, 80)
        Me.Student_Number.MaxLength = 9
        Me.Student_Number.Name = "Student_Number"
        Me.Student_Number.Size = New System.Drawing.Size(136, 20)
        Me.Student_Number.TabIndex = 1
        Me.Student_Number.Text = ""
        Me.ToolTip1.SetToolTip(Me.Student_Number, "Student Number")
        '
        'Student_Details_Display
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(420, 291)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(422, 293)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(422, 293)
        Me.Name = "Student_Details_Display"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Lab Tutor Payment System :: Add Student"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        MsgBox("Sorry, but it would appear that Lab Tutor Payment System has encountered an error situation. The following error has been encountered: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Lab Tutor Payment System")
    End Sub



    Private Sub Form_Keypress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress, MyBase.KeyPress
        Try
            If e.KeyChar = Chr(13) Then
                MyBase.DialogResult = DialogResult.OK
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
        
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

    Private Sub Surname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Surname.Leave
        Surname.Text = capitalize_first_letter(Surname.Text)
    End Sub

    Private Sub First_Name_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles First_Name.Leave
        First_Name.Text = capitalize_first_letter(First_Name.Text)
    End Sub
End Class
