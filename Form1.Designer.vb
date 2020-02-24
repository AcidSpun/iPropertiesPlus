<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.tbDescription = New System.Windows.Forms.TextBox()
        Me.cbTitle = New System.Windows.Forms.ComboBox()
        Me.lblNextProcess = New System.Windows.Forms.Label()
        Me.cbNextProcess = New System.Windows.Forms.ComboBox()
        Me.lblRawMaterial = New System.Windows.Forms.Label()
        Me.cbRawMaterial = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cbType = New System.Windows.Forms.ComboBox()
        Me.lblSPClass = New System.Windows.Forms.Label()
        Me.cbSPClass = New System.Windows.Forms.ComboBox()
        Me.lblManPartNum = New System.Windows.Forms.Label()
        Me.tbManPartNum = New System.Windows.Forms.TextBox()
        Me.lblNextProcessKey = New System.Windows.Forms.Label()
        Me.tbNextProcessKey = New System.Windows.Forms.TextBox()
        Me.lblRawMaterialPartNumber = New System.Windows.Forms.Label()
        Me.tbRawMaterialPartNumber = New System.Windows.Forms.TextBox()
        Me.lblTypeNumber = New System.Windows.Forms.Label()
        Me.tbTypeNumber = New System.Windows.Forms.TextBox()
        Me.lblPropertyType = New System.Windows.Forms.Label()
        Me.tbPropertyType = New System.Windows.Forms.TextBox()
        Me.lblManufaturer = New System.Windows.Forms.Label()
        Me.tbManufacturer = New System.Windows.Forms.TextBox()
        Me.btOK = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Location = New System.Drawing.Point(28, 14)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(27, 13)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title"
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Location = New System.Drawing.Point(28, 57)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(60, 13)
        Me.lblDescription.TabIndex = 2
        Me.lblDescription.Text = "Description"
        '
        'tbDescription
        '
        Me.tbDescription.Location = New System.Drawing.Point(31, 73)
        Me.tbDescription.Name = "tbDescription"
        Me.tbDescription.Size = New System.Drawing.Size(476, 20)
        Me.tbDescription.TabIndex = 3
        '
        'cbTitle
        '
        Me.cbTitle.FormattingEnabled = True
        Me.cbTitle.Location = New System.Drawing.Point(31, 31)
        Me.cbTitle.Name = "cbTitle"
        Me.cbTitle.Size = New System.Drawing.Size(476, 21)
        Me.cbTitle.TabIndex = 4
        Me.cbTitle.Text = "Select Title"
        '
        'lblNextProcess
        '
        Me.lblNextProcess.AutoSize = True
        Me.lblNextProcess.Location = New System.Drawing.Point(31, 101)
        Me.lblNextProcess.Name = "lblNextProcess"
        Me.lblNextProcess.Size = New System.Drawing.Size(67, 13)
        Me.lblNextProcess.TabIndex = 5
        Me.lblNextProcess.Text = "NextProcess"
        '
        'cbNextProcess
        '
        Me.cbNextProcess.FormattingEnabled = True
        Me.cbNextProcess.Location = New System.Drawing.Point(31, 118)
        Me.cbNextProcess.Name = "cbNextProcess"
        Me.cbNextProcess.Size = New System.Drawing.Size(218, 21)
        Me.cbNextProcess.TabIndex = 6
        Me.cbNextProcess.Text = "Select Next Process"
        '
        'lblRawMaterial
        '
        Me.lblRawMaterial.AutoSize = True
        Me.lblRawMaterial.Location = New System.Drawing.Point(31, 147)
        Me.lblRawMaterial.Name = "lblRawMaterial"
        Me.lblRawMaterial.Size = New System.Drawing.Size(69, 13)
        Me.lblRawMaterial.TabIndex = 7
        Me.lblRawMaterial.Text = "Raw Material"
        '
        'cbRawMaterial
        '
        Me.cbRawMaterial.FormattingEnabled = True
        Me.cbRawMaterial.Location = New System.Drawing.Point(31, 163)
        Me.cbRawMaterial.Name = "cbRawMaterial"
        Me.cbRawMaterial.Size = New System.Drawing.Size(218, 21)
        Me.cbRawMaterial.TabIndex = 8
        Me.cbRawMaterial.Text = "Select Raw Material"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Location = New System.Drawing.Point(31, 187)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(31, 13)
        Me.lblType.TabIndex = 9
        Me.lblType.Text = "Type"
        '
        'cbType
        '
        Me.cbType.FormattingEnabled = True
        Me.cbType.Location = New System.Drawing.Point(31, 204)
        Me.cbType.Name = "cbType"
        Me.cbType.Size = New System.Drawing.Size(218, 21)
        Me.cbType.TabIndex = 10
        Me.cbType.Text = "Select Type"
        '
        'lblSPClass
        '
        Me.lblSPClass.AutoSize = True
        Me.lblSPClass.Location = New System.Drawing.Point(31, 233)
        Me.lblSPClass.Name = "lblSPClass"
        Me.lblSPClass.Size = New System.Drawing.Size(121, 13)
        Me.lblSPClass.TabIndex = 11
        Me.lblSPClass.Text = "Spare Part Classification"
        '
        'cbSPClass
        '
        Me.cbSPClass.FormattingEnabled = True
        Me.cbSPClass.Location = New System.Drawing.Point(31, 249)
        Me.cbSPClass.Name = "cbSPClass"
        Me.cbSPClass.Size = New System.Drawing.Size(218, 21)
        Me.cbSPClass.TabIndex = 12
        Me.cbSPClass.Text = "Select Spare Part Classification"
        '
        'lblManPartNum
        '
        Me.lblManPartNum.AutoSize = True
        Me.lblManPartNum.Location = New System.Drawing.Point(31, 278)
        Me.lblManPartNum.Name = "lblManPartNum"
        Me.lblManPartNum.Size = New System.Drawing.Size(126, 13)
        Me.lblManPartNum.TabIndex = 13
        Me.lblManPartNum.Text = "Manufaturer Part Number"
        '
        'tbManPartNum
        '
        Me.tbManPartNum.Location = New System.Drawing.Point(31, 294)
        Me.tbManPartNum.Name = "tbManPartNum"
        Me.tbManPartNum.Size = New System.Drawing.Size(218, 20)
        Me.tbManPartNum.TabIndex = 16
        '
        'lblNextProcessKey
        '
        Me.lblNextProcessKey.AutoSize = True
        Me.lblNextProcessKey.Location = New System.Drawing.Point(300, 101)
        Me.lblNextProcessKey.Name = "lblNextProcessKey"
        Me.lblNextProcessKey.Size = New System.Drawing.Size(91, 13)
        Me.lblNextProcessKey.TabIndex = 17
        Me.lblNextProcessKey.Text = "Next Process Key"
        '
        'tbNextProcessKey
        '
        Me.tbNextProcessKey.Location = New System.Drawing.Point(303, 118)
        Me.tbNextProcessKey.Name = "tbNextProcessKey"
        Me.tbNextProcessKey.Size = New System.Drawing.Size(205, 20)
        Me.tbNextProcessKey.TabIndex = 18
        '
        'lblRawMaterialPartNumber
        '
        Me.lblRawMaterialPartNumber.AutoSize = True
        Me.lblRawMaterialPartNumber.Location = New System.Drawing.Point(300, 147)
        Me.lblRawMaterialPartNumber.Name = "lblRawMaterialPartNumber"
        Me.lblRawMaterialPartNumber.Size = New System.Drawing.Size(131, 13)
        Me.lblRawMaterialPartNumber.TabIndex = 19
        Me.lblRawMaterialPartNumber.Text = "Raw Material Part Number"
        '
        'tbRawMaterialPartNumber
        '
        Me.tbRawMaterialPartNumber.Location = New System.Drawing.Point(303, 163)
        Me.tbRawMaterialPartNumber.Name = "tbRawMaterialPartNumber"
        Me.tbRawMaterialPartNumber.Size = New System.Drawing.Size(205, 20)
        Me.tbRawMaterialPartNumber.TabIndex = 20
        '
        'lblTypeNumber
        '
        Me.lblTypeNumber.AutoSize = True
        Me.lblTypeNumber.Location = New System.Drawing.Point(300, 188)
        Me.lblTypeNumber.Name = "lblTypeNumber"
        Me.lblTypeNumber.Size = New System.Drawing.Size(71, 13)
        Me.lblTypeNumber.TabIndex = 21
        Me.lblTypeNumber.Text = "Type Number"
        '
        'tbTypeNumber
        '
        Me.tbTypeNumber.Location = New System.Drawing.Point(303, 204)
        Me.tbTypeNumber.Name = "tbTypeNumber"
        Me.tbTypeNumber.Size = New System.Drawing.Size(205, 20)
        Me.tbTypeNumber.TabIndex = 22
        '
        'lblPropertyType
        '
        Me.lblPropertyType.AutoSize = True
        Me.lblPropertyType.Location = New System.Drawing.Point(300, 233)
        Me.lblPropertyType.Name = "lblPropertyType"
        Me.lblPropertyType.Size = New System.Drawing.Size(73, 13)
        Me.lblPropertyType.TabIndex = 23
        Me.lblPropertyType.Text = "Property Type"
        '
        'tbPropertyType
        '
        Me.tbPropertyType.Location = New System.Drawing.Point(303, 250)
        Me.tbPropertyType.Name = "tbPropertyType"
        Me.tbPropertyType.Size = New System.Drawing.Size(205, 20)
        Me.tbPropertyType.TabIndex = 24
        '
        'lblManufaturer
        '
        Me.lblManufaturer.AutoSize = True
        Me.lblManufaturer.Location = New System.Drawing.Point(300, 278)
        Me.lblManufaturer.Name = "lblManufaturer"
        Me.lblManufaturer.Size = New System.Drawing.Size(70, 13)
        Me.lblManufaturer.TabIndex = 25
        Me.lblManufaturer.Text = "Manufacturer"
        '
        'tbManufacturer
        '
        Me.tbManufacturer.Location = New System.Drawing.Point(303, 294)
        Me.tbManufacturer.Name = "tbManufacturer"
        Me.tbManufacturer.Size = New System.Drawing.Size(205, 20)
        Me.tbManufacturer.TabIndex = 26
        '
        'btOK
        '
        Me.btOK.Location = New System.Drawing.Point(352, 330)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(75, 23)
        Me.btOK.TabIndex = 29
        Me.btOK.Text = "OK"
        Me.btOK.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Location = New System.Drawing.Point(433, 330)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 23)
        Me.btCancel.TabIndex = 30
        Me.btCancel.Text = "Cancel"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 378)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Controls.Add(Me.tbManufacturer)
        Me.Controls.Add(Me.lblManufaturer)
        Me.Controls.Add(Me.tbPropertyType)
        Me.Controls.Add(Me.lblPropertyType)
        Me.Controls.Add(Me.tbTypeNumber)
        Me.Controls.Add(Me.lblTypeNumber)
        Me.Controls.Add(Me.tbRawMaterialPartNumber)
        Me.Controls.Add(Me.lblRawMaterialPartNumber)
        Me.Controls.Add(Me.tbNextProcessKey)
        Me.Controls.Add(Me.lblNextProcessKey)
        Me.Controls.Add(Me.tbManPartNum)
        Me.Controls.Add(Me.lblManPartNum)
        Me.Controls.Add(Me.cbSPClass)
        Me.Controls.Add(Me.lblSPClass)
        Me.Controls.Add(Me.cbType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.cbRawMaterial)
        Me.Controls.Add(Me.lblRawMaterial)
        Me.Controls.Add(Me.cbNextProcess)
        Me.Controls.Add(Me.lblNextProcess)
        Me.Controls.Add(Me.cbTitle)
        Me.Controls.Add(Me.tbDescription)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "Form1"
        Me.Text = "iProperties Plus"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblTitle As Windows.Forms.Label
    Friend WithEvents lblDescription As Windows.Forms.Label
    Friend WithEvents tbDescription As Windows.Forms.TextBox
    Friend WithEvents cbTitle As Windows.Forms.ComboBox
    Friend WithEvents lblNextProcess As Windows.Forms.Label
    Friend WithEvents cbNextProcess As Windows.Forms.ComboBox
    Friend WithEvents lblRawMaterial As Windows.Forms.Label
    Friend WithEvents cbRawMaterial As Windows.Forms.ComboBox
    Friend WithEvents lblType As Windows.Forms.Label
    Friend WithEvents cbType As Windows.Forms.ComboBox
    Friend WithEvents lblSPClass As Windows.Forms.Label
    Friend WithEvents cbSPClass As Windows.Forms.ComboBox
    Friend WithEvents lblManPartNum As Windows.Forms.Label
    Friend WithEvents tbManPartNum As Windows.Forms.TextBox
    Friend WithEvents lblNextProcessKey As Windows.Forms.Label
    Friend WithEvents tbNextProcessKey As Windows.Forms.TextBox
    Friend WithEvents lblRawMaterialPartNumber As Windows.Forms.Label
    Friend WithEvents tbRawMaterialPartNumber As Windows.Forms.TextBox
    Friend WithEvents lblTypeNumber As Windows.Forms.Label
    Friend WithEvents tbTypeNumber As Windows.Forms.TextBox
    Friend WithEvents lblPropertyType As Windows.Forms.Label
    Friend WithEvents tbPropertyType As Windows.Forms.TextBox
    Friend WithEvents lblManufaturer As Windows.Forms.Label
    Friend WithEvents tbManufacturer As Windows.Forms.TextBox
    Friend WithEvents btOK As Windows.Forms.Button
    Friend WithEvents btCancel As Windows.Forms.Button
End Class
