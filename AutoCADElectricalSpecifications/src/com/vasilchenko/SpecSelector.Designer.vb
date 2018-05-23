<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ufSpecSelector
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.rbSheetCreate = New System.Windows.Forms.RadioButton()
        Me.rbSheetUpdate = New System.Windows.Forms.RadioButton()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.rbProjCreate = New System.Windows.Forms.RadioButton()
        Me.rbProjUpdate = New System.Windows.Forms.RadioButton()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'rbSheetCreate
        '
        Me.rbSheetCreate.AutoSize = True
        Me.rbSheetCreate.Location = New System.Drawing.Point(13, 35)
        Me.rbSheetCreate.Name = "rbSheetCreate"
        Me.rbSheetCreate.Size = New System.Drawing.Size(67, 17)
        Me.rbSheetCreate.TabIndex = 0
        Me.rbSheetCreate.TabStop = True
        Me.rbSheetCreate.Text = "Создать"
        Me.rbSheetCreate.UseVisualStyleBackColor = True
        '
        'rbSheetUpdate
        '
        Me.rbSheetUpdate.AutoSize = True
        Me.rbSheetUpdate.Location = New System.Drawing.Point(13, 58)
        Me.rbSheetUpdate.Name = "rbSheetUpdate"
        Me.rbSheetUpdate.Size = New System.Drawing.Size(74, 17)
        Me.rbSheetUpdate.TabIndex = 1
        Me.rbSheetUpdate.TabStop = True
        Me.rbSheetUpdate.Text = "Обновить"
        Me.rbSheetUpdate.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(15, 15)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rbSheetCreate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.rbSheetUpdate)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rbProjCreate)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rbProjUpdate)
        Me.SplitContainer1.Size = New System.Drawing.Size(250, 100)
        Me.SplitContainer1.SplitterDistance = 125
        Me.SplitContainer1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Чертеж"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Проект"
        '
        'rbProjCreate
        '
        Me.rbProjCreate.AutoSize = True
        Me.rbProjCreate.Location = New System.Drawing.Point(13, 35)
        Me.rbProjCreate.Name = "rbProjCreate"
        Me.rbProjCreate.Size = New System.Drawing.Size(67, 17)
        Me.rbProjCreate.TabIndex = 4
        Me.rbProjCreate.TabStop = True
        Me.rbProjCreate.Text = "Создать"
        Me.rbProjCreate.UseVisualStyleBackColor = True
        '
        'rbProjUpdate
        '
        Me.rbProjUpdate.AutoSize = True
        Me.rbProjUpdate.Location = New System.Drawing.Point(13, 58)
        Me.rbProjUpdate.Name = "rbProjUpdate"
        Me.rbProjUpdate.Size = New System.Drawing.Size(74, 17)
        Me.rbProjUpdate.TabIndex = 5
        Me.rbProjUpdate.TabStop = True
        Me.rbProjUpdate.Text = "Обновить"
        Me.rbProjUpdate.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(197, 125)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(116, 125)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 4
        Me.btnOk.Text = "Apply"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'SpecSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 160)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "SpecSelector"
        Me.Text = "SpecSelector"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents rbSheetCreate As Windows.Forms.RadioButton
    Friend WithEvents rbSheetUpdate As Windows.Forms.RadioButton
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents rbProjCreate As Windows.Forms.RadioButton
    Friend WithEvents rbProjUpdate As Windows.Forms.RadioButton
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnOk As Windows.Forms.Button
End Class
