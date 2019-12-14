Namespace com.vasilchenko.Forms
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class SpecSelector
        Inherits System.Windows.Forms.Form

        'Форма переопределяет dispose для очистки списка компонентов.
        <System.Diagnostics.DebuggerNonUserCode()>
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
        <System.Diagnostics.DebuggerStepThrough()>
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
        Me.rbPageCreate = New System.Windows.Forms.RadioButton()
        Me.rbPageUpdate = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        CType(Me.SplitContainer1,System.ComponentModel.ISupportInitialize).BeginInit
        Me.SplitContainer1.Panel1.SuspendLayout
        Me.SplitContainer1.Panel2.SuspendLayout
        Me.SplitContainer1.SuspendLayout
        Me.SuspendLayout
        '
        'rbSheetCreate
        '
        Me.rbSheetCreate.AutoSize = true
        Me.rbSheetCreate.Location = New System.Drawing.Point(13, 35)
        Me.rbSheetCreate.Name = "rbSheetCreate"
        Me.rbSheetCreate.Size = New System.Drawing.Size(67, 17)
        Me.rbSheetCreate.TabIndex = 0
        Me.rbSheetCreate.TabStop = true
        Me.rbSheetCreate.Text = "Создать"
        Me.rbSheetCreate.UseVisualStyleBackColor = true
        '
        'rbSheetUpdate
        '
        Me.rbSheetUpdate.AutoSize = true
        Me.rbSheetUpdate.Location = New System.Drawing.Point(13, 58)
        Me.rbSheetUpdate.Name = "rbSheetUpdate"
        Me.rbSheetUpdate.Size = New System.Drawing.Size(74, 17)
        Me.rbSheetUpdate.TabIndex = 1
        Me.rbSheetUpdate.TabStop = true
        Me.rbSheetUpdate.Text = "Обновить"
        Me.rbSheetUpdate.UseVisualStyleBackColor = true
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 12)
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
        Me.Label1.AutoSize = true
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(95, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Конструкторская"
        '
        'Label2
        '
        Me.Label2.AutoSize = true
        Me.Label2.Location = New System.Drawing.Point(10, 10)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Проект"
        '
        'rbProjCreate
        '
        Me.rbProjCreate.AutoSize = true
        Me.rbProjCreate.Location = New System.Drawing.Point(13, 35)
        Me.rbProjCreate.Name = "rbProjCreate"
        Me.rbProjCreate.Size = New System.Drawing.Size(67, 17)
        Me.rbProjCreate.TabIndex = 4
        Me.rbProjCreate.TabStop = true
        Me.rbProjCreate.Text = "Создать"
        Me.rbProjCreate.UseVisualStyleBackColor = true
        '
        'rbProjUpdate
        '
        Me.rbProjUpdate.AutoSize = true
        Me.rbProjUpdate.Enabled = false
        Me.rbProjUpdate.Location = New System.Drawing.Point(13, 58)
        Me.rbProjUpdate.Name = "rbProjUpdate"
        Me.rbProjUpdate.Size = New System.Drawing.Size(74, 17)
        Me.rbProjUpdate.TabIndex = 5
        Me.rbProjUpdate.TabStop = true
        Me.rbProjUpdate.Text = "Обновить"
        Me.rbProjUpdate.UseVisualStyleBackColor = true
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(289, 125)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = true
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(208, 125)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 4
        Me.btnOk.Text = "Apply"
        Me.btnOk.UseVisualStyleBackColor = true
        '
        'rbPageCreate
        '
        Me.rbPageCreate.AutoSize = true
        Me.rbPageCreate.Location = New System.Drawing.Point(282, 47)
        Me.rbPageCreate.Name = "rbPageCreate"
        Me.rbPageCreate.Size = New System.Drawing.Size(67, 17)
        Me.rbPageCreate.TabIndex = 5
        Me.rbPageCreate.TabStop = true
        Me.rbPageCreate.Text = "Создать"
        Me.rbPageCreate.UseVisualStyleBackColor = true
        '
        'rbPageUpdate
        '
        Me.rbPageUpdate.AutoSize = true
        Me.rbPageUpdate.Location = New System.Drawing.Point(282, 70)
        Me.rbPageUpdate.Name = "rbPageUpdate"
        Me.rbPageUpdate.Size = New System.Drawing.Size(74, 17)
        Me.rbPageUpdate.TabIndex = 6
        Me.rbPageUpdate.TabStop = true
        Me.rbPageUpdate.Text = "Обновить"
        Me.rbPageUpdate.UseVisualStyleBackColor = true
        '
        'Label3
        '
        Me.Label3.AutoSize = true
        Me.Label3.Location = New System.Drawing.Point(279, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Листы"
        '
        'SpecSelector
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(376, 160)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rbPageUpdate)
        Me.Controls.Add(Me.rbPageCreate)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "SpecSelector"
        Me.Text = "SpecSelector"
        Me.SplitContainer1.Panel1.ResumeLayout(false)
        Me.SplitContainer1.Panel1.PerformLayout
        Me.SplitContainer1.Panel2.ResumeLayout(false)
        Me.SplitContainer1.Panel2.PerformLayout
        CType(Me.SplitContainer1,System.ComponentModel.ISupportInitialize).EndInit
        Me.SplitContainer1.ResumeLayout(false)
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub

        Public WithEvents rbSheetCreate As Windows.Forms.RadioButton
        Public WithEvents rbSheetUpdate As Windows.Forms.RadioButton
        Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
        Friend WithEvents Label1 As Windows.Forms.Label
        Friend WithEvents Label2 As Windows.Forms.Label
        Public WithEvents rbProjCreate As Windows.Forms.RadioButton
        Public WithEvents rbProjUpdate As Windows.Forms.RadioButton
        Friend WithEvents btnCancel As Windows.Forms.Button
        Friend WithEvents btnOk As Windows.Forms.Button
        Public WithEvents rbPageCreate As Windows.Forms.RadioButton
        Public WithEvents rbPageUpdate As Windows.Forms.RadioButton
        Friend WithEvents Label3 As Windows.Forms.Label
    End Class
End Namespace