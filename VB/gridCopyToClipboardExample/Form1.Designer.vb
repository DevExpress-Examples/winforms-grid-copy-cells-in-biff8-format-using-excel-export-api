Namespace gridCopyToClipboardExample
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.gridControl1 = New DevExpress.XtraGrid.GridControl()
            Me.gridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.gridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.gridColumn2 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.gridColumn3 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.gridColumn4 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.gridColumn5 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.gridColumn6 = New DevExpress.XtraGrid.Columns.GridColumn()
            Me.panel1 = New System.Windows.Forms.Panel()
            Me.lblNote = New System.Windows.Forms.Label()
            DirectCast(Me.gridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            DirectCast(Me.gridView1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.panel1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' gridControl1
            ' 
            Me.gridControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.gridControl1.Location = New System.Drawing.Point(0, 41)
            Me.gridControl1.MainView = Me.gridView1
            Me.gridControl1.Name = "gridControl1"
            Me.gridControl1.Size = New System.Drawing.Size(986, 456)
            Me.gridControl1.TabIndex = 0
            Me.gridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() { Me.gridView1})
            ' 
            ' gridView1
            ' 
            Me.gridView1.Appearance.OddRow.BackColor = System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(224)))), (CInt((CByte(192)))))
            Me.gridView1.Appearance.OddRow.Font = New System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, (CByte(0)))
            Me.gridView1.Appearance.OddRow.Options.UseBackColor = True
            Me.gridView1.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() { Me.gridColumn1, Me.gridColumn2, Me.gridColumn3, Me.gridColumn4, Me.gridColumn5, Me.gridColumn6})
            Me.gridView1.GridControl = Me.gridControl1
            Me.gridView1.Name = "gridView1"
            Me.gridView1.OptionsSelection.MultiSelect = True
            Me.gridView1.OptionsView.EnableAppearanceOddRow = True
            Me.gridView1.OptionsView.ShowGroupPanel = False
            ' 
            ' gridColumn1
            ' 
            Me.gridColumn1.AppearanceCell.BackColor = System.Drawing.Color.White
            Me.gridColumn1.AppearanceCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            Me.gridColumn1.AppearanceCell.Options.UseFont = True
            Me.gridColumn1.AppearanceCell.Options.UseTextOptions = True
            Me.gridColumn1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.gridColumn1.AppearanceHeader.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            Me.gridColumn1.AppearanceHeader.Options.UseFont = True
            Me.gridColumn1.Caption = "Employee ID"
            Me.gridColumn1.FieldName = "Id"
            Me.gridColumn1.Name = "gridColumn1"
            Me.gridColumn1.Visible = True
            Me.gridColumn1.VisibleIndex = 0
            ' 
            ' gridColumn2
            ' 
            Me.gridColumn2.Caption = "Name"
            Me.gridColumn2.FieldName = "Name"
            Me.gridColumn2.Name = "gridColumn2"
            Me.gridColumn2.Visible = True
            Me.gridColumn2.VisibleIndex = 1
            ' 
            ' gridColumn3
            ' 
            Me.gridColumn3.Caption = "Salary"
            Me.gridColumn3.DisplayFormat.FormatString = "$ #,##0.00"
            Me.gridColumn3.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
            Me.gridColumn3.FieldName = "Salary"
            Me.gridColumn3.Name = "gridColumn3"
            Me.gridColumn3.Visible = True
            Me.gridColumn3.VisibleIndex = 2
            ' 
            ' gridColumn4
            ' 
            Me.gridColumn4.Caption = "Bonus"
            Me.gridColumn4.DisplayFormat.FormatString = "$ #,##0.00"
            Me.gridColumn4.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
            Me.gridColumn4.FieldName = "Bonus"
            Me.gridColumn4.Name = "gridColumn4"
            Me.gridColumn4.Visible = True
            Me.gridColumn4.VisibleIndex = 3
            ' 
            ' gridColumn5
            ' 
            Me.gridColumn5.Caption = "Department"
            Me.gridColumn5.FieldName = "Department"
            Me.gridColumn5.Name = "gridColumn5"
            Me.gridColumn5.Visible = True
            Me.gridColumn5.VisibleIndex = 4
            ' 
            ' gridColumn6
            ' 
            Me.gridColumn6.Caption = "Hired"
            Me.gridColumn6.DisplayFormat.FormatString = "d"
            Me.gridColumn6.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
            Me.gridColumn6.FieldName = "Hired"
            Me.gridColumn6.Name = "gridColumn6"
            Me.gridColumn6.Visible = True
            Me.gridColumn6.VisibleIndex = 5
            ' 
            ' panel1
            ' 
            Me.panel1.Controls.Add(Me.lblNote)
            Me.panel1.Dock = System.Windows.Forms.DockStyle.Top
            Me.panel1.Location = New System.Drawing.Point(0, 0)
            Me.panel1.Name = "panel1"
            Me.panel1.Size = New System.Drawing.Size(986, 41)
            Me.panel1.TabIndex = 1
            ' 
            ' lblNote
            ' 
            Me.lblNote.AutoSize = True
            Me.lblNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, (CByte(204)))
            Me.lblNote.Location = New System.Drawing.Point(12, 14)
            Me.lblNote.Name = "lblNote"
            Me.lblNote.Size = New System.Drawing.Size(517, 13)
            Me.lblNote.TabIndex = 1
            Me.lblNote.Text = "Select some grid rows, press ""Ctrl+C"", run Excel, select worksheet cell and press" & " ""Ctrl+V"""
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(986, 497)
            Me.Controls.Add(Me.gridControl1)
            Me.Controls.Add(Me.panel1)
            Me.Name = "Form1"
            Me.Text = "XtraGrid copy to clipboard example"
            DirectCast(Me.gridControl1, System.ComponentModel.ISupportInitialize).EndInit()
            DirectCast(Me.gridView1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.panel1.ResumeLayout(False)
            Me.panel1.PerformLayout()
            Me.ResumeLayout(False)

        End Sub

        #End Region

        Private WithEvents gridControl1 As DevExpress.XtraGrid.GridControl
        Private gridView1 As DevExpress.XtraGrid.Views.Grid.GridView
        Private panel1 As System.Windows.Forms.Panel
        Private gridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
        Private gridColumn2 As DevExpress.XtraGrid.Columns.GridColumn
        Private gridColumn3 As DevExpress.XtraGrid.Columns.GridColumn
        Private gridColumn4 As DevExpress.XtraGrid.Columns.GridColumn
        Private gridColumn5 As DevExpress.XtraGrid.Columns.GridColumn
        Private gridColumn6 As DevExpress.XtraGrid.Columns.GridColumn
        Private lblNote As System.Windows.Forms.Label
    End Class
End Namespace

