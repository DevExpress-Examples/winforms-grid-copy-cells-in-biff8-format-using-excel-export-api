Imports System
Imports System.ComponentModel
Imports System.Windows.Forms

Namespace gridCopyToClipboardExample

    Public Partial Class Form1
        Inherits Form

        Private _copyToClipboardBIFF8Helper As CopyToClipboardBIFF8Helper

        Public Sub New()
            InitializeComponent()
            InitializeGrid()
            _copyToClipboardBIFF8Helper = New CopyToClipboardBIFF8Helper()
        End Sub

        Private Sub InitializeGrid()
            Dim dataList As BindingList(Of Employee) = New BindingList(Of Employee)()
            dataList.Add(New Employee(10115, "Augusta Delono", 1100.0, 50.0, "Accounting", New DateTime(2002, 1, 20)))
            dataList.Add(New Employee(10501, "Berry Dafoe", 1650.0, 150.0, "IT", New DateTime(2004, 5, 6)))
            dataList.Add(New Employee(10709, "Chris Cadwell", 2000.0, 180.0, "Management", New DateTime(2006, 12, 5)))
            dataList.Add(New Employee(10356, "Esta Mangold", 1400.0, 75.0, "Logistics", New DateTime(2004, 3, 12)))
            dataList.Add(New Employee(10401, "Frank Diamond", 1750.0, 100.0, "Marketing", New DateTime(2002, 4, 2)))
            dataList.Add(New Employee(10202, "Liam Bell", 1200.0, 80.0, "Manufacturing", New DateTime(2002, 4, 3)))
            dataList.Add(New Employee(10205, "Simon Newman", 1250.0, 80.0, "Manufacturing", New DateTime(2006, 5, 20)))
            dataList.Add(New Employee(10403, "Wendy Underwood", 1100.0, 50.0, "Marketing", New DateTime(2007, 9, 11)))
            gridControl1.DataSource = dataList
        End Sub

        Private Sub gridControl1_ProcessGridKey(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.Control AndAlso e.KeyCode = Keys.C Then
                _copyToClipboardBIFF8Helper.CopySelectionToClipboard(gridView1)
                e.Handled = True
            End If
        End Sub
    End Class
End Namespace
