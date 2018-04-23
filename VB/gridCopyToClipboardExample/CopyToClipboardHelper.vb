Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.IO
Imports System.Windows.Forms
Imports DevExpress.Export.Xl
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.Utils
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraGrid.Drawing
Imports System.Reflection

Namespace gridCopyToClipboardExample
    Public Class CopyToClipboardBIFF8Helper
        Private view As GridView
        Private sheetName As String
        Private dataRange As XlCellRange
        Private Sub ExportColumns(ByVal sheet As IXlSheet)
            For Each column As GridColumn In Me.view.Columns
                ExportColumn(sheet, column)
            Next column
        End Sub

        Private Sub ExportColumn(ByVal sheet As IXlSheet, ByVal gridColumn As GridColumn)
            ' Skip hidden column
            If Not gridColumn.Visible Then
                Return
            End If

            Using column As IXlColumn = sheet.CreateColumn()
                ' Setup number format
                If gridColumn.DisplayFormat.FormatType = FormatType.DateTime Then
                    column.ApplyFormatting(XlCellFormatting.FromNetFormat(gridColumn.DisplayFormat.FormatString, True))
                ElseIf gridColumn.DisplayFormat.FormatType <> FormatType.None Then
                    column.ApplyFormatting(XlCellFormatting.FromNetFormat(gridColumn.DisplayFormat.FormatString, False))
                End If
            End Using
        End Sub

        Private Sub ExportRows(ByVal sheet As IXlSheet)
            Dim selectedRows() As Integer = Me.view.GetSelectedRows()
            For Each gridRowHandle As Integer In selectedRows
                view.UnselectRow(gridRowHandle)
                view.RefreshRow(gridRowHandle)
                ExportRow(sheet, gridRowHandle)
                view.SelectRow(gridRowHandle)
            Next gridRowHandle
        End Sub

        Private Sub ExportRow(ByVal sheet As IXlSheet, ByVal gridRowHandle As Integer)
            Using row As IXlRow = sheet.CreateRow()
                ExportCells(row, gridRowHandle)
            End Using
        End Sub

        Private Sub ExportCells(ByVal row As IXlRow, ByVal gridRowHandle As Integer)
            For Each column As GridColumn In Me.view.Columns
                If column.Visible Then
                    ExportCell(row, gridRowHandle, column)
                End If
            Next column
        End Sub

        Private Sub ExportCell(ByVal row As IXlRow, ByVal gridRowHandle As Integer, ByVal gridColumn As GridColumn)
            Using cell As IXlCell = row.CreateCell()
                ' Set cell value
                cell.Value = XlVariantValue.FromObject(Me.view.GetRowCellValue(gridRowHandle, gridColumn))

                ' Get cell appearance
                Dim appearance As AppearanceObject = GetCellAppearance(gridRowHandle, gridColumn)

                ' Apply alignment
                Dim alignment As New XlCellAlignment() With {.WrapText = appearance.TextOptions.WordWrap.HasFlag(WordWrap.Wrap), .VerticalAlignment = ConvertAlignment(appearance.TextOptions.VAlignment), .HorizontalAlignment = ConvertAlignment(appearance.TextOptions.HAlignment)}
                cell.ApplyFormatting(alignment)

                ' Apply borders
                Dim borderColor As Color = appearance.GetBorderColor()
                If Not DXColor.IsTransparentOrEmpty(borderColor) Then
                    cell.ApplyFormatting(XlBorder.OutlineBorders(borderColor))
                End If

                ' Apply fill
                If appearance.Options.UseBackColor Then
                    cell.ApplyFormatting(XlFill.SolidFill(appearance.BackColor))
                End If

                ' Apply font
                Dim appearanceFont As Font = appearance.Font
                Dim font As XlFont = XlFont.CustomFont(appearanceFont.Name)
                font.Size = appearanceFont.SizeInPoints
                font.Bold = appearanceFont.Bold
                font.Italic = appearanceFont.Italic
                font.StrikeThrough = appearanceFont.Strikeout
                font.Underline = If(appearanceFont.Underline, XlUnderlineType.Single, XlUnderlineType.None)
                If appearance.Options.UseForeColor Then
                    font.Color = appearance.ForeColor
                End If
                cell.ApplyFormatting(font)
            End Using
        End Sub

        Private Function GetCellAppearance(ByVal gridRowHandle As Integer, ByVal gridColumn As GridColumn) As AppearanceObject
            Dim viewInfo As GridViewInfo = TryCast(view.GetViewInfo(), GridViewInfo)
            Dim cellInfo As GridCellInfo = viewInfo.GetGridCellInfo(gridRowHandle, gridColumn)
            If cellInfo Is Nothing Then
                cellInfo = New GridCellInfo(New GridColumnInfoArgs(gridColumn), New GridDataRowInfo(viewInfo, gridRowHandle, view.GetRowLevel(gridRowHandle)), Rectangle.Empty)
            End If
            Dim [me] As MethodInfo = viewInfo.GetType().GetMethod("UpdateCellAppearance", BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.DeclaredOnly)
            If [me] IsNot Nothing Then
                [me].Invoke(viewInfo, New Object() {cellInfo, True })
            End If
            viewInfo.UpdateCellAppearance(cellInfo)
            Return cellInfo.Appearance
        End Function

        Private Function ConvertAlignment(ByVal verticalAlignment As VertAlignment) As XlVerticalAlignment
            Select Case verticalAlignment
                Case VertAlignment.Top
                    Return XlVerticalAlignment.Top
                Case VertAlignment.Center
                    Return XlVerticalAlignment.Center
                Case VertAlignment.Bottom
                    Return XlVerticalAlignment.Bottom
            End Select
            Return XlVerticalAlignment.Bottom
        End Function

        Private Function ConvertAlignment(ByVal horizontalAlignment As HorzAlignment) As XlHorizontalAlignment
            Select Case horizontalAlignment
                Case HorzAlignment.Center
                    Return XlHorizontalAlignment.Center
                Case HorzAlignment.Near
                    Return XlHorizontalAlignment.Left
                Case HorzAlignment.Far
                    Return XlHorizontalAlignment.Right
                Case Else
                    Return XlHorizontalAlignment.General
            End Select
        End Function

        Private Function CreateBIFF8DataStream() As MemoryStream
            Dim exporter As IXlExporter = XlExport.CreateExporter(XlDocumentFormat.Xls)
            Dim dataStream As New MemoryStream()
            Using document As IXlDocument = exporter.CreateDocument(dataStream)
                Using sheet As IXlSheet = document.CreateSheet()
                    ExportColumns(sheet)
                    ExportRows(sheet)
                    Me.sheetName = sheet.Name
                    Me.dataRange = sheet.DataRange
                End Using
            End Using
            dataStream.Position = 0
            Return dataStream
        End Function


        Private Function CreateLinkDataStream() As MemoryStream
            Dim link As String = String.Format("Excel" & ControlChars.NullChar & "[Book1]{0}" & ControlChars.NullChar & "{1}:{2}" & ControlChars.NullChar & ControlChars.NullChar, sheetName, GetR1C1(Me.dataRange.TopLeft), GetR1C1(Me.dataRange.BottomRight))
            Dim linkData() As Byte = DXEncoding.Default.GetBytes(link)
            Return New MemoryStream(linkData)
        End Function

        Private Function GetR1C1(ByVal cellPosition As XlCellPosition) As String
            Return String.Format("R{0}C{1}", cellPosition.Row + 1, cellPosition.Column + 1)
        End Function

        Public Sub CopySelectionToClipboard(ByVal view As GridView)
            Me.view = view
            If Me.view.SelectedRowsCount < 1 Then
                MessageBox.Show("Selection is empty", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim biff8DataStream As MemoryStream = CreateBIFF8DataStream()
            Dim linkDataStream As MemoryStream = CreateLinkDataStream()
            Dim dataObject As New DataObject()
            dataObject.SetData("Biff8", biff8DataStream)
            dataObject.SetData("Link", linkDataStream)
            Clipboard.Clear()
            Clipboard.SetDataObject(dataObject, True)
        End Sub
    End Class

End Namespace
