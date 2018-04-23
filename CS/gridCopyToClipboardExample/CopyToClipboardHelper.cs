using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using DevExpress.Export.Xl;
using DevExpress.XtraGrid.Columns;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Drawing;
using System.Reflection;

namespace gridCopyToClipboardExample {
    public class CopyToClipboardBIFF8Helper {
        GridView view;
        string sheetName;
        XlCellRange dataRange;
        void ExportColumns(IXlSheet sheet) {
            foreach (GridColumn column in this.view.Columns)
                ExportColumn(sheet, column);
        }

        void ExportColumn(IXlSheet sheet, GridColumn gridColumn) {
            // Skip hidden column
            if (!gridColumn.Visible)
                return;

            using (IXlColumn column = sheet.CreateColumn()) {
                // Setup number format
                if(gridColumn.DisplayFormat.FormatType == FormatType.DateTime)
                    column.ApplyFormatting(XlCellFormatting.FromNetFormat(gridColumn.DisplayFormat.FormatString, true));
                else if(gridColumn.DisplayFormat.FormatType != FormatType.None)
                    column.ApplyFormatting(XlCellFormatting.FromNetFormat(gridColumn.DisplayFormat.FormatString, false));
            }
        }

        void ExportRows(IXlSheet sheet) {
            int[] selectedRows = this.view.GetSelectedRows();
            foreach (int gridRowHandle in selectedRows) {
                view.UnselectRow(gridRowHandle);
                view.RefreshRow(gridRowHandle);
                ExportRow(sheet, gridRowHandle);
                view.SelectRow(gridRowHandle);
            }
        }

        void ExportRow(IXlSheet sheet, int gridRowHandle) {
            using (IXlRow row = sheet.CreateRow()) {
                ExportCells(row, gridRowHandle);
            }
        }

        void ExportCells(IXlRow row, int gridRowHandle) {
            foreach (GridColumn column in this.view.Columns) {
                if (column.Visible) {
                    ExportCell(row, gridRowHandle, column);
                }
            }
        }

        void ExportCell(IXlRow row, int gridRowHandle, GridColumn gridColumn) {
            using(IXlCell cell = row.CreateCell()) {
                // Set cell value
                cell.Value = XlVariantValue.FromObject(this.view.GetRowCellValue(gridRowHandle, gridColumn));

                // Get cell appearance
                AppearanceObject appearance = GetCellAppearance(gridRowHandle, gridColumn);

                // Apply alignment
                XlCellAlignment alignment = new XlCellAlignment() {
                    WrapText = appearance.TextOptions.WordWrap.HasFlag(WordWrap.Wrap),
                    VerticalAlignment = ConvertAlignment(appearance.TextOptions.VAlignment),
                    HorizontalAlignment = ConvertAlignment(appearance.TextOptions.HAlignment)
                };
                cell.ApplyFormatting(alignment);

                // Apply borders
                Color borderColor = appearance.GetBorderColor();
                if (!DXColor.IsTransparentOrEmpty(borderColor))
                    cell.ApplyFormatting(XlBorder.OutlineBorders(borderColor));

                // Apply fill
                if (appearance.Options.UseBackColor)
                    cell.ApplyFormatting(XlFill.SolidFill(appearance.BackColor));
                
                // Apply font
                Font appearanceFont = appearance.Font;
                XlFont font = XlFont.CustomFont(appearanceFont.Name);
                font.Size = appearanceFont.SizeInPoints;
                font.Bold = appearanceFont.Bold;
                font.Italic = appearanceFont.Italic;
                font.StrikeThrough = appearanceFont.Strikeout;
                font.Underline = appearanceFont.Underline ? XlUnderlineType.Single : XlUnderlineType.None;
                if (appearance.Options.UseForeColor)
                    font.Color = appearance.ForeColor;
                cell.ApplyFormatting(font);
            }
        }

        AppearanceObject GetCellAppearance(int gridRowHandle, GridColumn gridColumn) {
            GridViewInfo viewInfo = view.GetViewInfo() as GridViewInfo;
            GridCellInfo cellInfo = viewInfo.GetGridCellInfo(gridRowHandle, gridColumn);
            if (cellInfo == null) {
                cellInfo = new GridCellInfo(new GridColumnInfoArgs(gridColumn), new GridDataRowInfo(viewInfo, gridRowHandle, view.GetRowLevel(gridRowHandle)), Rectangle.Empty);
            }
            MethodInfo me = viewInfo.GetType().GetMethod("UpdateCellAppearance", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);
            if (me != null)
                me.Invoke(viewInfo, new object[] {cellInfo, true });
            viewInfo.UpdateCellAppearance(cellInfo);
            return cellInfo.Appearance;
        }

        XlVerticalAlignment ConvertAlignment(VertAlignment verticalAlignment) {
            switch (verticalAlignment) {
                case VertAlignment.Top:
                    return XlVerticalAlignment.Top;
                case VertAlignment.Center:
                    return XlVerticalAlignment.Center;
                case VertAlignment.Bottom:
                    return XlVerticalAlignment.Bottom;
            }
            return XlVerticalAlignment.Bottom;
        }

        XlHorizontalAlignment ConvertAlignment(HorzAlignment horizontalAlignment) {
            switch (horizontalAlignment) {
                case HorzAlignment.Center:
                    return XlHorizontalAlignment.Center;
                case HorzAlignment.Near:
                    return XlHorizontalAlignment.Left;
                case HorzAlignment.Far:
                    return XlHorizontalAlignment.Right;
                default:
                    return XlHorizontalAlignment.General;
            }
        }

        MemoryStream CreateBIFF8DataStream() {
            IXlExporter exporter = XlExport.CreateExporter(XlDocumentFormat.Xls);
            MemoryStream dataStream = new MemoryStream();
            using(IXlDocument document = exporter.CreateDocument(dataStream)) {
                using(IXlSheet sheet = document.CreateSheet()) {
                    ExportColumns(sheet);
                    ExportRows(sheet);
                    this.sheetName = sheet.Name;
                    this.dataRange = sheet.DataRange;
                }
            }
            dataStream.Position = 0;
            return dataStream;
        }


        MemoryStream CreateLinkDataStream() {
            string link = string.Format("Excel\0[Book1]{0}\0{1}:{2}\0\0", sheetName, GetR1C1(this.dataRange.TopLeft), GetR1C1(this.dataRange.BottomRight));
            byte[] linkData = DXEncoding.Default.GetBytes(link);
            return new MemoryStream(linkData);
        }

        string GetR1C1(XlCellPosition cellPosition) {
            return string.Format("R{0}C{1}", cellPosition.Row + 1, cellPosition.Column + 1);
        }

        public void CopySelectionToClipboard(GridView view) {
            this.view = view;
            if(this.view.SelectedRowsCount < 1) {
                MessageBox.Show("Selection is empty", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MemoryStream biff8DataStream = CreateBIFF8DataStream();
            MemoryStream linkDataStream = CreateLinkDataStream();
            DataObject dataObject = new DataObject();
            dataObject.SetData("Biff8", biff8DataStream);
            dataObject.SetData("Link", linkDataStream);
            Clipboard.Clear();
            Clipboard.SetDataObject(dataObject, true);
        }
    }
    
}
