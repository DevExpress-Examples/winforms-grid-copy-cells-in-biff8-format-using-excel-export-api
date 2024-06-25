<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128626442/15.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T266171)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# WinForms Data Grid - Copy selected cell values to the clipboard in BIFF8 format using Excel Export API

This example demonstrates how to use the DevExpress Excel Export API to export the values in selected grid cells in BIFF8 format and copy them to the clipboard. This allows you to paste clipboard data to an MS Excel document maintaining appearance and formatting settings (text alignment, borders, background color, and font settings).

![Copy selected cell values to the clipboard using Excel Export API in BIFF8 format](https://raw.githubusercontent.com/DevExpress-Examples/how-to-copy-cell-data-in-an-excel-format-biff8-to-the-clipboard-using-xl-export-library-t266171/15.1.4+/media/a605af40-2950-11e5-80bf-00155d62480c.png)

See the implementation of the `CopyToClipboardBIFF8Helper` class for details:

* Create an `IXlDocument` document and populate it with columns and rows. The Excel Export API writes a document directly to the stream. Note that you should first add all the columns to the sheet and then add the rows.
  ```csharp
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
  ```
* When the BIFF8 stream is ready, add the path to the data in the workbook to the clipboard. The `CreateLinkDataStream` method creates a memory stream.
  ```csharp
  MemoryStream CreateLinkDataStream() {
      string link = string.Format("Excel\0[Book1]{0}\0{1}:{2}\0\0", sheetName, GetR1C1(this.dataRange.TopLeft), GetR1C1(this.dataRange.BottomRight));
      byte[] linkData = DXEncoding.Default.GetBytes(link);
      return new MemoryStream(linkData);
  }
  ```
* Handle the `GridControl.ProcessGridKey` event to copy selected grid cells to the clipboard in BIFF8 format:
  ```csharp
  void gridControl1_ProcessGridKey(object sender, KeyEventArgs e) {
      if(e.Control && e.KeyCode == Keys.C) {
          _copyToClipboardBIFF8Helper.CopySelectionToClipboard(gridView1);
          e.Handled = true;
      }
  }
  ```


## Files to Review

* [CopyToClipboardHelper.cs](./CS/gridCopyToClipboardExample/CopyToClipboardHelper.cs) (VB: [CopyToClipboardHelper.vb](./VB/gridCopyToClipboardExample/CopyToClipboardHelper.vb))
* [Form1.cs](./CS/gridCopyToClipboardExample/Form1.cs) (VB: [Form1.vb](./VB/gridCopyToClipboardExample/Form1.vb))
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-grid-copy-cells-in-biff8-format-using-excel-export-api&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=winforms-grid-copy-cells-in-biff8-format-using-excel-export-api&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
