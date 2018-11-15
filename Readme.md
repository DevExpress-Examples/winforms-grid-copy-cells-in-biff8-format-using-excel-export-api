<!-- default file list -->
*Files to look at*:

* [CopyToClipboardHelper.cs](./CS/gridCopyToClipboardExample/CopyToClipboardHelper.cs) (VB: [CopyToClipboardHelper.vb](./VB/gridCopyToClipboardExample/CopyToClipboardHelper.vb))
* [Employee.cs](./CS/gridCopyToClipboardExample/Employee.cs) (VB: [Employee.vb](./VB/gridCopyToClipboardExample/Employee.vb))
* [Form1.cs](./CS/gridCopyToClipboardExample/Form1.cs) (VB: [Form1.vb](./VB/gridCopyToClipboardExample/Form1.vb))
* [Program.cs](./CS/gridCopyToClipboardExample/Program.cs) (VB: [Program.vb](./VB/gridCopyToClipboardExample/Program.vb))
<!-- default file list end -->
# How to copy cell data in an Excel format (BIFF8) to the clipboard using XL Export library


<p>This example demonstrates how to build data in BIFF8 format and pass it to the clipboard. This approach allows you to copy cells data within appearance (text alignment, borders, background color and font settings) and keep the cells number format as is while copy&paste data from a GridControl to an Excel-like document.<br /><br /><img src="https://raw.githubusercontent.com/DevExpress-Examples/how-to-copy-cell-data-in-an-excel-format-biff8-to-the-clipboard-using-xl-export-library-t266171/15.1.4+/media/a605af40-2950-11e5-80bf-00155d62480c.png"><br />                                                                        <br /><strong>Please see the "Implementation Details" (click the corresponding link below this text) to learn more about technical aspects of this approach implementation.</strong></p>


<h3>Description</h3>

<p>All the functionality to&nbsp;construct data in the BIFF8 format is encapsulated in the&nbsp;<strong>CopyToClipboardBIFF8Helper</strong>&nbsp;helper class. The main point of this approach is usage of the&nbsp;<strong>XL Export library</strong>&nbsp;for creating the BIFF8 data stream.<br /><br />First, a&nbsp;IXlDocument document is created and filled by columns and rows. Note that the XL Export library writes a document directly to the stream and requires adding&nbsp;all columns into a worksheet first and only then rows. The order is essential in this case. <br /><br />When the BIFF8 stream is ready, it is also necessary to add information of a path to data in the workbook&nbsp;to the clipboard. The link stream is constructed in the&nbsp;<strong>CreateLinkDataStream</strong> method. &nbsp;<br /><br />To substitute the default copy&amp;paste algorithm with this custom one, handle the<strong>&nbsp;GridControl.ProcessGridKey</strong>&nbsp;event and pass a&nbsp;<strong>GridView</strong>&nbsp;instance to the<strong>&nbsp;CopyToClipboardBIFF8Helper.CopySelectionToClipboard</strong>&nbsp;method.</p>

<br/>


