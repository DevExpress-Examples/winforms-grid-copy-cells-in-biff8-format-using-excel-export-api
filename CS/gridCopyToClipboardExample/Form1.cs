using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace gridCopyToClipboardExample {
    public partial class Form1 : Form {

        CopyToClipboardBIFF8Helper _copyToClipboardBIFF8Helper;
        public Form1() {
            InitializeComponent();
            InitializeGrid();
            _copyToClipboardBIFF8Helper = new CopyToClipboardBIFF8Helper();
        }

        void InitializeGrid() {
            BindingList<Employee> dataList = new BindingList<Employee>();
            dataList.Add(new Employee(10115, "Augusta Delono", 1100.0, 50.0, "Accounting", new DateTime(2002, 1, 20)));
            dataList.Add(new Employee(10501, "Berry Dafoe", 1650.0, 150.0, "IT", new DateTime(2004, 5, 6)));
            dataList.Add(new Employee(10709, "Chris Cadwell", 2000.0, 180.0, "Management", new DateTime(2006, 12, 5)));
            dataList.Add(new Employee(10356, "Esta Mangold", 1400.0, 75.0, "Logistics", new DateTime(2004, 3, 12)));
            dataList.Add(new Employee(10401, "Frank Diamond", 1750.0, 100.0, "Marketing", new DateTime(2002, 4, 2)));
            dataList.Add(new Employee(10202, "Liam Bell", 1200.0, 80.0, "Manufacturing", new DateTime(2002, 4, 3)));
            dataList.Add(new Employee(10205, "Simon Newman", 1250.0, 80.0, "Manufacturing", new DateTime(2006, 5, 20)));
            dataList.Add(new Employee(10403, "Wendy Underwood", 1100.0, 50.0, "Marketing", new DateTime(2007, 9, 11)));
            this.gridControl1.DataSource = dataList;
        }

        void gridControl1_ProcessGridKey(object sender, KeyEventArgs e) {
            if(e.Control && e.KeyCode == Keys.C) {
                _copyToClipboardBIFF8Helper.CopySelectionToClipboard(gridView1);
                e.Handled = true;
            }
        }




    }

}