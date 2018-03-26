using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;

namespace HandleXls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
            LoadXls.Click += delegate 
            {
               var OpenFile=new  OpenFileDialog();
                OpenFile.Filter =" Excel文件 | *.xls;*.xlsx";
                OpenFile.Title = "打开表格";
                OpenFile.ShowDialog();
                if (File.Exists(OpenFile.FileName))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(File.Open(OpenFile.FileName, FileMode.Open)))
                    {
                        foreach (DataTable item in reader.AsDataSet().Tables)
                        {
                            var DataSet = new DataGridView();
                            DataSet.ReadOnly = true;
                            DataSet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                            DataSet.AllowUserToAddRows = false;
                            DataSet.Dock = DockStyle.Fill;
                            DataSet.Name = item.TableName;
                            ShowTab.TabPages.Add(item.TableName, item.TableName);
                            foreach (var FirstRow in item.Rows[0].ItemArray)
                            {
                                DataSet.Columns.Add(FirstRow.ToString(), FirstRow.ToString());
                                DataSet.Columns[DataSet.Columns.Count-1].SortMode = DataGridViewColumnSortMode.NotSortable;
                            }
                            item.Rows.RemoveAt(0);
                            foreach (DataRow DataBase in item.Rows)
                            {
                                DataSet.Rows.Add(DataBase.ItemArray);
                            }
                            DataSet.Update();
                            ShowTab.TabPages[item.TableName].Controls.Add(DataSet);
                        }
                    }
                }
            };
            Handler.Click += delegate 
            {
                try
                {
                    var DataGrid = ShowTab.SelectedTab.Controls[0] as DataGridView;
                    foreach (DataGridViewRow item in DataGrid.Rows)
                    {
                        Console.WriteLine(item);
                        item.Cells[0].OwningColumn=
                           item.Cells[0].Value =
                    }
                }
                catch (Exception)
                {

                }
             
            };

        }
    }
}
