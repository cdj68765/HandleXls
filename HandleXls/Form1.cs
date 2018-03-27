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
using ExcelLibrary.SpreadSheet;

namespace HandleXls
{
    public sealed partial class Form1 : Form
    {
        public class BomInfo
        {
            internal string 品名;
            internal string 主件品号;
        }

        public Form1()
        {
            InitializeComponent();
            this.DoubleBuffered = true;
            Action<string> ReadXls = path =>
            {
                using (var reader = ExcelReaderFactory.CreateReader(File.Open(path, FileMode.Open)))
                {
                    foreach (DataTable item in reader.AsDataSet().Tables)
                    {
                        var DataSet = new DataGridView
                        {
                            ReadOnly = true,
                            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                            AllowUserToAddRows = false,
                            Dock = DockStyle.Fill,
                            Name = item.TableName
                        };
                        ShowTab.TabPages.Add(item.TableName, item.TableName);
                        foreach (var FirstRow in item.Rows[0].ItemArray)
                        {
                            DataSet.Columns.Add(FirstRow.ToString(), FirstRow.ToString());
                            DataSet.Columns[DataSet.Columns.Count - 1].SortMode = DataGridViewColumnSortMode.NotSortable;
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
            };
            ReadXls("2.xls");
            LoadXls.Click += delegate
            {
                var OpenFile = new OpenFileDialog
                {
                    Filter = " Excel文件 | *.xls;*.xlsx",
                    Title = "打开表格"
                };
                OpenFile.ShowDialog();
                if (File.Exists(OpenFile.FileName))
                {
                    ReadXls(OpenFile.FileName);
                }
            };
            Handler.Click += delegate
            {
                try
                {
                    int 元件品号序列 = -1;
                    int 品名序列 = -1;
                    int 主件品号序列 = -1;

                    var DataGrid = ShowTab.SelectedTab.Controls[0] as DataGridView;
                    for (int i = 0; i < DataGrid.ColumnCount; i++)
                    {
                        switch (DataGrid.Columns[i].Name)
                        {
                            case "元件品号":
                                元件品号序列 = i;
                                break;
                            case "品    名":
                                品名序列 = i;
                                break;
                            case "主件品号":
                                主件品号序列 = i;
                                break;
                        }
                    }
                    string 品名 = "";
                    var TempBomInfo = new BomInfo();
                    var SaveInfo = new Dictionary<string, List<BomInfo>>();
                    foreach (DataGridViewRow item in DataGrid.Rows)
                    {
                        if (string.IsNullOrWhiteSpace(item.Cells[元件品号序列].Value.ToString()))
                        {
                            if (品名 != "")
                            {
                                if (SaveInfo.ContainsKey(品名))
                                {
                                    SaveInfo[品名].Add(TempBomInfo);
                                }
                                else
                                {
                                    SaveInfo.Add(品名, new List<BomInfo>() { TempBomInfo });
                                }
                            }
                            TempBomInfo = new BomInfo();
                            TempBomInfo.品名 = item.Cells[品名序列].Value.ToString();
                            TempBomInfo.主件品号 = item.Cells[主件品号序列].Value.ToString();
                            品名 = "";
                        }
                        else
                        {
                            品名 += item.Cells[元件品号序列].Value + " |";
                        }
                    }
                }
                catch (Exception)
                {
                }
            };
        }
    }
}