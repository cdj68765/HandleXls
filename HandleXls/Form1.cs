using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDataReader;
using ExcelLibrary.SpreadSheet;

namespace HandleXls
{
    public sealed partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DoubleBuffered = true;
            导出处理后的数据();

            List<分类表> 读取并分析分类表(string Path = "分类表.xlsx")
            {
                var 分类列表 = new List<分类表>();
                using (var reader =
                    ExcelReaderFactory.CreateReader(new FileStream(Path, FileMode.Open, FileAccess.Read,
                        FileShare.ReadWrite)))
                {
                    var TempS = "";
                    var TempS2 = "";
                    foreach (DataRow Row in reader.AsDataSet().Tables[0].Rows)
                        if (string.IsNullOrWhiteSpace(Row[0].ToString()))
                        {
                            分类列表.Add(new 分类表(TempS, Row[1].ToString(), Row[2].ToString(), TempS2));
                        }
                        else
                        {
                            TempS = Row[0].ToString();
                            TempS2 = Row[3].ToString();
                            分类列表.Add(new 分类表(TempS, Row[1].ToString(), Row[2].ToString(), TempS2));
                        }
                }

                return 分类列表;
            }

            Tuple<List<object>, Dictionary<string, List<object[]>>> 读取并分析原始数据(string path = "品号信息20180417.xls")
            {
                Tuple<List<object>, Dictionary<string, List<object[]>>> Ret;
                using (var reader = ExcelReaderFactory.CreateReader(File.Open(path, FileMode.Open)))
                {
                    var Sheel = reader.AsDataSet().Tables[0].Rows;
                    Ret = new Tuple<List<object>, Dictionary<string, List<object[]>>>(Sheel[0].ItemArray.ToList(),
                        new Dictionary<string, List<object[]>>());
                    Sheel.RemoveAt(0);
                    var 分类号Index = Ret.Item1.FindIndex(x => x.ToString() == "品号");
                    foreach (DataRow Row in Sheel)
                    {
                        var 分类号 = Row[分类号Index].ToString().Split('-')[0];
                        if (Ret.Item2.ContainsKey(分类号))
                            Ret.Item2[分类号].Add(Row.ItemArray);
                        else
                            Ret.Item2.Add(分类号, new List<object[]> { Row.ItemArray });
                        //Interlocked.Increment(ref AllCount);
                    }
                }

                return Ret;
            }

            void 导出处理后的数据()
            {
                void InitSheet(ref Worksheet worksheet)
                {
                    worksheet.Cells.ColumnWidth[0] = 3200;
                    worksheet.Cells.ColumnWidth[1] = 3200;
                    worksheet.Cells.ColumnWidth[2] = 10000;
                    worksheet.Cells.ColumnWidth[3] = 10000;
                    worksheet.Cells.ColumnWidth[4] = 2000;
                    worksheet.Cells.ColumnWidth[5] = 2000;
                    worksheet.Cells.ColumnWidth[6] = 10000;
                    worksheet.Cells.ColumnWidth[7] = 2000;
                    worksheet.Cells.ColumnWidth[8] = 3200;
                    worksheet.Cells.ColumnWidth[9] = 2000;
                    worksheet.Cells.ColumnWidth[10] = 2000;
                    worksheet.Cells[0, 0] = new Cell("物料编码");
                    worksheet.Cells[1, 0] = new Cell("Id$");
                    worksheet.Cells[0, 1] = new Cell("分类（系列号）");
                    worksheet.Cells[1, 1] = new Cell("classification$<name>");
                    worksheet.Cells[0, 2] = new Cell("品名（按命名规则）");
                    worksheet.Cells[1, 2] = new Cell("name$");
                    worksheet.Cells[0, 3] = new Cell("规格（按命名规则）");
                    worksheet.Cells[1, 3] = new Cell("Specification");
                    worksheet.Cells[0, 4] = new Cell("单位");
                    worksheet.Cells[1, 4] = new Cell("UOM<description>");
                    worksheet.Cells[0, 5] = new Cell("品号属性");
                    worksheet.Cells[1, 5] = new Cell("SourceType<name>");
                    worksheet.Cells[0, 6] = new Cell("快捷码");
                    worksheet.Cells[1, 6] = new Cell("KJM");
                    worksheet.Cells[0, 7] = new Cell("主要仓库");
                    worksheet.Cells[1, 7] = new Cell("ZYCK<name>");
                    worksheet.Cells[0, 8] = new Cell("会计分类");
                    worksheet.Cells[1, 8] = new Cell("KJFL<name>");
                    worksheet.Cells[0, 9] = new Cell("备注");
                    worksheet.Cells[1, 9] = new Cell("Remark");
                    worksheet.Cells[0, 10] = new Cell("文件夹(默认即可，不输入）");
                    worksheet.Cells[1, 10] = new Cell("Folder");
                }

                /*    void 全部数据验证(Workbook book)
                    {
                        var Count = 0;
                        foreach (var Item in book.Worksheets) Count = Count + Item.Cells.Rows.Count - 1;
                        Console.WriteLine();
                    }*/

                var ClassifyList = 读取并分析分类表();
                var WorkbookDic = new Dictionary<string, Workbook>();
                var OpenXls = new OpenFileDialog();
                OpenXls.Filter = @"Excel|*.xls";
                OpenXls.ShowDialog();
                if (!File.Exists(OpenXls.FileName)) return;
                var OriData = 读取并分析原始数据(OpenXls.FileName);
                var HeaderIndex = new int[10];
                for (var i = 0; i < HeaderIndex.Length; i++)
                    switch (i)
                    {
                        case 0:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "品号");
                            break;

                        case 1:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "品名");
                            break;

                        case 2:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "规格");
                            break;

                        case 3:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "单位");
                            break;

                        case 4:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "品号属性");
                            break;

                        case 5:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "快捷码");
                            break;

                        case 6:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "主要仓库");
                            break;

                        case 7:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "会计");
                            break;

                        case 8:
                            HeaderIndex[i] = OriData.Item1.FindIndex(x => x.ToString() == "备注");
                            break;
                    }
                foreach (var Sheets in OriData.Item2)
                {
                    Workbook workbook;
                    Worksheet worksheet;
                    var Add = false;
                    var SheetName = (from Item in ClassifyList
                                     where Sheets.Key.StartsWith(Item.分类号)
                                     select Item).FirstOrDefault();

                    if (string.IsNullOrWhiteSpace(SheetName.一级))
                        continue;
                    if (SheetName.一级 == "未知分类")
                    {
                        // continue;
                    }
                    if (WorkbookDic.ContainsKey(SheetName.一级))
                    {
                        workbook = WorkbookDic[SheetName.一级];
                        worksheet = workbook.Worksheets[0];
                    }
                    else
                    {
                        Add = true;
                        workbook = new Workbook();
                        worksheet = new Worksheet(SheetName.页名);
                        InitSheet(ref worksheet);
                    }

                    var DicCount = worksheet.Cells.Rows.Count;
                    for (var i = 0; i < Sheets.Value.Count; i++, DicCount++)
                    {
                        worksheet.Cells[DicCount, 0] = new Cell(Sheets.Value[i][HeaderIndex[0]].ToString()); //品号
                        worksheet.Cells[DicCount, 1] = new Cell(Sheets.Key); //系列号
                        worksheet.Cells[DicCount, 2] = new Cell(Sheets.Value[i][HeaderIndex[1]].ToString()); //品名
                        worksheet.Cells[DicCount, 3] = new Cell(Sheets.Value[i][HeaderIndex[2]].ToString()); //规格
                        worksheet.Cells[DicCount, 4] = new Cell(Sheets.Value[i][HeaderIndex[3]].ToString()); //库存单位
                        worksheet.Cells[DicCount, 5] = new Cell(Sheets.Value[i][HeaderIndex[4]].ToString()); //品号属性
                        worksheet.Cells[DicCount, 6] = new Cell(Sheets.Value[i][HeaderIndex[5]].ToString()); //快捷码
                        worksheet.Cells[DicCount, 7] = new Cell(Sheets.Value[i][HeaderIndex[6]].ToString()); //主要仓库
                        worksheet.Cells[DicCount, 8] = new Cell(Sheets.Value[i][HeaderIndex[7]].ToString()); //会计
                        worksheet.Cells[DicCount, 9] = new Cell(HeaderIndex[8] == -1 ? "" : Sheets.Value[i][HeaderIndex[8]].ToString()); //备注缺失
                        worksheet.Cells[DicCount, 10] = new Cell(@"L:\"); //文件夹
                    }

                    if (Add)
                    {
                        workbook.Worksheets.Add(worksheet);
                        WorkbookDic.Add(SheetName.一级, workbook);
                    }
                }

                foreach (var item in WorkbookDic) item.Value.Save($"{item.Key}.xls");
                //全部数据验证(workbook);
            }

            void ReadXls(string path)
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
                            DataSet.Columns[DataSet.Columns.Count - 1].SortMode =
                                DataGridViewColumnSortMode.NotSortable;
                        }

                        item.Rows.RemoveAt(0);
                        foreach (DataRow DataBase in item.Rows) DataSet.Rows.Add(DataBase.ItemArray);

                        DataSet.Update();
                        ShowTab.TabPages[item.TableName].Controls.Add(DataSet);
                    }
                }
            }

            void SaveXls(string path, Dictionary<string, List<BomInfo>> SaveInfo)
            {
                var workbook = new Workbook();
                var count = 0;
                foreach (var BomItem in SaveInfo)
                {
                    var S1 = BomItem.Key.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
                    var S2 = S1[0].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    if (S1.Length != 1)
                    {
                        var S3 = S1[1].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        var worksheet = new Worksheet("Sheet" + count);
                        worksheet.Cells.ColumnWidth[0] = 3200;
                        worksheet.Cells.ColumnWidth[1] = 10000;
                        worksheet.Cells.ColumnWidth[2] = 10000;
                        worksheet.Cells[0, 0] = new Cell("元件组合");
                        worksheet.Cells[1, 0] = new Cell(S2[0]);
                        worksheet.Cells[1, 1] = new Cell(S2[1]);
                        worksheet.Cells[2, 0] = new Cell(S3[0]);
                        worksheet.Cells[2, 1] = new Cell(S3[1]);
                        worksheet.Cells[3, 0] = new Cell("品号");
                        worksheet.Cells[3, 1] = new Cell("品名");
                        for (var i = 0; i < BomItem.Value.Count; i++)
                        {
                            worksheet.Cells[4 + i, 0] = new Cell(BomItem.Value[i].主件品号);
                            worksheet.Cells[4 + i, 1] = new Cell(BomItem.Value[i].品名);
                        }

                        workbook.Worksheets.Add(worksheet);
                    }
                    else
                    {
                        var worksheet = new Worksheet("Sheet" + count);
                        worksheet.Cells.ColumnWidth[0] = 3200;
                        worksheet.Cells.ColumnWidth[1] = 3200;
                        worksheet.Cells.ColumnWidth[2] = 3200;
                        worksheet.Cells[0, 0] = new Cell("元件组合");
                        worksheet.Cells[1, 0] = new Cell(S2[0]);
                        worksheet.Cells[1, 1] = new Cell(S2[1]);
                        worksheet.Cells[2, 0] = new Cell("品号");
                        worksheet.Cells[2, 1] = new Cell("品名");
                        for (var i = 0; i < BomItem.Value.Count; i++)
                        {
                            worksheet.Cells[3 + i, 0] = new Cell(BomItem.Value[i].主件品号);
                            worksheet.Cells[3 + i, 1] = new Cell(BomItem.Value[i].品名);
                        }

                        workbook.Worksheets.Add(worksheet);
                    }

                    count++;
                }

                workbook.Save(path);
            }

            LoadXls.Click += delegate
            {
                var OpenFile = new OpenFileDialog
                {
                    Filter = @" Excel文件 | *.xls;*.xlsx",
                    Title = @"打开表格"
                };
                OpenFile.ShowDialog();
                if (File.Exists(OpenFile.FileName)) ReadXls(OpenFile.FileName);
            };
            Handler.Click += delegate
            {
                // try
                {
                    var 元件品号序列 = -1;
                    var 品名序列 = -1;
                    var 主件品号序列 = -1;
                    var DataGrid = ShowTab.SelectedTab.Controls[0] as DataGridView;
                    for (var i = 0; i < DataGrid.ColumnCount; i++)
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

                    var 品名 = "";
                    var TempBomInfo = new BomInfo();
                    var SaveInfo = new Dictionary<string, List<BomInfo>>();
                    foreach (DataGridViewRow item in DataGrid.Rows)
                        if (string.IsNullOrWhiteSpace(item.Cells[元件品号序列].Value.ToString()))
                        {
                            if (品名 != "")
                                if (SaveInfo.ContainsKey(品名))
                                    SaveInfo[品名].Add(TempBomInfo);
                                else
                                    SaveInfo.Add(品名, new List<BomInfo> { TempBomInfo });

                            TempBomInfo = new BomInfo
                            {
                                品名 = item.Cells[品名序列].Value.ToString(),
                                主件品号 = item.Cells[主件品号序列].Value.ToString()
                            };
                            品名 = "";
                        }
                        else
                        {
                            品名 += item.Cells[元件品号序列].Value + "|" + item.Cells[品名序列].Value + "&";
                        }

                    var SaveFile = new SaveFileDialog
                    {
                        Filter = @" Excel文件 | *.xls",
                        Title = @"保存表格"
                    };
                    SaveFile.ShowDialog();
                    if (!string.IsNullOrWhiteSpace(SaveFile.FileName)) SaveXls(SaveFile.FileName, SaveInfo);
                }
                /*  catch (Exception)
                  {
                  }*/
            };
            // this.Close();
        }

        public class BomInfo
        {
            internal string 品名;
            internal string 主件品号;
        }

        public class 分类表
        {
            internal string 二级;
            internal string 分类号;
            internal string 页名;
            internal string 一级;

            public 分类表(string v1, string v2, string v3, string v4)
            {
                一级 = v1;
                二级 = v2;
                分类号 = v3;
                页名 = v4;
            }
        }
    }
}