using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelDataReader;
using ExcelLibrary.SpreadSheet;
using NPOI.XSSF.UserModel;

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
                            Ret.Item2.Add(分类号, new List<object[]> {Row.ItemArray});
                        //Interlocked.Increment(ref AllCount);
                    }
                }

                return Ret;
            }

            void 导出处理后的数据()
            {
                /*    void 全部数据验证(Workbook book)
                    {
                        var Count = 0;
                        foreach (var Item in book.Worksheets) Count = Count + Item.Cells.Rows.Count - 1;
                        Console.WriteLine();
                    }*/

                var ClassifyList = 读取并分析分类表();
                var WorkbookDic = new Dictionary<string, XSSFWorkbook>();
                var WorkbookIndex = new Dictionary<string, int>();
                var OpenXls = new OpenFileDialog {Filter = @"Excel|*.xls"};
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
                    XSSFWorkbook XLSXbook;
                    XSSFSheet XLSXSheet;
                    var Add = false;
                    var SheetName = (from Item in ClassifyList
                        where Sheets.Key.StartsWith(Item.分类号)
                        select Item).FirstOrDefault();

                    if (string.IsNullOrWhiteSpace(SheetName.一级))
                        continue;
                    int DicCount;
                    if (WorkbookDic.ContainsKey(SheetName.一级))
                    {
                        XLSXbook = WorkbookDic[SheetName.一级];
                        DicCount = WorkbookIndex[SheetName.一级];
                        XLSXSheet = XLSXbook.GetSheetAt(0) as XSSFSheet;
                    }
                    else
                    {
                        Add = true;
                        XLSXbook = new XSSFWorkbook();
                        XLSXbook.CreateSheet(SheetName.页名);
                        XLSXSheet = XLSXbook.GetSheetAt(0) as XSSFSheet;
                        InitXSheet();
                        DicCount = 1;

                        void InitXSheet()
                        {
                            var row1 = XLSXSheet.CreateRow(0);
                            var row2 = XLSXSheet.CreateRow(1);
                            XLSXSheet.SetColumnWidth(0, 11 * 256);
                            XLSXSheet.SetColumnWidth(1, 23 * 256);
                            XLSXSheet.SetColumnWidth(2, 36 * 256);
                            XLSXSheet.SetColumnWidth(3, 46 * 256);
                            XLSXSheet.SetColumnWidth(4, 5 * 256);
                            XLSXSheet.SetColumnWidth(5, 8 * 256);
                            XLSXSheet.SetColumnWidth(6, 22 * 256);
                            XLSXSheet.SetColumnWidth(7, 11 * 256);
                            XLSXSheet.SetColumnWidth(8, 11 * 256);
                            XLSXSheet.SetColumnWidth(9, 8 * 256);
                            row1.CreateCell(0).SetCellValue("物料编码");
                            row2.CreateCell(0).SetCellValue("Id$");
                            row1.CreateCell(1).SetCellValue("分类（系列号）");
                            row2.CreateCell(1).SetCellValue("classification$<name>");
                            row1.CreateCell(2).SetCellValue("品名（按命名规则）");
                            row2.CreateCell(2).SetCellValue("name$");
                            row1.CreateCell(3).SetCellValue("规格（按命名规则）");
                            row2.CreateCell(3).SetCellValue("Specification");
                            row1.CreateCell(4).SetCellValue("单位");
                            row2.CreateCell(4).SetCellValue("UOM<description>");
                            row1.CreateCell(5).SetCellValue("品号属性");
                            row2.CreateCell(5).SetCellValue("SourceType<name>");
                            row1.CreateCell(6).SetCellValue("快捷码");
                            row2.CreateCell(6).SetCellValue("KJM");
                            row1.CreateCell(7).SetCellValue("主要仓库");
                            row2.CreateCell(7).SetCellValue("ZYCK<name>");
                            row1.CreateCell(8).SetCellValue("会计分类");
                            row2.CreateCell(8).SetCellValue("KJFL<name>");
                            row1.CreateCell(9).SetCellValue("备注");
                            row2.CreateCell(9).SetCellValue("Remark");
                            row1.CreateCell(10).SetCellValue("文件夹(默认即可，不输入）");
                            row2.CreateCell(10).SetCellValue("Folder");
                        }
                    }

                    for (var i = 0; i < Sheets.Value.Count; i++, DicCount++)
                    {
                        var row1 = XLSXSheet.CreateRow(DicCount);
                        row1.CreateCell(0).SetCellValue(Sheets.Value[i][HeaderIndex[0]].ToString());
                        row1.CreateCell(1).SetCellValue(Sheets.Key);
                        row1.CreateCell(2).SetCellValue(Sheets.Value[i][HeaderIndex[1]].ToString());
                        row1.CreateCell(3).SetCellValue(Sheets.Value[i][HeaderIndex[2]].ToString());
                        row1.CreateCell(4).SetCellValue(Sheets.Value[i][HeaderIndex[3]].ToString());
                        row1.CreateCell(5).SetCellValue(Sheets.Value[i][HeaderIndex[4]].ToString());
                        row1.CreateCell(6).SetCellValue(Sheets.Value[i][HeaderIndex[5]].ToString());
                        row1.CreateCell(7).SetCellValue(Sheets.Value[i][HeaderIndex[6]].ToString());
                        row1.CreateCell(8).SetCellValue(Sheets.Value[i][HeaderIndex[7]].ToString());
                        row1.CreateCell(9)
                            .SetCellValue(HeaderIndex[8] == -1 ? "" : Sheets.Value[i][HeaderIndex[8]].ToString());
                        row1.CreateCell(10).SetCellValue(@"L:\");
                    }

                    if (Add)
                    {
                        WorkbookDic.Add(SheetName.一级, XLSXbook);
                        WorkbookIndex.Add(SheetName.一级, DicCount);
                    }
                    else
                    {
                        WorkbookIndex[SheetName.一级] = DicCount;
                    }
                }

                foreach (var VARIABLE in WorkbookDic)
                    using (var File = new FileStream($"{VARIABLE.Key}.xlsx", FileMode.Create))
                    {
                        VARIABLE.Value.Write(File);
                    }

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
                    var S1 = BomItem.Key.Split(new[] {'&'}, StringSplitOptions.RemoveEmptyEntries);
                    var S2 = S1[0].Split(new[] {'|'}, StringSplitOptions.RemoveEmptyEntries);
                    if (S1.Length != 1)
                    {
                        var S3 = S1[1].Split(new[] {'|'}, StringSplitOptions.RemoveEmptyEntries);
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
                                    SaveInfo.Add(品名, new List<BomInfo> {TempBomInfo});

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