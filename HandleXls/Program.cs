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
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            //Application.Run(new Form1());
            var Form = new Form1();
            Form.Show();
            // new Run();
        }

        private class Run
        {
            internal class BomInfo
            {
                internal string 阶级;
                internal string 品号;
                internal string 品名;
                internal string 规格;
                internal List<BomInfo> 阶级下属 = new List<BomInfo>();
            }

            internal Run()
            {
                List<BomInfo> AddNew = new List<BomInfo>();
                var Path = "FA7-10 改自动化.xlsx";

                void ReadXls(string path = "FA7-10B 改自动化.xls")
                {
                    using (var reader =
                        ExcelReaderFactory.CreateReader(new FileStream(path, FileMode.Open, FileAccess.Read,
                            FileShare.ReadWrite)))
                    {
                        DataTable Dat = reader.AsDataSet().Tables[0];
                        Dat.Rows.RemoveAt(0);
                        BomInfo Temp = null;
                        foreach (DataRow Date in Dat.Rows)
                        {
                            var Temp1 = new BomInfo();
                            switch (Date[0].ToString())
                            {
                                case "0":
                                    {
                                        if (Temp != null)
                                        {
                                            AddNew.Add(Temp);
                                        }

                                        Temp = new BomInfo();
                                        Temp.品号 = Date[22].ToString();
                                        Temp.品名 = Date[3].ToString();
                                    }
                                    break;

                                case ".1":

                                    Temp1.阶级 = Date[0].ToString();
                                    Temp1.品号 = Date[1].ToString();
                                    Temp1.品名 = Date[3].ToString();
                                    Temp1.规格 = Date[4].ToString();
                                    Temp.阶级下属.Add(Temp1);
                                    break;

                                case "..2":
                                    Temp1.阶级 = Date[0].ToString();
                                    Temp1.品号 = Date[1].ToString();
                                    Temp1.品名 = Date[3].ToString();
                                    Temp1.规格 = Date[4].ToString();
                                    Temp.阶级下属.Last().阶级下属.Add(Temp1);
                                    break;

                                case "...3":
                                    Temp1.阶级 = Date[0].ToString();
                                    Temp1.品号 = Date[1].ToString();
                                    Temp1.品名 = Date[3].ToString();
                                    Temp1.规格 = Date[4].ToString();
                                    Temp.阶级下属.Last().阶级下属.Last().阶级下属.Add(Temp1);
                                    break;
                            }
                        }
                    }
                }

                var _job = new Dictionary<string, BomInfo>();
                var Bomp = new Dictionary<string, BomInfo>();
                var Add6 = new Dictionary<string, List<BomInfo>>();
                var Cov6 = new Dictionary<string, List<BomInfo>>();
                var Cov6List = new Dictionary<string, BomInfo>();
                var Add3 = new Dictionary<string, List<BomInfo>>();
                var Cov3 = new Dictionary<string, List<BomInfo>>();
                var 常开List = new Dictionary<string, BomInfo>();
                var OFS = new List<BomInfo>();
                var Other = new List<BomInfo>();

                void Classify()
                {
                    //先得到加工件
                    AddNew.ForEach(T1 =>
                    {
                        var 加工件 = T1.阶级下属.Find(X => X.品名.IndexOf("加工件", StringComparison.Ordinal) != -1);
                        if (加工件 != null)
                        {
                            Bomp.Add(T1.品号, T1);
                            var 银点 = (from T in 加工件.阶级下属
                                      where T.品名.IndexOf("银点铆压件FA7-10A 接线柱", StringComparison.Ordinal) != -1
                                      select T.品号).FirstOrDefault();

                            if (银点 != null && !_job.ContainsKey(银点))
                            {
                                _job.Add(银点, 加工件);
                            }
                        }
                    });
                    AddNew.ForEach(T1 =>
                    {
                        if (!Bomp.ContainsKey(T1.品号))
                        {
                            var 六号基座 = T1.阶级下属.Find(X => X.品号 == "18230-0726");
                            var 三号基座 = T1.阶级下属.Find(X => X.品号 == "18230-0569");
                            if (六号基座 != null)
                            {
                                //整理思路，检索银点加工件规格，如果检索到就添加，开始操作、
                                var 银点 = (from T in T1.阶级下属
                                          where T.品名.IndexOf("银点铆压件FA7-10A 接线柱", StringComparison.Ordinal) != -1
                                          select T).FirstOrDefault();
                                if (银点 != null)
                                {
                                    //寻找加工件
                                    if (_job.ContainsKey(银点.品号))
                                    {
                                        if (Add6.ContainsKey(银点.品号))
                                        {
                                            Add6[银点.品号].Add(T1);
                                        }
                                        else
                                        {
                                            Add6.Add(银点.品号, new List<BomInfo>() { T1 });
                                        }
                                    }
                                    else
                                    {
                                        if (Cov6.ContainsKey(银点.品号))
                                        {
                                            Cov6[银点.品号].Add(T1);
                                        }
                                        else
                                        {
                                            Cov6.Add(银点.品号, new List<BomInfo>() { T1 });
                                            Cov6List.Add(银点.品号, 银点);
                                        }
                                    }
                                }
                            }
                            else if (三号基座 != null)
                            {
                                var 常闭 = (from T in T1.阶级下属
                                          where T.品名.IndexOf("银点铆压件FA7-10A 常闭", StringComparison.Ordinal) != -1
                                          select T.品号).FirstOrDefault();
                                if (常闭 != null)
                                {
                                    OFS.Add(T1);
                                }
                                else
                                {
                                    var 常开 = (from T in T1.阶级下属
                                              where T.品名.IndexOf("银点铆压件FA7-10A 常开", StringComparison.Ordinal) != -1
                                              select T).FirstOrDefault();
                                    if (常开 != null)
                                    {
                                        //寻找加工件
                                        if (_job.ContainsKey(常开.品号))
                                        {
                                            if (Add3.ContainsKey(常开.品号))
                                            {
                                                Add3[常开.品号].Add(T1);
                                            }
                                            else
                                            {
                                                Add3.Add(常开.品号, new List<BomInfo>() { T1 });
                                            }
                                        }
                                        else
                                        {
                                            if (Cov3.ContainsKey(常开.品号))
                                            {
                                                Cov3[常开.品号].Add(T1);
                                            }
                                            else
                                            {
                                                Cov3.Add(常开.品号, new List<BomInfo>() { T1 });
                                                常开List.Add(常开.品号, 常开);
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Other.Add(T1);
                            }
                        }
                    });
                }

                void DateHandle()
                {
                    XSSFWorkbook XLSXbook = new XSSFWorkbook();
                    {
                        XLSXbook.CreateSheet("不需要处理");
                        var 不需要处理 = XLSXbook.GetSheet("不需要处理");
                        不需要处理.SetColumnWidth(0, 11 * 256);
                        不需要处理.SetColumnWidth(1, 42 * 256);
                        //有6号基座加工件处理为跳过
                        var bomp = Bomp.Keys.ToArray();
                        for (int i = 0; i < bomp.Length; i++)
                        {
                            var row1 = 不需要处理.CreateRow(i);
                            row1.CreateCell(0).SetCellValue(Bomp[bomp[i]].品号);
                            row1.CreateCell(1).SetCellValue(Bomp[bomp[i]].品名);
                        }
                    }
                    {
                        //有6号基座处理为添加
                        //现有6号基座加工件
                        var bomp = Add6.Keys.ToArray();
                        for (int i = 0; i < Add6.Count; i++)
                        {
                            XLSXbook.CreateSheet($"6基座替换{i}");
                            var 替换 = XLSXbook.GetSheet($"6基座替换{i}");
                            替换.SetColumnWidth(0, 11 * 256);
                            替换.SetColumnWidth(1, 42 * 256);
                            替换.SetColumnWidth(2, 42 * 256);
                            int h = 0;
                            var row1 = 替换.CreateRow(h);
                            row1.CreateCell(1).SetCellValue("以下Bom清单做如下处理");
                            ++h;
                            row1 = 替换.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("删除以下料件");
                            ++h;
                            for (int j = 0; j < _job[bomp[i]].阶级下属.Count; j++, h++)
                            {
                                row1 = 替换.CreateRow(h);
                                row1.CreateCell(0).SetCellValue(_job[bomp[i]].阶级下属[j].品号);
                                row1.CreateCell(1).SetCellValue(_job[bomp[i]].阶级下属[j].品名);
                            }
                            ++h;
                            row1 = 替换.CreateRow(h);
                            row1.CreateCell(1).SetCellValue("添加以下加工件");
                            row1 = 替换.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue(_job[bomp[i]].品号);
                            row1.CreateCell(1).SetCellValue(_job[bomp[i]].品名);
                            ++h;
                            row1 = 替换.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("需要处理的开关品号");
                            ++h;
                            for (int j = 0; j < Add6[bomp[i]].Count; j++, h++)
                            {
                                row1 = 替换.CreateRow(h);
                                row1.CreateCell(0).SetCellValue(Add6[bomp[i]][j].品号);
                                row1.CreateCell(1).SetCellValue(Add6[bomp[i]][j].品名);
                            }
                        }
                    }
                    {
                        //创建6#基座加工件
                        var bomp = Cov6.Keys.ToArray();
                        for (int i = 0; i < Cov6.Count; i++)
                        {
                            XLSXbook.CreateSheet($"6基座添加{i}");
                            var 添加 = XLSXbook.GetSheet($"6基座添加{i}");
                            添加.SetColumnWidth(0, 11 * 256);
                            添加.SetColumnWidth(1, 42 * 256);
                            添加.SetColumnWidth(2, 42 * 256);
                            int h = 0;
                            var row1 = 添加.CreateRow(h);
                            row1.CreateCell(1).SetCellValue("以下Bom清单做如下处理");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("删除以下料件");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue(Cov6List[bomp[i]].品号);
                            row1.CreateCell(1).SetCellValue(Cov6List[bomp[i]].品名);
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue("18230-0726");
                            row1.CreateCell(1).SetCellValue("FA7-10A 6#基座");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue("5120-00054");
                            row1.CreateCell(1).SetCellValue("螺丝-(FA7-10A)");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("创建并添加以下加工件");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue($"加工件FA7-10A 6#基座({Cov6List[bomp[i]].规格})");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("需要处理的开关品号");
                            ++h;
                            for (int j = 0; j < Cov6[bomp[i]].Count; j++, h++)
                            {
                                row1 = 添加.CreateRow(h);
                                row1.CreateCell(0).SetCellValue(Cov6[bomp[i]][j].品号);
                                row1.CreateCell(1).SetCellValue(Cov6[bomp[i]][j].品名);
                            }
                        }
                    }
                    {
                        //创建3#基座加工件
                        var bomp = Cov3.Keys.ToArray();
                        for (int i = 0; i < Cov3.Count; i++)
                        {
                            XLSXbook.CreateSheet($"3基座添加{i}");
                            var 添加 = XLSXbook.GetSheet($"3基座添加{i}");
                            添加.SetColumnWidth(0, 11 * 256);
                            添加.SetColumnWidth(1, 42 * 256);
                            添加.SetColumnWidth(2, 42 * 256);
                            int h = 0;
                            var row1 = 添加.CreateRow(h);
                            row1.CreateCell(1).SetCellValue("以下Bom清单做如下处理");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("删除以下料件");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue(常开List[bomp[i]].品号);
                            row1.CreateCell(1).SetCellValue(常开List[bomp[i]].品名);
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue("18230-0569");
                            row1.CreateCell(1).SetCellValue("FA7-10A 3#基座");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(0).SetCellValue("5120-00054");
                            row1.CreateCell(1).SetCellValue("螺丝-(FA7-10A)");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("创建并添加以下加工件");
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue($"加工件FA7-10A 3#基座({常开List[bomp[i]].规格})");
                            ++h;
                            row1 = 添加.CreateRow(++h);
                            row1.CreateCell(1).SetCellValue("需要处理的开关品号");
                            ++h;
                            for (int j = 0; j < Cov3[bomp[i]].Count; j++, h++)
                            {
                                row1 = 添加.CreateRow(h);
                                row1.CreateCell(0).SetCellValue(Cov3[bomp[i]][j].品号);
                                row1.CreateCell(1).SetCellValue(Cov3[bomp[i]][j].品名);
                            }
                        }
                    }
                    using (var File = new FileStream(Path, FileMode.Create))
                    {
                        XLSXbook.Write(File);
                    }
                }

                ReadXls("FA7-10A 改自动化.xls");
                ReadXls();
                Classify();
                DateHandle();
                /*   var OpenFile = new OpenFileDialog
                   {
                       Filter = @" Excel文件 | *.xls;*.xlsx",
                       Title = @"打开表格"
                   };
                   OpenFile.ShowDialog();
                   if (File.Exists(OpenFile.FileName)) ReadXls(OpenFile.FileName);*/
            }
        }
    }
}