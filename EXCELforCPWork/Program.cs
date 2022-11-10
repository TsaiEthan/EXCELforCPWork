using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace EXCELforCPWork
{
    internal class Program
    {

        static void Main(string[] args)
        {
            for (int monthToAdd = 0; monthToAdd < 1; monthToAdd++)
            {
                DateTime date = DateTime.Now.AddMonths(monthToAdd);
                string month = DateTime.Now.AddMonths(monthToAdd).ToString("MM");
                string dirPath = @"H:\ChinPoonWork\";
                string dirPathNewFolder = dirPath + month + "月";
                string dirPathMaintenanceForm = dirPathNewFolder + @"\保養表\";
                string dirPathAppointmentMaintenanceForm = dirPathNewFolder + @"\後三月預保養表\";
                string dirPathAttachment = dirPathNewFolder + @"\附件\";

                CreateFolder(dirPathNewFolder, dirPathMaintenanceForm, dirPathAppointmentMaintenanceForm, dirPathAttachment);
                if (month.Substring(0, 1) == "0")
                    month = month.Remove(0, 1);
                //CopyFileToNewFolder(dirPath, dirPathMaintenanceForm, dirPathAppointmentMaintenanceForm, dirPathAttachment);
                DoMaintenanceFormExcelFile(dirPath, dirPathMaintenanceForm, date, month);
                DoAppointmentMaintenanceFormExcelFile(dirPath, dirPathAppointmentMaintenanceForm, date, month);
            }
            //Console.ReadLine();
        }
        static void CreateFolder(string dirPathNewFolder, string dirPathMaintenanceForm, string dirPathAppointmentMaintenanceForm, string dirPathAttachment)
        {
            //建立資料夾，以月份區分
            if (!Directory.Exists(dirPathNewFolder))
            {
                Directory.CreateDirectory(dirPathNewFolder);

                //建立保養表、後三月預保養及附件資料夾
                Directory.CreateDirectory(dirPathMaintenanceForm);
                Directory.CreateDirectory(dirPathAppointmentMaintenanceForm);
                Directory.CreateDirectory(dirPathAttachment);
                Console.WriteLine("資料夾創建成功");
            }
        }

        static void CopyFileToNewFolder(string dirPath, string dirPathMaintenanceForm, string dirPathAppointmentMaintenanceForm, string dirPathAttachment)
        {
            // 取得資料夾內所有檔案
            FileInfo[] directoryFiles = new FileInfo[] { };
            DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
            directoryFiles = directoryInfo.GetFiles("*.xls");

            //Copy原始Excel檔到新資料
            if (Directory.Exists(dirPathMaintenanceForm) && Directory.Exists(dirPathAppointmentMaintenanceForm))
            {
                foreach (FileInfo directoryFile in directoryFiles)
                {
                    System.IO.File.Copy(directoryFile.FullName, dirPathMaintenanceForm + directoryFile.Name, true);
                    System.IO.File.Copy(directoryFile.FullName, dirPathAppointmentMaintenanceForm + directoryFile.Name, true);
                }
                Console.WriteLine("保養表Copy及更名成功");
            }
        }

        static void DoMaintenanceFormExcelFile(string dirPath, string folderPath, DateTime date, string month)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo[] directoryFiles = new FileInfo[] { };
                if (Directory.Exists(dirPath))
                {
                    // 取得資料夾內所有檔案
                    DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                    directoryFiles = directoryInfo.GetFiles("*.xls");
                }
                int monthInteger = StringToInt(month);
                foreach (FileInfo directoryFile in directoryFiles)
                {
                    if (File.Exists(dirPath + directoryFile.Name))
                    {
                        FileStream file;
                        IWorkbook workBook;
                        ISheet workSheet;
                        file = new FileStream(dirPath + directoryFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                        Console.WriteLine(directoryFile.Name + "開啟成功");
                        workBook = new HSSFWorkbook(file);
                        workSheet = workBook.GetSheetAt(0);

                        HSSFCellStyle cellStyle1 = (HSSFCellStyle)workBook.CreateCellStyle();
                        //置中的Style
                        cellStyle1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        cellStyle1.VerticalAlignment = VerticalAlignment.Center;
                        IFont font = workBook.CreateFont();
                        //字型
                        font.FontName = "Times New Roman";
                        //字體尺寸
                        font.FontHeightInPoints = 16;
                        //字體粗體
                        font.IsBold = true;
                        cellStyle1.SetFont(font);

                        IFont font2 = workBook.CreateFont();
                        //字型
                        font2.FontName = "Times New Roman";
                        //字體尺寸
                        font2.FontHeightInPoints = 16;
                        //字體粗體
                        font2.IsBold = false;

                        //填入保養月份
                        workSheet.GetRow(1).GetCell(3).SetCellValue(month);
                        workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle1;

                        //填入執行日期
                        workSheet.GetRow(1).GetCell(8).SetCellValue(date.ToString("yyyy   /    M    /     "));
                        workSheet.GetRow(1).GetCell(8).CellStyle.SetFont(font2);

                        for (int i = 3; i < workSheet.LastRowNum - 1; i++)
                        {
                            //表格中增加逗號
                            if (workSheet.GetRow(i).GetCell(4).ToString() == "感測值≧500")
                            {
                                workSheet.GetRow(i).GetCell(7).SetCellValue(",");
                            }
                            string[] maintenanceMonths = workSheet.GetRow(i).GetCell(6).ToString().Split(',');
                            int x1 = 0;
                            int x2 = 0;
                            //單個月分圈起的位置
                            if (maintenanceMonths.Length == 1 && maintenanceMonths[0] == month)
                            {
                                x1 = 422;
                                x2 = 602;
                                DrowingCircle(true, workBook, workSheet, i, x1, x2);
                            }
                            else if (maintenanceMonths.Length == 2)
                            {
                                //兩個月分圈起的位置(位置1)
                                if (monthInteger <= 6 && maintenanceMonths[0] == month)
                                {
                                    x1 = 350;
                                    x2 = 530;
                                }
                                //兩個月分圈起的位置(位置2)
                                else if (monthInteger >= 7 && maintenanceMonths[1] == month)
                                {
                                    x1 = 490;
                                    x2 = 670;
                                }
                            }
                            else if (maintenanceMonths.Length == 4)
                            {
                                //四個月分圈起的位置(位置1)
                                if (monthInteger <= 3 && maintenanceMonths[0] == month)
                                {
                                    x1 = 200;
                                    x2 = 380;
                                }
                                //四個月分圈起的位置(位置2)
                                else if (monthInteger >= 4 && monthInteger <= 6 && maintenanceMonths[1] == month)
                                {
                                    x1 = 330;
                                    x2 = 510;
                                }
                                //四個月分圈起的位置(位置3)
                                else if (monthInteger >= 7 && monthInteger <= 9 && maintenanceMonths[2] == month)
                                {
                                    x1 = 440;
                                    x2 = 620;
                                }
                                //四個月分圈起的位置(位置4)
                                else if (monthInteger >= 10 && maintenanceMonths[3] == month)
                                {
                                    x1 = 590;
                                    x2 = 770;
                                }

                            }
                            //表格中圈起保養月及畫刪除線
                            if (x1 != 0 && x2 != 0)
                            {
                                DrowingCircle(true, workBook, workSheet, i, x1, x2);
                            }
                            else if (workSheet.GetRow(i).GetCell(6).ToString() != ""
                                    && workSheet.GetRow(i).GetCell(6).ToString() != "1~12")
                            {
                                DrowingLine(workSheet, i);
                            }
                        }
                        //Console.WriteLine("Excel檔案讀取完成，Sheet：" + workSheet.SheetName);
                        file = new FileStream(folderPath + directoryFile.Name, FileMode.Create, FileAccess.Write);
                        workBook.Write(file);
                        workBook.Close();
                        file.Close();
                    }
                    else
                    {
                        Console.WriteLine("Excel檔案不存在，未開啟");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel檔案開啟出錯：" + ex.Message);
            }
        }
        static void DoAppointmentMaintenanceFormExcelFile(string dirPath, string folderPath, DateTime date, string month)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo[] directoryFiles = new FileInfo[] { };
                if (Directory.Exists(dirPath))
                {
                    // 取得資料夾內所有檔案
                    DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                    directoryFiles = directoryInfo.GetFiles("*.xls");
                }
                int monthInteger = StringToInt(month);
                int[] monthAdd = new int[3] { monthInteger + 1, monthInteger + 2, monthInteger + 3};
                for (int j = 0; j < 3; j++)
                {
                    if (monthAdd[j] > 12)
                        monthAdd[j] = monthAdd[j] - 12;
                }
                string monthAddOne = (monthAdd[0]).ToString();
                string monthAddTwo = (monthAdd[1]).ToString();
                string monthAddThree = (monthAdd[2]).ToString();
                foreach (FileInfo directoryFile in directoryFiles)
                {
                    if (File.Exists(dirPath + directoryFile.Name))
                    {
                        FileStream file;
                        IWorkbook workBook;
                        ISheet workSheet;
                        file = new FileStream(dirPath + directoryFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                        Console.WriteLine(directoryFile.Name + "開啟成功");
                        workBook = new HSSFWorkbook(file);
                        workSheet = workBook.GetSheetAt(0);

                        HSSFCellStyle cellStyle1 = (HSSFCellStyle)workBook.CreateCellStyle();
                        //置中的Style
                        cellStyle1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        cellStyle1.VerticalAlignment = VerticalAlignment.Center;
                        IFont font = workBook.CreateFont();
                        //字型
                        font.FontName = "Times New Roman";
                        //字體尺寸
                        font.FontHeightInPoints = 16;
                        //字體粗體
                        font.IsBold = true;
                        cellStyle1.SetFont(font);

                        IFont font2 = workBook.CreateFont();
                        //字型
                        font2.FontName = "Times New Roman";
                        //字體尺寸
                        font2.FontHeightInPoints = 16;
                        //字體粗體
                        font2.IsBold = false;

                        //填入保養月份
                        workSheet.GetRow(1).GetCell(3).SetCellValue(month);
                        workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle1;

                        //填入預保養執行日期
                        DateTime lastWorkDate = date.AddMonths(1).AddDays(-DateTime.Now.Day);
                        if (lastWorkDate.DayOfWeek == DayOfWeek.Saturday)
                        {
                            lastWorkDate = lastWorkDate.AddDays(-1);
                        }
                        else if (lastWorkDate.DayOfWeek == DayOfWeek.Sunday)
                        {
                            lastWorkDate = lastWorkDate.AddDays(-2);
                        }
                        workSheet.GetRow(1).GetCell(8).SetCellValue(lastWorkDate.ToString("yyyy   /    M    /   dd"));
                        workSheet.GetRow(1).GetCell(8).CellStyle.SetFont(font2);

                        for (int i = 3; i < workSheet.LastRowNum - 1; i++)
                        {
                            string[] maintenanceMonths = workSheet.GetRow(i).GetCell(6).ToString().Split(',');
                            int x1 = 0;
                            int x2 = 0;
                            //單個月分圈起的位置
                            if (maintenanceMonths.Length == 1)
                            {
                                if (maintenanceMonths[0] == monthAddOne || maintenanceMonths[0] == monthAddTwo || maintenanceMonths[0] == monthAddThree)
                                {
                                    x1 = 422;
                                    x2 = 602;
                                    DrowingCircle(false, workBook, workSheet, i, x1, x2);
                                }
                            }
                            else if (maintenanceMonths.Length == 2)
                            {
                                //兩個月分圈起的位置(位置1)
                                if (maintenanceMonths[0] == monthAddOne || maintenanceMonths[0] == monthAddTwo || maintenanceMonths[0] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[0]) <= 6)
                                    {
                                        x1 = 350;
                                        x2 = 530;
                                    }
                                }
                                //兩個月分圈起的位置(位置2)
                                else if (maintenanceMonths[1] == monthAddOne || maintenanceMonths[1] == monthAddTwo || maintenanceMonths[1] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[1]) >= 7)
                                    {
                                        x1 = 490;
                                        x2 = 670;
                                    }
                                }
                            }
                            else if (maintenanceMonths.Length == 4)
                            {
                                //四個月分圈起的位置(位置1)
                                if (maintenanceMonths[0] == monthAddOne || maintenanceMonths[0] == monthAddTwo || maintenanceMonths[0] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[0]) <= 3)
                                    {
                                        x1 = 200;
                                        x2 = 380;
                                    }
                                }
                                //四個月分圈起的位置(位置2)
                                else if (maintenanceMonths[1] == monthAddOne || maintenanceMonths[1] == monthAddTwo || maintenanceMonths[1] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[1]) >= 4 && StringToInt(maintenanceMonths[1]) <= 6)
                                    {
                                        x1 = 330;
                                        x2 = 510;
                                    }
                                }
                                //四個月分圈起的位置(位置3)
                                else if (maintenanceMonths[2] == monthAddOne || maintenanceMonths[2] == monthAddTwo || maintenanceMonths[2] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[2]) >= 7 && StringToInt(maintenanceMonths[2]) <= 9)
                                    {
                                        x1 = 440;
                                        x2 = 620;
                                    }
                                }
                                //四個月分圈起的位置(位置4)
                                else if (maintenanceMonths[3] == monthAddOne || maintenanceMonths[3] == monthAddTwo || maintenanceMonths[3] == monthAddThree)
                                {
                                    if (StringToInt(maintenanceMonths[3]) >= 10)
                                    {
                                        x1 = 590;
                                        x2 = 770;
                                    }
                                }
                            }
                            //表格中圈起保養月及畫刪除線
                            if (x1 != 0 && x2 != 0)
                            {
                                DrowingCircle(false, workBook, workSheet, i, x1, x2);
                            }
                            else if (workSheet.GetRow(i).GetCell(6).ToString() != ""
                                    && workSheet.GetRow(i).GetCell(6).ToString() != "1~12")
                            {
                                DrowingLine(workSheet, i);
                            }
                        }
                        //Console.WriteLine("Excel檔案讀取完成，Sheet：" + workSheet.SheetName);
                        file = new FileStream(folderPath + directoryFile.Name, FileMode.Create, FileAccess.Write);
                        workBook.Write(file);
                        workBook.Close();
                        file.Close();
                    }
                    else
                    {
                        Console.WriteLine("Excel檔案不存在，未開啟");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel檔案開啟出錯：" + ex.Message);
            }
        }
        static void DrowingCircle(bool Maintenance, IWorkbook workBook, ISheet workSheet, int i, int x1, int x2)
        {
            //儲存格畫圈
            HSSFPatriarch patriarchCircle = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            HSSFClientAnchor c1 = new HSSFClientAnchor(x1, 30, x2, 226, 6, i, 6, i);
            HSSFSimpleShape circle1 = patriarchCircle.CreateSimpleShape(c1);
            circle1.ShapeType = HSSFSimpleShape.OBJECT_TYPE_OVAL;
            circle1.LineStyle = HSSFShape.LINESTYLE_SOLID;
            circle1.IsNoFill = true;
            circle1.LineWidth = 6350;
            //表格中增加依附件
            if (Maintenance)
            {
                if (workSheet.GetRow(i).GetCell(2).ToString() == "依校驗表"
                   || workSheet.GetRow(i).GetCell(2).ToString() == "附檢測資料")
                {
                    workSheet.GetRow(i).GetCell(7).SetCellValue("依 附 件");
                    IFont font = workBook.CreateFont();
                    //字型
                    font.FontName = "新細明體";
                    //字體尺寸
                    font.FontHeightInPoints = 12;
                    workSheet.GetRow(i).GetCell(7).CellStyle.SetFont(font);
                }
            }
        }
        static void DrowingLine(ISheet workSheet, int i)
        {
            //儲存格畫斜線
            HSSFPatriarch patriarch1 = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            HSSFClientAnchor a1 = new HSSFClientAnchor(0, 0, 0, 0, 7, i, 7 + 1, i + 1);
            HSSFSimpleShape line1 = patriarch1.CreateSimpleShape(a1);
            line1.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
            line1.LineStyle = HSSFShape.LINESTYLE_SOLID;
            // 在NPOI中線的寬度12700表示1pt,所以這裡是0.5pt粗的線條。
            line1.LineWidth = 6350;

            //儲存格畫斜線
            HSSFPatriarch patriarch2 = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            HSSFClientAnchor a2 = new HSSFClientAnchor(0, 0, 0, 0, 8, i, 8 + 1, i + 1);
            HSSFSimpleShape line2 = patriarch2.CreateSimpleShape(a2);
            line2.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
            line2.LineStyle = HSSFShape.LINESTYLE_SOLID;
            // 在NPOI中線的寬度12700表示1pt,所以這裡是0.5pt粗的線條。
            line2.LineWidth = 6350;
        }
        static int StringToInt(string stringForChange)
        {
            bool result = int.TryParse(stringForChange, out int integer);
            if (result)
                return integer;
            else
                return -1;
        }
    }
}
