﻿using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
                string month = date.ToString("MM");
                //string dirPath = @"H:\ChinPoonWork\";
                string dirPath = System.IO.Directory.GetCurrentDirectory() + @"\";
                string dirPathNewFolder = dirPath + month + "月";
                string dirPathMaintenanceForm = dirPathNewFolder + @"\保養表\";
                string dirPathAppointmentMaintenanceForm = dirPathNewFolder + @"\後三月預保養表\";
                string dirPathAttachment = dirPathNewFolder + @"\附件\";
                //產生需要的資料夾
                CreateFolder(dirPathNewFolder, dirPathMaintenanceForm, dirPathAppointmentMaintenanceForm, dirPathAttachment);
                if (month.Substring(0, 1) == "0")
                    month = month.Remove(0, 1);
                //CopyFileToNewFolder(dirPath, dirPathMaintenanceForm, dirPathAppointmentMaintenanceForm, dirPathAttachment);

                //製作保養表
                DoMaintenanceFormExcelFile(dirPath, dirPathMaintenanceForm, date, month);
                //製作預保養表
                DoAppointmentMaintenanceFormExcelFile(dirPath, dirPathAppointmentMaintenanceForm, date, month);
            }
            Console.ReadLine();
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

                //獲取月份第一日及天數
                DateTime monthFirstDay;
                int daysOfMonth;
                MonthFirstDayAndDays(date, out monthFirstDay, out daysOfMonth);

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
                        //抓取表單的名子
                        string[] formName = directoryFile.Name.Split('-');

                        ICellStyle cellStyle = workBook.CreateCellStyle();
                        //置中的Style
                        cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        cellStyle.VerticalAlignment = VerticalAlignment.Center;
                        IFont font = workBook.CreateFont();
                        //字型
                        font.FontName = "Times New Roman";
                        //字體尺寸
                        font.FontHeightInPoints = 16;
                        //字體粗體
                        font.IsBold = true;
                        cellStyle.SetFont(font);

                        IFont font2 = workBook.CreateFont();
                        //字型
                        font2.FontName = "Times New Roman";
                        //字體尺寸
                        font2.FontHeightInPoints = 16;
                        //字體粗體
                        font2.IsBold = false;

                        //填入保養月份
                        workSheet.GetRow(1).GetCell(3).SetCellValue(month);
                        workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle;

                        //根據不同線別選定保養日期
                        List<DateTime> executionDate = new List<DateTime>() { };
                        switch (formName[0])
                        {
                            //DESMEAR#3，第1個星期四保
                            case "G01":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Thursday");
                                break;
                            //DESMEAR#4，第2個星期四保
                            case "G02":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Thursday");
                                break;
                            //DESMEAR#5，第3個星期四保
                            case "G03":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Thursday");
                                break;
                            //DUBURR#1，第3個星期五保
                            case "G04":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Friday");
                                break;
                            //PTH#4，第2個星期三保
                            case "G22":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Wednesday");
                                break;                                
                            //PTH#5，第2個星期一保
                            case "G05":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                break;
                            //水5，第2個星期一保
                            case "G07":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                break;
                            //PTH#6，第2個星期二保
                            case "G06":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                break;
                            //水6，第2個星期二保
                            case "G08":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                break;
                            //水7，第1個星期四保
                            case "G09":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Thursday");
                                break;
                            //水8，第1個星期一保
                            case "G10":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Monday");
                                break;
                            //水9，第2個星期四保
                            case "G11":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Thursday");
                                break;
                            //水10，第1個星期三保
                            case "G12":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Wednesday");
                                break;
                            //雷射孔微蝕#2，第3個星期二保
                            case "G13":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Tuesday");
                                break;
                            //文坦讀孔機，第2個星期日保
                            case "G18":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Sunday");
                                break;
                            //水11，第1個星期二保
                            case "G24":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Tuesday");
                                break;
                            //水12，第1個星期五保
                            case "G25":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Friday");
                                break;
                            //PLASMA，第2個星期五保
                            case "G26":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Friday");
                                break;
                        }

                        //填入執行日期
                        workSheet.GetRow(1).GetCell(8).SetCellValue(executionDate[0].ToString("yyyy   /    M    /    d"));
                        workSheet.GetRow(1).GetCell(8).CellStyle.SetFont(font2);

                        for (int i = 3; i < workSheet.LastRowNum - 1; i++)
                        {
                            if (workSheet.GetRow(i).GetCell(6) == null)
                            {
                                workSheet.GetRow(i).CreateCell(6).SetCellValue("");
                            }
                                                       
                            //表格中增加逗號
                            if (workSheet.GetRow(i).GetCell(4) != null
                                && workSheet.GetRow(i).GetCell(4).ToString() == "感測值≧500")
                            {
                                workSheet.GetRow(i).GetCell(7).SetCellValue(",");
                            }
                            string[] maintenanceMonths = workSheet.GetRow(i).GetCell(6).ToString().Split(',');
                            maintenanceMonths = workSheet.GetRow(i).GetCell(6).ToString().Split(',');
                            int x1 = 0;
                            int x2 = 0;
                            //單個月分圈起的位置
                            if (maintenanceMonths.Length == 1 && maintenanceMonths[0] == month)
                            {
                                x1 = 430;
                                x2 = 610;
                                DrowingCircle(true, workBook, workSheet, i, x1, x2, 0);
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
                                    x1 = 610;
                                    x2 = 790;
                                }
                            }
                            //表格中圈起保養月及畫刪除線
                            if (x1 != 0 && x2 != 0)
                            {
                                DrowingCircle(true, workBook, workSheet, i, x1, x2, 0);
                            }
                            else if (workSheet.GetRow(i).GetCell(6).ToString() != ""
                                    && workSheet.GetRow(i).GetCell(6).ToString() != "1~12")
                            {
                                DrowingLine(workSheet, i);
                            }
                            if(workSheet.GetRow(i).GetCell(6).ToString() != "")
                                SetCellStyle(workBook, workSheet, i);
                        }

                        //文坦讀孔機
                        if(formName[0] == "G18")
                        {                           
                            MachineCodeDrowingCircle("G19", 700, 1000, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G20", 15, 245, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G21", 315, 545, workBook, folderPath, directoryFile);
                            //For G18
                            DrowingCircle(false, workBook, workSheet, 28, 315, 615, 18);
                        }
                        //PLASMA
                        else if (formName[0] == "G26")
                        {
                            MachineCodeDrowingCircle("G28", 400, 630, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G29", 710, 940, workBook, folderPath, directoryFile);
                            //For G26
                            DrowingCircle(false, workBook, workSheet, 28, 80, 310, 26);
                            DrowingCircle(false, workBook, workSheet, 1, 315, 373, 26);
                        }
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
        static void MonthFirstDayAndDays(DateTime date, out DateTime monthFirstDay, out int daysOfMonth)
        {
            monthFirstDay = date.AddDays(-DateTime.Now.Day + 1);
            DateTime monthLastDay = date.AddMonths(1).AddDays(-DateTime.Now.Day);
            //兩時間天數相減
            TimeSpan ts = monthLastDay.Subtract(monthFirstDay);
            //相距天數
            daysOfMonth = ts.Days;
        }

        static List<DateTime> DateToWeekDay(DateTime monthFirstDay, int daysOfMonth, int whichWeek, string whatDayIsIt)
        {
            List<DateTime> executionDate = new List<DateTime>() { };
            executionDate.Add(new DateTime());
            for (int i = 0; i <= daysOfMonth; i++)
            {
                if (monthFirstDay.AddDays(i).DayOfWeek.ToString() == whatDayIsIt)
                {
                    executionDate.Add(monthFirstDay.AddDays(i));
                }
            }
            executionDate[0] = executionDate[whichWeek];
            return executionDate;
        }
        static void SetCellStyle(IWorkbook workBook, ISheet workSheet, int i)
        {
            ICellStyle cellStyleOriginal = workBook.CreateCellStyle();
            //置中的Style
            cellStyleOriginal.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyleOriginal.VerticalAlignment = VerticalAlignment.Center;
            //下邊框
            cellStyleOriginal.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            IFont fontOriginal = workBook.CreateFont();
            //字型
            fontOriginal.FontName = "Times New Roman";
            //字體尺寸
            fontOriginal.FontHeightInPoints = 14;
            //字體粗體
            fontOriginal.IsBold = false;
            cellStyleOriginal.SetFont(fontOriginal);
            workSheet.GetRow(i).GetCell(6).CellStyle = cellStyleOriginal;
        }
        static void MachineCodeDrowingCircle(string machineCode, int x3, int x4, IWorkbook workBook, string folderPath, FileInfo directoryFile)
        {
            int machineCodeNumber = StringToInt(machineCode.Substring(1,2));
            ISheet workSheet = workBook.GetSheetAt(0);
            HSSFPatriarch circle = DrowingCircle(false, workBook, workSheet, 28, x3, x4, machineCodeNumber);
            HSSFPatriarch circle2 = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            if (machineCodeNumber == 28)
                circle2 = DrowingCircle(false, workBook, workSheet, 1, 370, 428, machineCodeNumber);
            else if(machineCodeNumber == 29)
                circle2 = DrowingCircle(false, workBook, workSheet, 1, 425, 483, machineCodeNumber);
            FileStream newFile = new FileStream(folderPath + machineCode + "-" + directoryFile.Name, FileMode.Create, FileAccess.Write);
            workBook.Write(newFile);
            circle.Clear();
            if (machineCodeNumber == 28 || machineCodeNumber == 29) 
                circle2.Clear();
        }
        static void DoCurrentCheckForm(FileInfo directoryFile)
        {
            //directoryFile.Name;
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

                        ICellStyle cellStyle2 = workBook.CreateCellStyle();
                        //置中的Style
                        cellStyle2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        cellStyle2.VerticalAlignment = VerticalAlignment.Center;
                        IFont font = workBook.CreateFont();
                        //字型
                        font.FontName = "Times New Roman";
                        //字體尺寸
                        font.FontHeightInPoints = 16;
                        //字體粗體
                        font.IsBold = true;
                        cellStyle2.SetFont(font);

                        IFont font2 = workBook.CreateFont();
                        //字型
                        font2.FontName = "Times New Roman";
                        //字體尺寸
                        font2.FontHeightInPoints = 16;
                        //字體粗體
                        font2.IsBold = false;

                        //填入保養月份
                        workSheet.GetRow(1).GetCell(3).SetCellValue(monthAddOne + "、" + monthAddTwo + "、" + monthAddThree);
                        workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle2;

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
                            if (workSheet.GetRow(i).GetCell(6) == null)
                            {
                                workSheet.GetRow(i).CreateCell(6).SetCellValue("");
                            }
                            string[] maintenanceMonths = workSheet.GetRow(i).GetCell(6).ToString().Split(',');
                            int x1 = 0;
                            int x2 = 0;
                            //單個月分圈起的位置
                            if (maintenanceMonths.Length == 1)
                            {
                                if (maintenanceMonths[0] == monthAddOne || maintenanceMonths[0] == monthAddTwo || maintenanceMonths[0] == monthAddThree)
                                {
                                    x1 = 430;
                                    x2 = 610;
                                    DrowingCircle(false, workBook, workSheet, i, x1, x2, 0);
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
                                        x1 = 610;
                                        x2 = 790;
                                    }
                                }
                            }
                            //表格中圈起保養月及畫刪除線
                            if (x1 != 0 && x2 != 0)
                            {
                                DrowingCircle(false, workBook, workSheet, i, x1, x2, 0);
                            }
                            else if (workSheet.GetRow(i).GetCell(6).ToString() != ""
                                    && workSheet.GetRow(i).GetCell(6).ToString() != "1~12")
                            {
                                DrowingLine(workSheet, i);
                            }
                            if (workSheet.GetRow(i).GetCell(6).ToString() != "")
                                SetCellStyle(workBook, workSheet, i);
                        }
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
        static HSSFPatriarch DrowingCircle(bool Maintenance, IWorkbook workBook, ISheet workSheet, int i, int x1, int x2, int machineCodeNumber)
        {
            int initial = 6;
            if(machineCodeNumber == 18 || machineCodeNumber == 19)
            {
                initial = 8;
            }
            else if (machineCodeNumber == 20 || machineCodeNumber == 21
                     || machineCodeNumber == 26 || machineCodeNumber == 28 || machineCodeNumber == 29)
            {
                if(i == 28)
                    initial = 9;
                else
                    initial = 1;
            }
            //儲存格畫圈
            HSSFPatriarch patriarchCircle = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            HSSFClientAnchor c1 = new HSSFClientAnchor(x1, 30, x2, 226, initial, i, initial, i);
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
            return patriarchCircle;
        }
        static void DrowingLine(ISheet workSheet, int i)
        {
            //儲存格畫斜線
            for (int j = 7; j < 9; j++)
            {
                HSSFPatriarch patriarch1 = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
                HSSFClientAnchor a1 = new HSSFClientAnchor(0, 0, 0, 0, j, i, j + 1, i + 1);
                HSSFSimpleShape line1 = patriarch1.CreateSimpleShape(a1);
                line1.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                line1.LineStyle = HSSFShape.LINESTYLE_SOLID;
                // 在NPOI中線的寬度12700表示1pt,所以這裡是0.5pt粗的線條。
                line1.LineWidth = 6350;
            }
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
