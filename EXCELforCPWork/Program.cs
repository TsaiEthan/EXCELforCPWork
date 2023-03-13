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
        static int a01Count = 0, maintenanceCount = 0;
        static void Main(string[] args)
        {
            for (int monthToAdd = 0; monthToAdd < 1; monthToAdd++)
            {
                DateTime date = DateTime.Now.AddMonths(monthToAdd);
                string yearMonth = date.ToString("yyyy" + "年" + "MM" + "月");
                string month = date.ToString("MM");
                string dirPath = System.IO.Directory.GetCurrentDirectory() + @"\";
                string newDirPath = dirPath + yearMonth + @"\";

                //產生需要的資料夾
                CreateFolder(newDirPath);
                if (month.Substring(0, 1) == "0")
                    month = month.Remove(0, 1);

                //製作保養表及產生相關附件
                DoMaintenanceFormExcelFile(dirPath, newDirPath, date, month);

                //刪除已製作好的附件內空白的Sheet
                List<string> fileName = new List<string> { "A02~A06-設備性能檢測數值記錄表.xls", "A07~A08電流比對紀錄表.xls" };
                RemoveBlankExcelSheet(newDirPath, fileName);

                //製作預保養表
                DoAppointmentMaintenanceFormExcelFile(dirPath, newDirPath, date, month);

                //製作LAYOUT圖(緊急開關、液位開關)
                DoLayoutFormExcelFile(dirPath, newDirPath, month);
            }
            Console.ReadLine();
        }
        static void RemoveBlankExcelSheet(string newDirPath, List<string> fileName)
        {
            foreach (string cunrrentFile in fileName)
            {
                //讀取原始檔
                if (File.Exists(newDirPath + cunrrentFile))
                {
                    FileStream file = new FileStream(newDirPath + cunrrentFile, FileMode.Open, FileAccess.Read);
                    IWorkbook workBook = new HSSFWorkbook(file);
                    file.Close();
                    int sheetCount = 0;
                    if (cunrrentFile == "A02~A06-設備性能檢測數值記錄表.xls")
                        sheetCount = 6;
                    else if (cunrrentFile == "A07~A08電流比對紀錄表.xls")
                        sheetCount = 2;
                    for(int i = 0; i < sheetCount; i++)
                        workBook.RemoveSheetAt(0);
                    if(workBook.NumberOfSheets == 0)
                        workBook.CreateSheet("此月無需此附件");
                    workBook.SetActiveSheet(0);
                    file = new FileStream(newDirPath + cunrrentFile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    workBook.Write(file, true);
                    workBook.Close();
                    file.Close();
                }
            }
        }
        static void CreateFolder(string newDirPath)
        {
            //建立資料夾，以月份區分
            if (!Directory.Exists(newDirPath))
            {
                Directory.CreateDirectory(newDirPath);
                string[] folderName = newDirPath.Split(@"\");
                Console.WriteLine(folderName[folderName.Length - 2] + " 資料夾創建成功");
            }
        }
        static void DoMaintenanceFormExcelFile(string dirPath, string newDirPath, DateTime date, string month)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo directoryGFile = null;                
                if (File.Exists(dirPath + "G01~G26-設備定期保養項目表.xls"))
                {
                    directoryGFile = new FileInfo(dirPath + "G01~G26-設備定期保養項目表.xls");
                }

                int monthInteger = StringToInt(month);

                //獲取月份第一日及天數
                DateTime monthFirstDay;
                int daysOfMonth;
                MonthFirstDayAndDays(date, out monthFirstDay, out daysOfMonth);
                //獲取下個月份第一日
                DateTime nextMonthFirstDay;
                NextMonthFirstDay(date, out nextMonthFirstDay);
                //複製檔案
                //讀取原始檔
                FileStream file = new FileStream(directoryGFile.FullName, FileMode.Open, FileAccess.Read);
                IWorkbook workBook = new HSSFWorkbook(file);
                file.Close();

                //複製Sheet文坦，並放到正確位置
                for (int i = 1; i <= 3; i++)
                {
                    int machineCode = 18 + i;
                    ISheet newWorkSheet = workBook.CloneSheet(13);
                    workBook.SetSheetName(workBook.NumberOfSheets - 1, "G" + machineCode.ToString() + "- 文坦");
                    workBook.SetSheetOrder("G" + machineCode.ToString() + "- 文坦", 13 + i);
                }
                //複製Sheet PLASMA，並放到正確位置
                for (int i = 2; i <= 3; i++)
                {
                    int machineCode = 26 + i;
                    if (machineCode != 27)
                    {
                        ISheet newWorkSheet = workBook.CloneSheet(20);
                        workBook.SetSheetName(workBook.NumberOfSheets - 1, "G" + machineCode.ToString() + "- PLASMA");
                        workBook.SetSheetOrder("G" + machineCode.ToString() + "- PLASMA", 20 + i - 1);
                    }
                }
                file = new FileStream(newDirPath + directoryGFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                file.Close();

                file = new FileStream(newDirPath + "後三月預保養表.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                file.Close();

                FileInfo directoryAFile = null;
                if (File.Exists(dirPath + "A02~A06-設備性能檢測數值記錄表.xls"))
                {
                    directoryAFile = new FileInfo(dirPath + "A02~A06-設備性能檢測數值記錄表.xls");
                    file = new FileStream(directoryAFile.FullName, FileMode.Open, FileAccess.Read);
                    IWorkbook copyWorkBook = new HSSFWorkbook(file);
                    file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    copyWorkBook.Write(file, true);
                    copyWorkBook.Close();
                    file.Close();
                }

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
                for (int i = 0; i < workBook.NumberOfSheets; i++)
                {
                    ISheet workSheet = workBook.GetSheetAt(i);

                    //填入保養月份
                    workSheet.GetRow(1).GetCell(3).SetCellValue(month);
                    workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle;

                    bool heaterCheck = false;
                    HSSFSimpleShape circle1;
                    //圈保養月份及劃刪除線
                    for (int j = 3; j < workSheet.LastRowNum - 1; j++)
                    {
                        int x1 = 0;
                        int x2 = 0;
                        if (workSheet.GetRow(j).GetCell(6) == null)
                        {
                            workSheet.GetRow(j).CreateCell(6).SetCellValue("");
                        }

                        //表格中增加逗號
                        if (workSheet.GetRow(j).GetCell(4) != null
                            && workSheet.GetRow(j).GetCell(4).ToString() == "感測值≧500")
                        {
                            workSheet.GetRow(j).GetCell(7).SetCellValue(",");
                        }
                        string[] maintenanceMonths = workSheet.GetRow(j).GetCell(6).ToString().Split(',');
                        maintenanceMonths = workSheet.GetRow(j).GetCell(6).ToString().Split(',');

                        if(workSheet.GetRow(j).GetCell(6).ToString() == "1~12")
                        {
                            maintenanceCount++;
                        }

                        //單個月分圈起的位置
                        if (maintenanceMonths.Length == 1 && maintenanceMonths[0] == month)
                        {
                            x1 = 430;
                            x2 = 610;
                            DrowingCircle(true, workBook, workSheet, j, x1, x2, 0);
                            maintenanceCount++;
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
                            DrowingCircle(true, workBook, workSheet, j, x1, x2, 0, out circle1, ref heaterCheck);
                            maintenanceCount++;
                        }
                        else if (workSheet.GetRow(j).GetCell(6).ToString() != ""
                                && workSheet.GetRow(j).GetCell(6).ToString() != "1~12")
                        {
                            DrowingLine(workSheet, j);
                        }
                        if (workSheet.GetRow(j).GetCell(6).ToString() != "")
                            SetCellStyle(workBook, workSheet, j , 6, 14, 1);
                    }

                    //抓取表單的名子
                    string[] formName = workSheet.SheetName.Split('-');
                    //根據不同線別選定保養日期
                    List<DateTime> executionDate = new List<DateTime>() { };
                    List<DateTime> nextMonthExecutionDate = new List<DateTime>() { };
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
                        //DESMEAR#5，第2個星期五保
                        case "G03":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Friday");
                            break;
                        //DEBURR#1，第3個星期五保
                        case "G04":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Friday");
                            if (heaterCheck)
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G04", "DEBURR#1");
                            break;
                        //PTH#4，第2個星期三保
                        case "G22":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Wednesday");
                            if (heaterCheck)
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G22", "PTH#4");
                            break;
                        //PTH#5，第2個星期一保
                        case "G05":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                            if (heaterCheck)
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G05", "PTH#5");
                            break;
                        //水5，第2個星期一保
                        case "G07":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G07", "水5");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G07", "水5");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G07", "水5");
                            }
                            break;
                        //PTH#6，第2個星期二保
                        case "G06":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G06", "PTH#6");
                            }
                            break;
                        //水6，第2個星期二保
                        case "G08":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G08", "水6");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G08", "水6");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G08", "水6");
                            }
                            break;
                        //水7，第1個星期四保
                        case "G09":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Thursday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G09", "水7");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G09", "水7");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G09", "水7");
                            }
                            break;
                        //水8，第1個星期一保
                        case "G10":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Monday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G10", "水8");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G10", "水8");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G10", "水8A");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G10", "水8B");
                            }
                            break;
                        //水9，第2個星期四保
                        case "G11":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Thursday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G11", "水9");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G11", "水9");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G11", "水9");
                            }
                            break;
                        //水10，第1個星期三保
                        case "G12":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Wednesday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G12", "水10");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G12", "水10");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Wednesday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G12", "水10A");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G12", "水10B");
                            }
                            break;
                        //雷射孔微蝕#2，第3個星期二保
                        case "G13":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Tuesday");
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G13", "雷射孔微蝕#2");
                            }
                            break;
                        //文坦讀孔機，第2個星期二保
                        case "G18":
                        case "G19":
                        case "G20":
                        case "G21":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                            break;
                        //水11，第1個星期二保
                        case "G24":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Tuesday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G24", "水11");
                            a01Count++;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G24", "水11");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G24", "水11");
                            }
                            break;
                        //水12，第1個星期五保
                        case "G25":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Friday");
                            DoForm_A01(dirPath, newDirPath, executionDate, "G25", "水12");
                            a01Count = 0;
                            if (heaterCheck)
                            {
                                DoForm_A02ToA06(dirPath, newDirPath, executionDate, "G25", "水12");
                            }
                            if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                            {
                                nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Friday");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G25", "水12A");
                                DoForm_A07A08(dirPath, newDirPath, nextMonthExecutionDate, "G25", "水12B");
                            }
                            break;
                        //PLASMA，第2個星期五保
                        case "G26":
                        case "G28":
                        case "G29":
                            executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Friday");
                            break;
                    }

                    //填入執行日期
                    workSheet.GetRow(1).GetCell(8).SetCellValue(executionDate[0].ToString("yyyy   /    M    /    d"));
                    workSheet.GetRow(1).GetCell(8).CellStyle.SetFont(font2);

                    //圈取設備代碼
                    CircleMachineCode(formName[0], workBook, workSheet);

                    SetPrintStyle(workSheet);
                    workSheet.GetRow(0).CreateCell(25).SetAsActiveCell();
                    workBook.SetActiveSheet(0);
                    file = new FileStream(newDirPath + directoryGFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    workBook.Write(file, true);
                    file.Close();
                }
                Console.WriteLine("寫入 " + GetFileName(file.Name) + " 成功");
                workBook.Close();

                //紀錄保養數量                
                workBook = new HSSFWorkbook();
                ISheet recordSheet = workBook.CreateSheet("數量統計");
                recordSheet.CreateRow(0).CreateCell(0).SetCellValue("本月保養數量");
                //將Column 0，欄寬設定為12
                recordSheet.SetColumnWidth(0, (int)((12 + 0.71) * 256));
                SetCellStyle(workBook, recordSheet, 0, 0, 10, 2);
                recordSheet.GetRow(0).CreateCell(1).SetCellValue(maintenanceCount);
                SetCellStyle(workBook, recordSheet, 0, 1, 10, 2);
                file = new FileStream(newDirPath + "數量統計.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                workBook.Close();
                file.Close();
                Console.WriteLine("寫入 " + GetFileName(file.Name) + " 成功");
                maintenanceCount = 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel檔案開啟出錯：" + ex.Message);
            }
        }
        static void CircleMachineCode(string machineCode, IWorkbook workBook, ISheet workSheet)
        {
            //圈取設備代碼
            switch (machineCode)
            {
                //文坦讀孔機
                case "G18":
                    DrowingCircle(false, workBook, workSheet, 28, 250, 550, 18);
                    break;
                case "G19":
                    DrowingCircle(false, workBook, workSheet, 28, 640, 940, 19);
                    break;
                case "G20":
                    DrowingCircle(false, workBook, workSheet, 28, 8, 238, 20);
                    break;
                case "G21":
                    DrowingCircle(false, workBook, workSheet, 28, 310, 540, 21);
                    break;
                //PLASMA
                case "G26":
                    DrowingCircle(false, workBook, workSheet, 28, 99, 328, 26);
                    DrowingCircle(false, workBook, workSheet, 1, 275, 332, 26);
                    break;
                case "G28":
                    DrowingCircle(false, workBook, workSheet, 28, 398, 625, 28);
                    DrowingCircle(false, workBook, workSheet, 1, 324, 380, 28);
                    break;
                case "G29":
                    DrowingCircle(false, workBook, workSheet, 28, 702, 932, 29);
                    DrowingCircle(false, workBook, workSheet, 1, 373, 428, 29);
                    break;
            }
        }

        static void SetPrintStyle(ISheet workSheet)
        {
            //設定列印邊界，0.2=0.5CM
            workSheet.SetMargin(NPOI.SS.UserModel.MarginType.TopMargin, 0.2);
            workSheet.SetMargin(NPOI.SS.UserModel.MarginType.RightMargin, 0.2);
            workSheet.SetMargin(NPOI.SS.UserModel.MarginType.BottomMargin, 0.2);
            workSheet.SetMargin(NPOI.SS.UserModel.MarginType.LeftMargin, 0.2);
            //水平置中
            workSheet.HorizontallyCenter = true;
            //垂直置中
            workSheet.VerticallyCenter = true;
        }
        static string GetFileName(string fullFileName)
        {
            string[] fileNameWithoutPath = fullFileName.Split('\\');
            return fileNameWithoutPath[fileNameWithoutPath.Length - 1];
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

        static void NextMonthFirstDay(DateTime date, out DateTime nextMonthFirstDay)
        {
            date = date.AddMonths(1);
            nextMonthFirstDay = date.AddDays(-DateTime.Now.Day + 1);
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
        static void SetCellStyle(IWorkbook workBook, ISheet workSheet, int row, int column, int fontHeightInPoints, int allBorder)
        {
            ICellStyle cellStyleOriginal = workBook.CreateCellStyle();
            //置中的Style
            cellStyleOriginal.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyleOriginal.VerticalAlignment = VerticalAlignment.Center;
            if(allBorder == 0)
            {
                //上邊框
                cellStyleOriginal.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                //左邊框
                cellStyleOriginal.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                //右邊框
                cellStyleOriginal.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                //下邊框
                cellStyleOriginal.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            }
            else if(allBorder == 1)
            {
                //下邊框
                cellStyleOriginal.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            }

            IFont fontOriginal = workBook.CreateFont();
            //字型
            fontOriginal.FontName = "Times New Roman";
            //字體尺寸
            fontOriginal.FontHeightInPoints = fontHeightInPoints;
            //字體粗體
            fontOriginal.IsBold = false;
            cellStyleOriginal.SetFont(fontOriginal);
            workSheet.GetRow(row).GetCell(column).CellStyle = cellStyleOriginal;
        }
        static void DoForm_A01(string dirPath, string newDirPath, List<DateTime> executionDate, string machineCode, string lineName)
        {
            FileInfo directoryAFile = new FileInfo(dirPath + "A01-亞碩競銘線纜線熱顯像檢查表.xls");
            FileStream file = null;
            IWorkbook workBook = null;
            //複製檔案
            if (!File.Exists(newDirPath + directoryAFile.Name))
            {
                file = new FileStream(directoryAFile.FullName, FileMode.Open, FileAccess.Read);
                workBook = new HSSFWorkbook(file);
                file.Close();
                file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                file.Close();
            }
            
            file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workBook = new HSSFWorkbook(file);
            file.Close();
            ISheet workSheet = workBook.GetSheetAt(0);
            ISheet newWorkSheet;
            if (workSheet.SheetName != "熱顯像檢查表")
            {
                newWorkSheet = workBook.CloneSheet(0);
                workBook.SetSheetName(a01Count, lineName);
                workBook.SetActiveSheet(a01Count);
            }
            else
            {
                newWorkSheet = workBook.GetSheetAt(0);
                workBook.SetSheetName(0, lineName);
            }
            newWorkSheet.GetRow(0).GetCell(0).SetCellValue(lineName);

            if (lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                int j = 1;
                for (int i = 1; i < executionDate.Count; i++)
                {
                    int ramdomCurrentA = 0, ramdomCurrentB = 0, ramdomTemperatureA = 0, ramdomTemperatureB = 0, addTemperature = 3;
                    //亂數電流，介於540~1480
                    ramdomCurrentA = RandomCurrent(newWorkSheet, 540, 1480, 0, 0, false);
                    ramdomTemperatureA = GetTemperature(ramdomCurrentA);
                    var random = new Random();
                    if (random.NextDouble() >= 0.5)
                    {
                        ramdomCurrentB = ramdomCurrentA;
                        ramdomTemperatureB = ramdomTemperatureA;
                    }
                    else
                    {
                        //亂數電流，介於540~1480
                        ramdomCurrentB = RandomCurrent(newWorkSheet, 540, 1480, 0, 0, false);
                        ramdomTemperatureB = GetTemperature(ramdomCurrentB);
                    }
                    if (ramdomTemperatureA == 37)
                        addTemperature = 2;
                    newWorkSheet.GetRow(j + 1).GetCell(0).SetCellValue(executionDate[i].Year);
                    newWorkSheet.GetRow(j + 1).GetCell(1).SetCellValue(executionDate[i].Month);
                    newWorkSheet.GetRow(j + 1).GetCell(2).SetCellValue(executionDate[i].Day);
                    newWorkSheet.GetRow(j + 1).GetCell(3).SetCellValue(lineName + "A");
                    newWorkSheet.GetRow(j + 1).GetCell(4).SetCellValue(ramdomCurrentA + " A");
                    newWorkSheet.GetRow(j + 1).GetCell(5).SetCellValue("端子  " + ramdomTemperatureA + " ~ " + (ramdomTemperatureA + addTemperature) + " ℃");

                    newWorkSheet.GetRow(j + 2).GetCell(0).SetCellValue(executionDate[i].Year);
                    newWorkSheet.GetRow(j + 2).GetCell(1).SetCellValue(executionDate[i].Month);
                    newWorkSheet.GetRow(j + 2).GetCell(2).SetCellValue(executionDate[i].Day);
                    newWorkSheet.GetRow(j + 2).GetCell(3).SetCellValue(lineName + "B");
                    newWorkSheet.GetRow(j + 2).GetCell(4).SetCellValue(ramdomCurrentB + " A");
                    newWorkSheet.GetRow(j + 2).GetCell(5).SetCellValue("端子  " + ramdomTemperatureB + " ~ " + (ramdomTemperatureB + addTemperature) + " ℃");
                    j = j + 2;
                }
            }
            else
            {
                for (int i = 1; i < executionDate.Count; i++)
                {
                    int ramdomCurrentA = 0, ramdomTemperatureA = 0, addTemperature = 3;
                    if (lineName == "水5" || lineName == "水6")
                    {
                        //亂數電流，介於180~450
                        ramdomCurrentA = RandomCurrent(newWorkSheet, 180, 450, 0, 0, false);
                    }
                    else
                    {
                        //亂數電流，介於540~1480
                        ramdomCurrentA = RandomCurrent(newWorkSheet, 540, 1480, 0, 0, false);
                    }
                    ramdomTemperatureA = GetTemperature(ramdomCurrentA);
                    if (ramdomTemperatureA == 37)
                        addTemperature = 2;
                    newWorkSheet.GetRow(i + 1).GetCell(0).SetCellValue(executionDate[i].Year);
                    newWorkSheet.GetRow(i + 1).GetCell(1).SetCellValue(executionDate[i].Month);
                    newWorkSheet.GetRow(i + 1).GetCell(2).SetCellValue(executionDate[i].Day);
                    newWorkSheet.GetRow(i + 1).GetCell(3).SetCellValue(lineName);
                    newWorkSheet.GetRow(i + 1).GetCell(4).SetCellValue(ramdomCurrentA + " A");
                    newWorkSheet.GetRow(i + 1).GetCell(5).SetCellValue("端子  " + ramdomTemperatureA + " ~ " + (ramdomTemperatureA + addTemperature) + " ℃");
                }
            }

            SetPrintStyle(workSheet);
            workSheet.GetRow(0).CreateCell(25).SetAsActiveCell();
            workBook.SetActiveSheet(0);
            file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workBook.Write(file, true);
            Console.WriteLine(lineName + " 寫入 " + GetFileName(file.Name) + " 成功");
            workBook.Close();
            file.Close();
        }
        static int GetTemperature(int ramdomCurrent)
        {
            int temperature = 0;
            if (ramdomCurrent <= 180)
                temperature = 30;
            else if (ramdomCurrent <= 365 && ramdomCurrent > 180)
                temperature = 31;
            else if (ramdomCurrent <= 550 && ramdomCurrent > 365)
                temperature = 32;
            else if (ramdomCurrent <= 735 && ramdomCurrent > 550)
                temperature = 33;
            else if (ramdomCurrent <= 920 && ramdomCurrent > 735)
                temperature = 34;
            else if (ramdomCurrent <= 1105 && ramdomCurrent > 920)
                temperature = 35;
            else if (ramdomCurrent <= 1290 && ramdomCurrent > 1105)
                temperature = 36;
            else if (ramdomCurrent <= 1480 && ramdomCurrent > 1290)
                temperature = 37;
            return temperature;
        }
        static void DoForm_A02ToA06(string dirPath, string newDirPath, List<DateTime> executionDate, string machineCode, string lineName)
        {
            string machineName = "";
            Dictionary<string, int> cloneSheetIndexs = new Dictionary<string, int> { };
            Dictionary<string, int> blockToWriteDatas = new Dictionary<string, int> { };
            //FOR VCP
            if (lineName == "水5" || lineName == "水6")
            {
                machineName = "水平電鍍線(VCP)(" + lineName.Remove(0, 1) + ")線";
                cloneSheetIndexs.Add(lineName, 0);
                blockToWriteDatas.Add(lineName, 1);
            }
            //FOR SVCP
            else if (lineName == "水7" || lineName == "水8" || lineName == "水9" || lineName == "水10" || lineName == "水11" || lineName == "水12")
            {
                machineName = "水平電鍍線(SVCP)(" + lineName.Remove(0, 1) + ")線";
                cloneSheetIndexs.Add(lineName, 1);
                blockToWriteDatas.Add(lineName, 2);
            }
            //FOR PTH
            else if (lineName == "PTH#4" || lineName == "PTH#5" || lineName == "PTH#6")
            {
                machineName = "水平PTH(" + lineName.Remove(0, 4) + ")線";
                cloneSheetIndexs.Add(lineName + "_1", 2);
                cloneSheetIndexs.Add(lineName + "_2", 3);
                blockToWriteDatas.Add(lineName + "_1", 6);
                blockToWriteDatas.Add(lineName + "_2", 2);
            }
            //FOR DEBURR#1
            else if (lineName == "DEBURR#1")
            {
                machineName = "DEBURR(1)線";
                cloneSheetIndexs.Add(lineName, 4);
                blockToWriteDatas.Add(lineName, 2);
            }
            //FOR 雷射孔微蝕#2
            else if (lineName == "雷射孔微蝕#2")
            {
                machineName = "雷射孔微蝕(2)線";
                cloneSheetIndexs.Add(lineName, 5);
                blockToWriteDatas.Add(lineName, 1);
            }

            FileStream file;
            IWorkbook workBook = null;
            ISheet newWorkSheet = null;
            if (File.Exists(newDirPath + "A02~A06-設備性能檢測數值記錄表.xls"))
            {
                foreach(var sheetIndex in cloneSheetIndexs)
                {
                    file = new FileStream(newDirPath + "A02~A06-設備性能檢測數值記錄表.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    workBook = new HSSFWorkbook(file);
                    file.Close();
                    newWorkSheet = workBook.CloneSheet(sheetIndex.Value);
                    workBook.SetSheetName(workBook.NumberOfSheets - 1, sheetIndex.Key);

                    newWorkSheet.GetRow(1).GetCell(0).SetCellValue("設備名稱:  " + machineName);
                    newWorkSheet.GetRow(1).GetCell(8).SetCellValue("檢測日期:" + executionDate[0].ToString("  yyyy   /    M    /   dd"));

                    //填入亂數產生的調整前電流數值                    
                    string standerCurrent = "";
                    int row = 5;
                    int column = 1;
                    for (int j = 0; j < blockToWriteDatas[sheetIndex.Key]; j++)
                    {
                        int ii = j;
                        if (j >= 3)
                        {
                            row = 15;
                            ii = j - 3;
                        }
                        standerCurrent = newWorkSheet.GetRow(row).GetCell(column + (ii * 3)).StringCellValue.TrimEnd('A');
                        int standerCurrentInt = StringToInt(standerCurrent);
                        //亂數決定增加多少電流，0.5~3.9A
                        double randomCurrentForAdd = (double)RandomCurrent(newWorkSheet, 5, 39, 0, 0, false) / 10;
                        double toWriteRandomCurrent = (double)standerCurrentInt + randomCurrentForAdd;
                        for (int k = 0; k < 3; k++)
                        {
                            //亂數決定增加多少量測誤差電流，-0.1~0.2A
                            double tolerance = (double)RandomCurrent(newWorkSheet, -1, 2, 0, 0, false) / 10;
                            //寫入表格，保留至小數第一位
                            newWorkSheet.GetRow(row + 2).GetCell(column + (ii * 3) + k).SetCellValue((toWriteRandomCurrent + tolerance).ToString("0.0") + " A");
                        }
                    }
                    SetPrintStyle(newWorkSheet);

                    file = new FileStream(newDirPath + "A02~A06-設備性能檢測數值記錄表.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    workBook.Write(file, true);
                    file.Close();
                    workBook.Close();
                    Console.WriteLine(lineName + " 寫入 " + GetFileName(file.Name) + " 成功");
                }
            }            
        }

        static bool DoCurrentCheckForm(string storageGridWords, DateTime date)
        {
            bool doCurrentCheckForm = false;
            string[] storageGridWord = storageGridWords.Split(',');
            foreach (string nextMonth in storageGridWord)
            {
                if (nextMonth == date.AddMonths(1).Month.ToString())
                    doCurrentCheckForm = true;
            }
            return doCurrentCheckForm;
        }
        static void DoForm_A07A08(string dirPath, string newDirPath, List<DateTime> executionDate, string machineCode, string lineName)
        {
            FileInfo directoryAFile = new FileInfo(dirPath + "A07~A08電流比對紀錄表.xls");
            FileStream file = null;
            IWorkbook workBook = null;
            //複製檔案
            if (!File.Exists(newDirPath + directoryAFile.Name))
            {
                file = new FileStream(directoryAFile.FullName, FileMode.Open, FileAccess.Read);
                workBook = new HSSFWorkbook(file);
                file.Close();
                file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                file.Close();
            }

            file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workBook = new HSSFWorkbook(file);
            file.Close();
            ISheet newWorkSheet;
            if (machineCode == "G07" || machineCode == "G08")
            {
                newWorkSheet = workBook.CloneSheet(0);
                workBook.SetSheetName(workBook.NumberOfSheets - 1, lineName);
                workBook.SetActiveSheet(workBook.NumberOfSheets - 1);
            }
            else if (machineCode == "G09" || machineCode == "G10" || machineCode == "G11" || machineCode == "G12" || machineCode == "G24" || machineCode == "G25")
            {
                newWorkSheet = workBook.CloneSheet(1);
                workBook.SetSheetName(workBook.NumberOfSheets - 1, lineName);
            }
            else
            {
                newWorkSheet = workBook.GetSheetAt(0);
                workBook.SetSheetName(0, lineName);
            }
            file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workBook.Write(file, true);
            file.Close();
            int gridRow = 0, gridColumn = 0, checkBoxIndex = 0, checkBoxIndex2 = 0;
            //FOR 水5、水6
            if (machineCode == "G07" || machineCode == "G08")
            {
                gridRow = 1;
                gridColumn = 15;
                if (lineName == "水5")
                    checkBoxIndex = 9;
                else if (lineName == "水6")
                    checkBoxIndex = 21;
            }
            //FOR 水7~水12
            else if (machineCode == "G09" || machineCode == "G10" || machineCode == "G11" || machineCode == "G12" || machineCode == "G24" || machineCode == "G25")
            {
                gridRow = 2;
                gridColumn = 9;
                switch (lineName)
                {
                    case "水7":
                        checkBoxIndex = 3;
                        break;
                    case "水9":
                        checkBoxIndex = 19;
                        break;
                    case "水11":
                        checkBoxIndex = 35;
                        break;
                    case "水8A":
                        checkBoxIndex = 7;
                        checkBoxIndex2 = 11;
                        break;
                    case "水8B":
                        checkBoxIndex = 7;
                        checkBoxIndex2 = 14;
                        break;
                    case "水10A":
                        checkBoxIndex = 23;
                        checkBoxIndex2 = 27;
                        break;
                    case "水10B":
                        checkBoxIndex = 23;
                        checkBoxIndex2 = 30;
                        break;
                    case "水12A":
                        checkBoxIndex = 40;
                        checkBoxIndex2 = 45;
                        break;
                    case "水12B":
                        checkBoxIndex = 40;
                        checkBoxIndex2 = 48;
                        break;
                }
            }
            WriteRamdomDataIntoA07A08Form(directoryAFile, newDirPath, gridRow, gridColumn, executionDate[0], lineName, checkBoxIndex, checkBoxIndex2);
        }
        static void WriteRamdomDataIntoA07A08Form(FileInfo directoryAFile, string newDirPath, int gridRow, int gridColumn, DateTime executionDate, string lineName, int checkBoxIndex, int checkBoxIndex2)
        {
            FileStream file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IWorkbook workBook = new HSSFWorkbook(file);
            ISheet workSheet = workBook.GetSheet(lineName);

            //填入資料
            if (lineName == "水5" || lineName == "水6")
            {
                //亂數產生設定電流值後填表，介於180~450
                RandomCurrent(workSheet, 180, 450, 18, 5, true);
                
            }
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11"
                     || lineName == "水8A" || lineName == "水10A" || lineName == "水12A"
                     || lineName == "水8B" || lineName == "水10B" || lineName == "水12B")
            {
                //亂數產生設定電流值後填表，介於540~1480
                RandomCurrent(workSheet, 540, 1480, 10, 6, true);
            }

            //填入執行日期
            string storageGridDate = workSheet.GetRow(gridRow).GetCell(gridColumn).StringCellValue;
            if (storageGridDate != "" && storageGridDate != null)
            {
                storageGridDate = storageGridDate.Remove(4, executionDate.Year.ToString().Length);
                storageGridDate = storageGridDate.Insert(4, executionDate.Year.ToString());
                storageGridDate = storageGridDate.Remove(12, executionDate.Month.ToString().Length);
                storageGridDate = storageGridDate.Insert(12, executionDate.Month.ToString());
                storageGridDate = storageGridDate.Remove(17, executionDate.Day.ToString().Length);
                storageGridDate = storageGridDate.Insert(17, executionDate.Day.ToString());
                workSheet.GetRow(gridRow).GetCell(gridColumn).SetCellValue(storageGridDate);
            }

            //勾選線別
            string storageGridLineName = workSheet.GetRow(gridRow).GetCell(0).StringCellValue;
            storageGridLineName = storageGridLineName.Remove(checkBoxIndex, 1);
            storageGridLineName = storageGridLineName.Insert(checkBoxIndex, "R");
            if (lineName == "水8A" || lineName == "水10A" || lineName == "水12A"
                || lineName == "水8B" || lineName == "水10B" || lineName == "水12B")
            {
                storageGridLineName = storageGridLineName.Remove(checkBoxIndex2, 1);
                storageGridLineName = storageGridLineName.Insert(checkBoxIndex2, "R");
            }
            HSSFRichTextString lineNameToGrid = new HSSFRichTextString(storageGridLineName);
            IFont font = workBook.CreateFont();
            //字型
            font.FontName = "Wingdings 2";
            //字體尺寸
            font.FontHeightInPoints = 14;
            //FOR PTH#5、PTH#6
            if (lineName == "水5" || lineName == "水6")
            {
                lineNameToGrid.ApplyFont(9, 10, font);
                lineNameToGrid.ApplyFont(21, 22, font);
            }
            //FOR 水7~水12
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11"
                     || lineName == "水8A" || lineName == "水10A" || lineName == "水12A"
                     || lineName == "水8B" || lineName == "水10B" || lineName == "水12B")
            {
                lineNameToGrid.ApplyFont(3, 4, font);
                lineNameToGrid.ApplyFont(19, 20, font);
                lineNameToGrid.ApplyFont(35, 36, font);
                lineNameToGrid.ApplyFont(7, 8, font);
                lineNameToGrid.ApplyFont(11, 12, font);
                lineNameToGrid.ApplyFont(14, 15, font);
                lineNameToGrid.ApplyFont(23, 24, font);
                lineNameToGrid.ApplyFont(27, 28, font);
                lineNameToGrid.ApplyFont(30, 31, font);
                lineNameToGrid.ApplyFont(40, 41, font);
                lineNameToGrid.ApplyFont(45, 46, font);
                lineNameToGrid.ApplyFont(48, 49, font);
            }
            workSheet.GetRow(gridRow).GetCell(0).SetCellValue(lineNameToGrid);
            SetPrintStyle(workSheet);
            file = new FileStream(newDirPath + directoryAFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workBook.Write(file, true);
            Console.WriteLine(lineName + " 寫入 " + GetFileName(file.Name) + " 成功");
            file.Close();
        }
        static int RandomCurrent(ISheet workSheet, int minCurrent, int maxCurrent, int forCount, int startRow, bool ifA07A08)
        {
            Random randomNumber = new Random(Guid.NewGuid().GetHashCode());
            //亂數產生設定電流值後填表，介於minCurrent~maxCurrent
            int randomSetCurrent = randomNumber.Next(minCurrent, maxCurrent);
            if (ifA07A08)
            {
                for (int i = 3; i <= forCount; i++)
                {
                    workSheet.GetRow(startRow).GetCell(i).SetCellValue(randomSetCurrent + "A");
                    //亂數產生實際電流值，介於(設定電流值的96%)~(設定電流值+1)
                    int randomActualCurrent = randomNumber.Next(Convert.ToInt32(randomSetCurrent * 0.96), randomSetCurrent + 1);
                    workSheet.GetRow(startRow + 1).GetCell(i).SetCellValue(randomActualCurrent + "A");
                    double errorPercentTemp = Math.Abs(randomSetCurrent - randomActualCurrent);
                    double errorPercentTemp2 = errorPercentTemp / randomSetCurrent * 100;
                    double errorPercent = Math.Round(errorPercentTemp2, 1, MidpointRounding.AwayFromZero);
                    workSheet.GetRow(startRow + 2).GetCell(i).SetCellValue(errorPercent + "%");
                }
            }
            return randomSetCurrent;
        }
        static void DoAppointmentMaintenanceFormExcelFile(string dirPath, string newDirPath, DateTime date, string month)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo directoryGFile = new FileInfo(newDirPath + "後三月預保養表.xls"); ;
                FileInfo[] directoryAFiles = new FileInfo[] { };
                if (Directory.Exists(dirPath))
                {
                    // 取得資料夾內所有檔案
                    DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                    //所有A開頭的EXCLE檔
                    directoryAFiles = directoryInfo.GetFiles("A*.xls");
                }

                //獲取月份第一日及天數
                DateTime monthFirstDay;
                int daysOfMonth;
                MonthFirstDayAndDays(date, out monthFirstDay, out daysOfMonth);
                //獲取下個月份第一日
                DateTime nextMonthFirstDay;
                NextMonthFirstDay(date, out nextMonthFirstDay);

                FileStream file = new FileStream(directoryGFile.FullName, FileMode.Open, FileAccess.Read);
                IWorkbook workBook = new HSSFWorkbook(file);
                file.Close();
                HSSFSimpleShape circle1;
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

                int monthInteger = StringToInt(month);
                int[] monthAdd = new int[3] { monthInteger + 1, monthInteger + 2, monthInteger + 3 };
                for (int k = 0; k < 3; k++)
                {
                    if (monthAdd[k] > 12)
                        monthAdd[k] = monthAdd[k] - 12;
                }
                string monthAddOne = (monthAdd[0]).ToString();
                string monthAddTwo = (monthAdd[1]).ToString();
                string monthAddThree = (monthAdd[2]).ToString();

                for (int j = 0; j < workBook.NumberOfSheets; j++)
                {
                    ISheet workSheet = workBook.GetSheetAt(j);

                    //填入保養月份
                    workSheet.GetRow(1).GetCell(3).SetCellValue(monthAddOne + "、" + monthAddTwo + "、" + monthAddThree);
                    workSheet.GetRow(1).GetCell(3).CellStyle = cellStyle;

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
                                DrowingCircle(false, workBook, workSheet, i, x1, x2, 0, out circle1);
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
                            SetCellStyle(workBook, workSheet, i, 6, 14, 1);
                    }

                    //抓取表單的名子
                    string[] formName = workSheet.SheetName.Split('-');
                    //圈取設備代碼
                    CircleMachineCode(formName[0], workBook, workSheet);

                    SetPrintStyle(workSheet);
                    workSheet.GetRow(0).CreateCell(25).SetAsActiveCell();
                    workBook.SetActiveSheet(0);
                    file = new FileStream(newDirPath + "後三月預保養表.xls", FileMode.Create, FileAccess.Write);
                    workBook.Write(file, true);
                    file.Close();
                }
                Console.WriteLine("寫入 " + GetFileName(file.Name) + " 成功");
                workBook.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel檔案開啟出錯：" + ex.Message);
            }
        }
        static void DoLayoutFormExcelFile(string dirPath, string newDirPath, string month)
        {
            //開啟Excel 2003檔案
            FileInfo directoryGFile = null;
            if (File.Exists(dirPath + "製七部LAYOUT圖(緊急開關、液位開關)A0018656更新.xls"))
            {
                directoryGFile = new FileInfo(dirPath + "製七部LAYOUT圖(緊急開關、液位開關)A0018656更新.xls");
            }
            //複製檔案
            //讀取原始檔
            FileStream file = new FileStream(directoryGFile.FullName, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = new HSSFWorkbook(file);
            file.Close();
            int[] totalDatas = new int[6] { 0, 0, 0, 0, 0, 0 };
            int[] datas = new int[6] { 0, 0, 0, 0, 0, 0 };
            for (int i = 0; i < 18; i++)
            {
                ISheet workSheet = workBook.GetSheetAt(i);
                int row = 0;
                int column = 0;
                //datas = new int[6] { 0, 0, 0, 0, 0, 0 };
                switch (workSheet.SheetName)
                {
                    case "Desmear#3":
                        row = 4;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 4;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 11;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 2;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 25;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 4;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 6;
                                datas[4] = 0;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "Desmear#4":
                        row = 4;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 5;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 21;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 15;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 6;
                                datas[4] = 0;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "Desmear#5":
                        row = 4;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 4;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 21;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 15;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 2;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 6;
                                datas[4] = 0;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "Deburr#1":
                        row = 4;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 5;
                                datas[1] = 0;
                                datas[2] = 2;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 0;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 2;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "PTH#5":
                        row = 3;
                        column = 13;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 4;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 7;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 8;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 3;
                                datas[1] = 4;
                                datas[2] = 0;
                                datas[3] = 13;
                                datas[4] = 4;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "PTH#6":
                        row = 5;
                        column = 13;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 4;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 7;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 4;
                                datas[2] = 0;
                                datas[3] = 9;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 3;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 10;
                                datas[4] = 0;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "水平電鍍五線":
                        row = 3;
                        column = 17;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 0;
                                datas[2] = 3;
                                datas[3] = 2;
                                datas[4] = 2;
                                datas[5] = 1;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 0;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 4;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                        }
                        break;
                    case "水平電鍍六線":
                        row = 3;
                        column = 14;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 0;
                                datas[2] = 3;
                                datas[3] = 2;
                                datas[4] = 2;
                                datas[5] = 1;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 0;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 3;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                        }
                        break;
                    case "水平電鍍七線":
                        row = 5;
                        column = 14;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 3;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 1;
                                datas[4] = 2;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "水平電鍍八線":
                        row = 2;
                        column = 17;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 6;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 7;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 2;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 3;
                                datas[3] = 2;
                                datas[4] = 0;
                                datas[5] = 2;
                                break;
                        }
                        break;
                    case "水平電鍍九線":
                        row = 2;
                        column = 20;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 3;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 1;
                                datas[4] = 2;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "水平電鍍十線":
                        row = 2;
                        column = 17;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 6;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 7;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 2;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 3;
                                datas[3] = 2;
                                datas[4] = 0;
                                datas[5] = 2;
                                break;
                        }
                        break;
                    case "雷燒孔微蝕#2":
                        row = 2;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 0;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 5;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 0;
                                datas[3] = 5;
                                datas[4] = 1;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "2F 讀孔機":
                        row = 9;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "2":
                            case "3":
                            case "4":
                            case "5":
                            case "6":
                            case "7":
                            case "8":
                            case "9":
                            case "10":
                            case "11":
                            case "12":
                                datas[0] = 4;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                        }
                        break;
                    case "PTH#4":
                        row = 2;
                        column = 11;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 4;
                                datas[2] = 0;
                                datas[3] = 9;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 2;
                                datas[1] = 3;
                                datas[2] = 0;
                                datas[3] = 6;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 2;
                                datas[1] = 2;
                                datas[2] = 0;
                                datas[3] = 16;
                                datas[4] = 4;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "水平電鍍十一線":
                        row = 2;
                        column = 17;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 3;
                                datas[1] = 1;
                                datas[2] = 3;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 1;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 1;
                                datas[4] = 2;
                                datas[5] = 1;
                                break;
                        }
                        break;
                    case "水平電鍍十二線":
                        row = 2;
                        column = 18;
                        switch (month)
                        {
                            case "1":
                            case "4":
                            case "7":
                            case "10":
                                datas[0] = 6;
                                datas[1] = 1;
                                datas[2] = 3;
                                datas[3] = 3;
                                datas[4] = 0;
                                datas[5] = 0;
                                break;
                            case "2":
                            case "5":
                            case "8":
                            case "11":
                                datas[0] = 7;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 0;
                                datas[4] = 1;
                                datas[5] = 0;
                                break;
                            case "3":
                            case "6":
                            case "9":
                            case "12":
                                datas[0] = 6;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 2;
                                datas[4] = 1;
                                datas[5] = 2;
                                break;
                        }
                        break;
                    case "2F 烤箱":
                        row = 4;
                        column = 10;
                        switch (month)
                        {
                            case "1":
                            case "2":
                            case "3":
                            case "4":
                            case "5":
                            case "6":
                            case "7":
                            case "8":
                            case "9":
                            case "10":
                            case "11":
                            case "12":
                                datas[0] = 8;
                                datas[1] = 0;
                                datas[2] = 0;
                                datas[3] = 2;
                                datas[4] = 0;
                                datas[5] = 2;
                                break;
                        }
                        break;
                }
                
                for (int j = 0; j < datas.Length; j++)
                {
                    workSheet.GetRow(row + j).GetCell(column).SetCellValue(datas[j]);
                    SetCellStyle(workBook, workSheet, row + j, column, 12, 2);
                    totalDatas[j] += datas[j];
                    workSheet.GetRow(row + j).GetCell(column + 1).SetCellValue(0);
                    SetCellStyle(workBook, workSheet, row + j, column + 1, 12, 2);
                }
                file = new FileStream(newDirPath + directoryGFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                Console.WriteLine(workSheet.SheetName + " 寫入 " + GetFileName(file.Name) + " 成功");
                file.Close();
            }
            workBook.Close();
            if (File.Exists(newDirPath + "數量統計.xls"))
            {
                file = new FileStream(newDirPath + "數量統計.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook = new HSSFWorkbook(file);
                file.Close();
                ISheet workSheet = workBook.GetSheet("數量統計");

                for(int k = 0; k < datas.Length; k++)
                {
                    string cellString = "";
                    switch (k)
                    {
                        case 0:
                            cellString = "緊急開關";
                            break;
                        case 1:
                            cellString = "液位浮球";
                            break;
                        case 2:
                            cellString = "液位極棒";
                            break;
                        case 3:
                            cellString = "加熱器";
                            break;
                        case 4:
                            cellString = "機械浮球";
                            break;
                        case 5:
                            cellString = "烘乾段風車";
                            break;
                    }
                    workSheet.CreateRow(2 + k).CreateCell(0).SetCellValue(cellString);
                    SetCellStyle(workBook, workSheet, 2 + k, 0, 10, 2);
                    workSheet.GetRow(2 + k).CreateCell(1).SetCellValue(totalDatas[k]);
                    SetCellStyle(workBook, workSheet, 2 + k, 1, 10, 2);
                }
                file = new FileStream(newDirPath + "數量統計.xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook.Write(file, true);
                file.Close();
                workBook.Close();
            }
        }
        static void DrowingCircle(bool Maintenance, IWorkbook workBook, ISheet workSheet, int i, int x1, int x2, int machineCodeNumber)
        {
            bool notUse = false;
            DrowingCircle(Maintenance, workBook, workSheet, i, x1, x2, machineCodeNumber, out HSSFSimpleShape circle1, ref notUse);
        }
        static HSSFPatriarch DrowingCircle(bool Maintenance, IWorkbook workBook, ISheet workSheet, int i, int x1, int x2, int machineCodeNumber, out HSSFSimpleShape circle1)
        {
            bool notUse = false;
            return DrowingCircle(Maintenance, workBook, workSheet, i, x1, x2, machineCodeNumber, out circle1, ref notUse);
        }
        static HSSFPatriarch DrowingCircle(bool Maintenance, IWorkbook workBook, ISheet workSheet, int i, int x1, int x2, int machineCodeNumber, out HSSFSimpleShape circle1, ref bool heaterCheck)
        {
            //heaterCheck = false;
            int initial = 6;
            if (machineCodeNumber == 18 || machineCodeNumber == 19)
            {
                initial = 8;
            }
            else if (machineCodeNumber == 20 || machineCodeNumber == 21
                     || machineCodeNumber == 26 || machineCodeNumber == 28 || machineCodeNumber == 29)
            {
                if (i == 28)
                    initial = 9;
                else
                    initial = 1;
            }
            //儲存格畫圈
            HSSFPatriarch patriarchCircle = (HSSFPatriarch)workSheet.CreateDrawingPatriarch();
            HSSFClientAnchor c1 = new HSSFClientAnchor(x1, 30, x2, 226, initial, i, initial, i);
            circle1 = patriarchCircle.CreateSimpleShape(c1);
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
                //附加熱器檢查表
                if (workSheet.GetRow(i).GetCell(2).ToString() == "附檢測資料")
                {
                    heaterCheck = true;
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
