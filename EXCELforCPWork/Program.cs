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
            for (int monthToAdd = -1; monthToAdd < 1; monthToAdd++)
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

                //製作保養表及產生相關附件
                DoMaintenanceFormExcelFile(dirPath, dirPathMaintenanceForm, date, month, dirPathAttachment);
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

        static void DoMaintenanceFormExcelFile(string dirPath, string folderPath, DateTime date, string month, string dirPathAttachment)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo[] directoryGFiles = new FileInfo[] { };
                FileInfo[] directoryAFiles = new FileInfo[] { };
                if (Directory.Exists(dirPath))
                {
                    // 取得資料夾內所有檔案
                    DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                    //所有G開頭的EXCLE檔
                    directoryGFiles = directoryInfo.GetFiles("G*.xls");
                    directoryAFiles = directoryInfo.GetFiles("A*.xls");
                }
                int monthInteger = StringToInt(month);

                //獲取月份第一日及天數
                DateTime monthFirstDay;
                int daysOfMonth;
                MonthFirstDayAndDays(date, out monthFirstDay, out daysOfMonth);
                //獲取下個月份第一日
                DateTime nextMonthFirstDay;
                NextMonthFirstDay(date, out nextMonthFirstDay);

                foreach (FileInfo directoryFile in directoryGFiles)
                {
                    if (File.Exists(dirPath + directoryFile.Name))
                    {
                        FileStream file;
                        IWorkbook workBook;
                        ISheet workSheet;
                        file = new FileStream(dirPath + directoryFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);                        
                        workBook = new HSSFWorkbook(file);
                        workSheet = workBook.GetSheetAt(0);
                        
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

                        bool heaterCheck = false;
                        HSSFSimpleShape circle1;
                        for (int i = 3; i < workSheet.LastRowNum - 1; i++)
                        {
                            int x1 = 0;
                            int x2 = 0;
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
                                DrowingCircle(true, workBook, workSheet, i, x1, x2, 0, out circle1, ref heaterCheck);
                            }
                            else if (workSheet.GetRow(i).GetCell(6).ToString() != ""
                                    && workSheet.GetRow(i).GetCell(6).ToString() != "1~12")
                            {
                                DrowingLine(workSheet, i);
                            }
                            if(workSheet.GetRow(i).GetCell(6).ToString() != "")
                                SetCellStyle(workBook, workSheet, i);
                        }
                        //抓取表單的名子
                        string[] formName = directoryFile.Name.Split('-');
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
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G04", "DEBURR#1");
                                break;
                            //PTH#4，第2個星期三保
                            case "G22":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Wednesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G22", "PTH#4");
                                break;
                            //PTH#5，第2個星期一保
                            case "G05":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G05", "PTH#5");
                                break;
                            //水5，第2個星期一保
                            case "G07":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G07", "水5");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G07", "水5");
                                if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G07", "水5");
                                }
                                break;
                            //PTH#6，第2個星期二保
                            case "G06":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G06", "PTH#6");
                                break;
                            //水6，第2個星期二保
                            case "G08":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G08", "水6");
                                if(heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G08", "水6");
                                if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G08", "水6");
                                }
                                break;
                            //水7，第1個星期四保
                            case "G09":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Thursday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G09", "水7");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G09", "水7");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G09", "水7");
                                }
                                break;
                            //水8，第1個星期一保
                            case "G10":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Monday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G10", "水8");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G10", "水8");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G10", "水8");
                                }
                                break;
                            //水9，第2個星期四保
                            case "G11":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Thursday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G11", "水9");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G11", "水9");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G11", "水9");
                                }
                                break;
                            //水10，第1個星期三保
                            case "G12":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Wednesday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G12", "水10");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G12", "水10");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Wednesday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G12", "水10");
                                }
                                break;
                            //雷射孔微蝕#2，第3個星期二保
                            case "G13":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Tuesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G13", "雷射孔微蝕#2");
                                break;
                            //文坦讀孔機，第2個星期日保
                            case "G18":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Sunday");
                                break;
                            //水11，第1個星期二保
                            case "G24":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Tuesday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G24", "水11");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G24", "水11");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G24", "水11");
                                }
                                break;
                            //水12，第1個星期五保
                            case "G25":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Friday");
                                DoForm_A01(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G25", "水12");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathAttachment, directoryAFiles, executionDate, "G25", "水12");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Friday");
                                    DoForm_A07A08(dirPath, dirPathAttachment, directoryAFiles, nextMonthExecutionDate, "G25", "水12");
                                }
                                break;
                            //PLASMA，第2個星期五保
                            case "G26":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Friday");
                                break;
                        }

                        //填入執行日期
                        workSheet.GetRow(1).GetCell(8).SetCellValue(executionDate[0].ToString("yyyy   /    M    /    d"));
                        workSheet.GetRow(1).GetCell(8).CellStyle.SetFont(font2);
                        //文坦讀孔機
                        if (formName[0] == "G18")
                        {
                            MachineCodeDrowingCircle("G19", 693, 993, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G20", 8, 238, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G21", 310, 540, workBook, folderPath, directoryFile);
                            //For G18
                            DrowingCircle(false, workBook, workSheet, 28, 304, 604, 18);
                        }
                        //PLASMA
                        else if (formName[0] == "G26")
                        {
                            MachineCodeDrowingCircle("G28", 398, 625, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G29", 702, 932, workBook, folderPath, directoryFile);
                            //For G26
                            DrowingCircle(false, workBook, workSheet, 28, 99, 328, 26);
                            DrowingCircle(false, workBook, workSheet, 1, 255, 312, 26);
                        }

                        SetPrintStyle(workSheet);

                        file = new FileStream(folderPath + directoryFile.Name, FileMode.Create, FileAccess.Write);
                        workBook.Write(file, true);
                        Console.WriteLine(GetFileName(file.Name) + "寫入成功");
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
            HSSFSimpleShape c1;
            HSSFSimpleShape c2 = null;
            HSSFPatriarch circle = DrowingCircle(false, workBook, workSheet, 28, x3, x4, machineCodeNumber, out c1);
            HSSFPatriarch circle2 = null;
            if (machineCodeNumber == 28)
                circle2 = DrowingCircle(false, workBook, workSheet, 1, 304, 360, machineCodeNumber, out c2);
            else if(machineCodeNumber == 29)
                circle2 = DrowingCircle(false, workBook, workSheet, 1, 350, 406, machineCodeNumber, out c2);
            SetPrintStyle(workSheet);
            FileStream newFile = new FileStream(folderPath + machineCode + "-" + directoryFile.Name, FileMode.Create, FileAccess.Write);
            workBook.Write(newFile, true);
            circle.RemoveShape(c1);
            if (machineCodeNumber == 28 || machineCodeNumber == 29) 
                circle2.RemoveShape(c2);
        }
        static void DoForm_A01(string dirPath, string dirPathAttachment, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {
            FileStream file;
            IWorkbook workBook;
            ISheet workSheet;
            file = new FileStream(dirPath + directoryAFiles[0].Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);            
            workBook = new HSSFWorkbook(file);
            workSheet = workBook.GetSheetAt(0);

            workSheet.GetRow(0).GetCell(0).SetCellValue(lineName);
            if (lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                int j = 1;
                for (int i = 1; i < executionDate.Count; i++)
                {
                    workSheet.GetRow(j + 1).GetCell(0).SetCellValue(executionDate[i].Year);
                    workSheet.GetRow(j + 1).GetCell(1).SetCellValue(executionDate[i].Month);
                    workSheet.GetRow(j + 1).GetCell(2).SetCellValue(executionDate[i].Day);
                    workSheet.GetRow(j + 1).GetCell(3).SetCellValue(lineName + "A");
                    workSheet.GetRow(j + 1).GetCell(4).SetCellValue("               A");
                    workSheet.GetRow(j + 1).GetCell(5).SetCellValue("端子      -     ℃");

                    workSheet.GetRow(j + 2).GetCell(0).SetCellValue(executionDate[i].Year);
                    workSheet.GetRow(j + 2).GetCell(1).SetCellValue(executionDate[i].Month);
                    workSheet.GetRow(j + 2).GetCell(2).SetCellValue(executionDate[i].Day);
                    workSheet.GetRow(j + 2).GetCell(3).SetCellValue(lineName + "B");
                    workSheet.GetRow(j + 2).GetCell(4).SetCellValue("               A");
                    workSheet.GetRow(j + 2).GetCell(5).SetCellValue("端子      -     ℃");
                    j = j + 2;
                }
            }
            else
            {
                for (int i = 1; i < executionDate.Count; i++)
                {
                    workSheet.GetRow(i + 1).GetCell(0).SetCellValue(executionDate[i].Year);
                    workSheet.GetRow(i + 1).GetCell(1).SetCellValue(executionDate[i].Month);
                    workSheet.GetRow(i + 1).GetCell(2).SetCellValue(executionDate[i].Day);
                    workSheet.GetRow(i + 1).GetCell(3).SetCellValue(lineName);
                    workSheet.GetRow(i + 1).GetCell(4).SetCellValue("               A");
                    workSheet.GetRow(i + 1).GetCell(5).SetCellValue("端子      -     ℃");
                }
            }

            SetPrintStyle(workSheet);

            file = new FileStream(dirPathAttachment + "A01-" + machineCode + "-" + lineName + "-亞碩競銘線纜線熱顯像檢查表.xls", FileMode.Create, FileAccess.Write);
            workBook.Write(file, true);
            Console.WriteLine(GetFileName(file.Name) + "寫入成功");
            workBook.Close();
            file.Close();
        }
        static void DoForm_A02ToA06(string dirPath, string dirPathAttachment, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {
            FileStream file;
            IWorkbook workBook = null;
            ISheet workSheet;
            string machineName = "";
            string openPath = "", writePath = "";
            string openPath2 = "", writePath2 = "";
            //FOR DEBURR#1
            if (lineName == "DEBURR#1")
            {
                machineName = "DEBURR(1)線";
                openPath = dirPath + directoryAFiles[5].Name;
                writePath = dirPathAttachment + "A05-" + machineCode + "-" + lineName + "-DEBURR設備性能檢測數值記錄表.xls";
            }
            //FOR 雷射孔微蝕#2
            else if (lineName == "雷射孔微蝕#2")
            {
                machineName = "雷射孔微蝕(2)線";
                openPath = dirPath + directoryAFiles[6].Name;
                writePath = dirPathAttachment + "A06-" + machineCode + "-" + lineName + "-雷射孔微蝕設備性能檢測數值記錄表.xls";
            }
            //FOR VCP
            else if (lineName == "水5" || lineName == "水6")
            {
                machineName = "水平電鍍線(VCP)(" + lineName.Remove(0,1) + ")線";
                openPath = dirPath + directoryAFiles[1].Name;
                writePath = dirPathAttachment + "A02-" + machineCode + "-" + lineName + "-水平電鍍線(VCP)設備性能檢測數值記錄表.xls";
            }
            //FOR PTH
            else if (lineName == "PTH#4" || lineName == "PTH#5" || lineName == "PTH#6")
            {
                machineName = "水平PTH(" + lineName.Remove(0, 4) + ")線";
                openPath = dirPath + directoryAFiles[3].Name;
                writePath = dirPathAttachment + "A041-" + machineCode + "-" + lineName + "-PTH設備性能檢測數值記錄表.xls";

                openPath2 = dirPath + directoryAFiles[4].Name;
                writePath2 = dirPathAttachment + "A042-" + machineCode + "-" + lineName + "-PTH設備性能檢測數值記錄表.xls";
            }
            //FOR SVCP
            else
            {
                machineName = "水平電鍍線(SVCP)(" + lineName.Remove(0, 1) + ")線";
                openPath = dirPath + directoryAFiles[2].Name;
                writePath = dirPathAttachment + "A03-" + machineCode + "-" + lineName + "-水平電鍍線(SVCP)設備性能檢測數值記錄表.xls";
            }

            file = new FileStream(openPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);            
            workBook = new HSSFWorkbook(file);
            workSheet = workBook.GetSheetAt(0);

            workSheet.GetRow(1).GetCell(0).SetCellValue("設備名稱:  " + machineName);
            workSheet.GetRow(1).GetCell(8).SetCellValue("檢測日期:" + executionDate[0].ToString("  yyyy   /    M    /   dd"));

            SetPrintStyle(workSheet);

            file = new FileStream(writePath, FileMode.Create, FileAccess.Write);
            workBook.Write(file, true);
            Console.WriteLine(GetFileName(file.Name) + "寫入成功");
            if (lineName == "PTH#4" || lineName == "PTH#5" || lineName == "PTH#6")
            {
                file = new FileStream(openPath2, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workBook = new HSSFWorkbook(file);
                workSheet = workBook.GetSheetAt(0);

                workSheet.GetRow(1).GetCell(0).SetCellValue("設備名稱:  " + machineName);
                workSheet.GetRow(1).GetCell(8).SetCellValue("檢測日期:" + executionDate[0].ToString("  yyyy   /    M    /   dd"));

                SetPrintStyle(workSheet);

                file = new FileStream(writePath2, FileMode.Create, FileAccess.Write);
                workBook.Write(file, true);
                Console.WriteLine(GetFileName(file.Name) + "寫入成功");
            }
            workBook.Close();
            file.Close();
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
        static void DoForm_A07A08(string dirPath, string dirPathAttachment, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {
            //設定開啟及儲存的路徑跟檔名
            string openPath = "", writePath = "", writePath2 = "";
            //FOR PTH#5、PTH#6
            if (lineName == "水5" || lineName == "水6")
            {
                openPath = dirPath + directoryAFiles[7].Name;
                writePath = dirPathAttachment + "A07-" + machineCode + "-" + lineName + "-PTH電流比對紀錄表.xls";
            }
            //FOR 奇數水平電鍍線(水7、水9、水11)
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11")
            {
                openPath = dirPath + directoryAFiles[8].Name;
                writePath = dirPathAttachment + "A08-" + machineCode + "-" + lineName + "-水平電鍍線電流比對紀錄表.xls";
            }
            //FOR 偶數水平電鍍線(水8、水10、水12)
            else if (lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                openPath = dirPath + directoryAFiles[8].Name;
                writePath = dirPathAttachment + "A08-" + machineCode + "-" + lineName + "A-水平電鍍線電流比對紀錄表.xls";
                writePath2 = dirPathAttachment + "A08-" + machineCode + "-" + lineName + "B-水平電鍍線電流比對紀錄表.xls";
            }

            int gridRow = 0, gridColumn = 0, checkBoxIndex = 0, checkBoxIndexA = 0, checkBoxIndexB = 0;
            //FOR PTH#5、PTH#6
            if (lineName == "水5" || lineName == "水6")
            {
                gridRow = 1;
                gridColumn = 15;
                if(lineName == "水5")
                    checkBoxIndex = 9;
                else if(lineName == "水6")
                    checkBoxIndex = 21;
            }
            //FOR 水7~水12
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11"
                     || lineName == "水8" || lineName == "水10" || lineName == "水12")
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
                    case "水8":
                        checkBoxIndex = 7;
                        checkBoxIndexA = 11;
                        checkBoxIndexB = 14;
                        break;
                    case "水10":
                        checkBoxIndex = 23;
                        checkBoxIndexA = 27;
                        checkBoxIndexB = 30;
                        break;
                    case "水12":
                        checkBoxIndex = 40;
                        checkBoxIndexA = 45;
                        checkBoxIndexB = 48;
                        break;
                }
            }
            DoA07A08Form(openPath, writePath, gridRow, gridColumn, executionDate[0], lineName, checkBoxIndex, checkBoxIndexA);
            if (lineName == "水8" || lineName == "水10" || lineName == "水12")
                DoA07A08Form(openPath, writePath2, gridRow, gridColumn, executionDate[0], lineName, checkBoxIndex, checkBoxIndexB);
        }
        static void DoA07A08Form(string openPath, string writePath, int gridRow, int gridColumn, DateTime executionDate, string lineName, int checkBoxIndex, int checkBoxIndexAB)
        {
            FileStream file = new FileStream(openPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IWorkbook workBook = new HSSFWorkbook(file);
            ISheet workSheet = workBook.GetSheetAt(0);

            Random randomNumber = new Random(Guid.NewGuid().GetHashCode());

            //填入資料
            if (lineName == "水5" || lineName == "水6")
            {
                //亂數產生設定電流值，介於300~500
                int randomSetCurrent = randomNumber.Next(300, 500);
                for (int i = 3; i <= 18; i++)
                {
                    workSheet.GetRow(5).GetCell(i).SetCellValue(randomSetCurrent + "A");
                    //亂數產生實際電流值，介於(設定電流值的96%)~(設定電流值+2)
                    int randomActualCurrent = randomNumber.Next(Convert.ToInt32(randomSetCurrent * 0.96), randomSetCurrent + 2);
                    workSheet.GetRow(6).GetCell(i).SetCellValue(randomActualCurrent + "A");
                    double errorPercentTemp = Math.Abs(randomSetCurrent - randomActualCurrent);
                    double errorPercentTemp2 = errorPercentTemp / randomSetCurrent * 100;
                    double errorPercent = Math.Round(errorPercentTemp2, 1, MidpointRounding.AwayFromZero);
                    workSheet.GetRow(7).GetCell(i).SetCellValue(errorPercent + "%");
                }
            }
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11"
                     || lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                //亂數產生設定電流值，介於540~1250
                int randomSetCurrent = randomNumber.Next(540, 1250);
                for (int i = 3; i <= 10; i++)
                {
                    workSheet.GetRow(6).GetCell(i).SetCellValue(randomSetCurrent + "A");
                    //亂數產生實際電流值，介於(設定電流值的96%)~(設定電流值+2)
                    int randomActualCurrent = randomNumber.Next(Convert.ToInt32(randomSetCurrent * 0.96), randomSetCurrent + 2);
                    workSheet.GetRow(7).GetCell(i).SetCellValue(randomActualCurrent + "A");
                    double errorPercentTemp = Math.Abs(randomSetCurrent - randomActualCurrent);
                    double errorPercentTemp2 = errorPercentTemp / randomSetCurrent * 100;
                    double errorPercent = Math.Round(errorPercentTemp2, 1, MidpointRounding.AwayFromZero);
                    workSheet.GetRow(8).GetCell(i).SetCellValue(errorPercent + "%");
                }
            }

            //填入執行日期
            string storageGridDate = workSheet.GetRow(gridRow).GetCell(gridColumn).StringCellValue;
            storageGridDate = storageGridDate.Remove(4, 4);
            storageGridDate = storageGridDate.Insert(4, executionDate.Year.ToString());
            storageGridDate = storageGridDate.Remove(11, 2);
            storageGridDate = storageGridDate.Insert(11, executionDate.Month.ToString());
            storageGridDate = storageGridDate.Remove(16, 2);
            storageGridDate = storageGridDate.Insert(16, executionDate.Day.ToString());
            workSheet.GetRow(gridRow).GetCell(gridColumn).SetCellValue(storageGridDate);

            //勾選線別
            string storageGridLineName = workSheet.GetRow(gridRow).GetCell(0).StringCellValue;
            storageGridLineName = storageGridLineName.Remove(checkBoxIndex, 1);
            storageGridLineName = storageGridLineName.Insert(checkBoxIndex, "R");
            if (lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                storageGridLineName = storageGridLineName.Remove(checkBoxIndexAB, 1);
                storageGridLineName = storageGridLineName.Insert(checkBoxIndexAB, "R");
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
                     || lineName == "水8" || lineName == "水10" || lineName == "水12")
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
            file = new FileStream(writePath, FileMode.Create, FileAccess.Write);
            workBook.Write(file, true);
            Console.WriteLine(GetFileName(file.Name) + "寫入成功");
            workBook.Close();
            file.Close();
        }
        static void DoAppointmentMaintenanceFormExcelFile(string dirPath, string folderPath, DateTime date, string month)
        {
            try
            {
                //開啟Excel 2003檔案
                FileInfo[] directoryGFiles = new FileInfo[] { };
                FileInfo[] directoryAFiles = new FileInfo[] { };
                if (Directory.Exists(dirPath))
                {
                    // 取得資料夾內所有檔案
                    DirectoryInfo directoryInfo = new DirectoryInfo(dirPath);
                    //所有G開頭的EXCLE檔
                    directoryGFiles = directoryInfo.GetFiles("G*.xls");
                    directoryAFiles = directoryInfo.GetFiles("A*.xls");
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
                foreach (FileInfo directoryFile in directoryGFiles)
                {
                    if (File.Exists(dirPath + directoryFile.Name))
                    {
                        FileStream file;
                        IWorkbook workBook;
                        ISheet workSheet;
                        file = new FileStream(dirPath + directoryFile.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                        workBook = new HSSFWorkbook(file);
                        workSheet = workBook.GetSheetAt(0);
                        HSSFSimpleShape circle1;
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
                                SetCellStyle(workBook, workSheet, i);
                        }
                        //抓取表單的名子
                        string[] formName = directoryFile.Name.Split('-');
                        //文坦讀孔機
                        if (formName[0] == "G18")
                        {
                            MachineCodeDrowingCircle("G19", 693, 993, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G20", 8, 238, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G21", 310, 540, workBook, folderPath, directoryFile);
                            //For G18
                            DrowingCircle(false, workBook, workSheet, 28, 304, 604, 18);
                        }
                        //PLASMA
                        else if (formName[0] == "G26")
                        {
                            MachineCodeDrowingCircle("G28", 398, 625, workBook, folderPath, directoryFile);
                            MachineCodeDrowingCircle("G29", 702, 932, workBook, folderPath, directoryFile);
                            //For G26
                            DrowingCircle(false, workBook, workSheet, 28, 99, 328, 26);
                            DrowingCircle(false, workBook, workSheet, 1, 255, 312, 26);
                        }
                        SetPrintStyle(workSheet);
                        file = new FileStream(folderPath + directoryFile.Name, FileMode.Create, FileAccess.Write);
                        workBook.Write(file, true);
                        Console.WriteLine(GetFileName(file.Name) + "寫入成功");
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
