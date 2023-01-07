using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using System.Data;

namespace EXCELforCPWork
{
    internal class Program
    {
        static string[] fileName = new string[3] { "保養表", "附件", "預保養表" };
        static int maintenanceFormCount, attachmentCount, appointmentMaintenanceFormCount;
        static void Main(string[] args)
        {
            for (int monthToAdd = 0; monthToAdd < 2; monthToAdd++)
            {
                maintenanceFormCount = 0;
                appointmentMaintenanceFormCount = 0;
                attachmentCount = 0;
                DateTime date = DateTime.Now.AddMonths(monthToAdd);
                string year = date.ToString("yyyy");
                string month = date.ToString("MM");
                //string dirPath = @"H:\ChinPoonWork\";
                string dirPath = System.IO.Directory.GetCurrentDirectory() + @"\";
                string dirPathNewFolder = dirPath + year + "年" + month + "月";
                //string dirPathMaintenanceForm = dirPathNewFolder + @"\保養表\";
                //string dirPathAppointmentMaintenanceForm = dirPathNewFolder + @"\後三月預保養表\";
                //string dirPathAttachment = dirPathNewFolder + @"\附件\";

                //產生需要的資料夾
                CreateFolder(dirPathNewFolder);
                if (month.Substring(0, 1) == "0")
                    month = month.Remove(0, 1);

                //製作保養表及產生相關附件
                DoMaintenanceFormExcelFile(dirPath, dirPathNewFolder, date);
                //製作預保養表
                DoAppointmentMaintenanceFormExcelFile(dirPath, dirPathNewFolder, date);
            }
            Console.ReadLine();
        }
        static void CreateFolder(string dirPathNewFolder)
        {
            //建立資料夾，以月份區分
            if (!Directory.Exists(dirPathNewFolder))
            {
                Directory.CreateDirectory(dirPathNewFolder);

                //建立保養表、後三月預保養及附件資料夾
                //Directory.CreateDirectory(dirPathMaintenanceForm);
                //Directory.CreateDirectory(dirPathAppointmentMaintenanceForm);
                //Directory.CreateDirectory(dirPathAttachment);
                Console.WriteLine("資料夾創建成功");
            }
            //創建"保養表", "附件", "預保養表"空白檔案            
            foreach (string name in fileName)
            {
                IWorkbook workBook = new HSSFWorkbook(); ;
                workBook.CreateSheet();
                FileStream file = new FileStream(dirPathNewFolder + @"\" + name + ".xls", FileMode.Create, FileAccess.Write);
                workBook.Write(file, true);
                Console.WriteLine(GetFileName(file.Name) + "寫入成功");
                workBook.Close();
                file.Close();
            }
        }

        static void DoMaintenanceFormExcelFile(string dirPath, string dirPathNewFolder, DateTime date)
        {
            try
            {
                string month = date.ToString("MM");
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
                int monthInteger = StringToInt(date.ToString("MM"));

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
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G04", "DEBURR#1");
                                break;
                            //PTH#4，第2個星期三保
                            case "G22":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Wednesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G22", "PTH#4");
                                break;
                            //PTH#5，第2個星期一保
                            case "G05":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G05", "PTH#5");
                                break;
                            //水5，第2個星期一保
                            case "G07":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Monday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G07", "水5");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G07", "水5");
                                if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G07", "水5");
                                }
                                break;
                            //PTH#6，第2個星期二保
                            case "G06":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G06", "PTH#6");
                                break;
                            //水6，第2個星期二保
                            case "G08":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Tuesday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G08", "水6");
                                if(heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G08", "水6");
                                if (DoCurrentCheckForm(workSheet.GetRow(9).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G08", "水6");
                                }
                                break;
                            //水7，第1個星期四保
                            case "G09":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Thursday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G09", "水7");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G09", "水7");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G09", "水7");
                                }
                                break;
                            //水8，第1個星期一保
                            case "G10":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Monday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G10", "水8");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G10", "水8");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Monday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G10", "水8");
                                }
                                break;
                            //水9，第2個星期四保
                            case "G11":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Thursday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G11", "水9");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G11", "水9");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Thursday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G11", "水9");
                                }
                                break;
                            //水10，第1個星期三保
                            case "G12":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Wednesday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G12", "水10");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G12", "水10");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Wednesday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G12", "水10");
                                }
                                break;
                            //雷射孔微蝕#2，第3個星期二保
                            case "G13":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 3, "Tuesday");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G13", "雷射孔微蝕#2");
                                break;
                            //文坦讀孔機，第2個星期日保
                            case "G18":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 2, "Sunday");
                                break;
                            //水11，第1個星期二保
                            case "G24":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Tuesday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G24", "水11");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G24", "水11");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Tuesday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G24", "水11");
                                }
                                break;
                            //水12，第1個星期五保
                            case "G25":
                                executionDate = DateToWeekDay(monthFirstDay, daysOfMonth, 1, "Friday");
                                DoForm_A01(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G25", "水12");
                                if (heaterCheck)
                                    DoForm_A02ToA06(dirPath, dirPathNewFolder, directoryAFiles, executionDate, "G25", "水12");
                                if (DoCurrentCheckForm(workSheet.GetRow(10).GetCell(6).ToString(), date))
                                {
                                    nextMonthExecutionDate = DateToWeekDay(nextMonthFirstDay, 28, 1, "Friday");
                                    DoForm_A07A08(dirPath, dirPathNewFolder, directoryAFiles, nextMonthExecutionDate, "G25", "水12");
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
                            MachineCodeDrowingCircle("G19", 693, 993, workBook, dirPathNewFolder, directoryFile);
                            MachineCodeDrowingCircle("G20", 8, 238, workBook, dirPathNewFolder, directoryFile);
                            MachineCodeDrowingCircle("G21", 310, 540, workBook, dirPathNewFolder, directoryFile);
                            //For G18
                            DrowingCircle(false, workBook, workSheet, 28, 304, 604, 18);
                        }
                        //PLASMA
                        else if (formName[0] == "G26")
                        {
                            MachineCodeDrowingCircle("G28", 398, 625, workBook, dirPathNewFolder, directoryFile);
                            MachineCodeDrowingCircle("G29", 702, 932, workBook, dirPathNewFolder, directoryFile);
                            //For G26
                            DrowingCircle(false, workBook, workSheet, 28, 99, 328, 26);
                            DrowingCircle(false, workBook, workSheet, 1, 255, 312, 26);
                        }

                        SetPrintStyle(workSheet);

                        file = new FileStream(dirPathNewFolder + @"\" + directoryFile.Name, FileMode.Create, FileAccess.Write);
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
            FileStream newFile = new FileStream(folderPath + @"\" + machineCode + "-" + directoryFile.Name, FileMode.Create, FileAccess.Write);
            workBook.Write(newFile, true);
            circle.RemoveShape(c1);
            if (machineCodeNumber == 28 || machineCodeNumber == 29) 
                circle2.RemoveShape(c2);
        }
        static void DoForm_A01(string dirPath, string dirPathNewFolder, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {
            FileStream readFile = new FileStream(dirPath + directoryAFiles[0].Name, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = new HSSFWorkbook(readFile);
            ISheet workSheet = workBook.GetSheetAt(0);
            readFile.Close();
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

            FileStream writeFile = new FileStream(dirPathNewFolder + @"/A01-" + machineCode + "-" + lineName + "-亞碩競銘線纜線熱顯像檢查表.xls", FileMode.Create, FileAccess.Write);
            workBook.Write(writeFile, true);
            Console.WriteLine(GetFileName(writeFile.Name) + "寫入成功");
            workBook.Close();
            writeFile.Close();
        }
        static void DoForm_A02ToA06(string dirPath, string dirPathNewFolder, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {

            string machineName = "";
            string openPath = "", writePath = "";
            string openPath2 = "", writePath2 = "";
            //FOR DEBURR#1
            if (lineName == "DEBURR#1")
            {
                machineName = "DEBURR(1)線";
                openPath = dirPath + directoryAFiles[5].Name;
                writePath = dirPathNewFolder + @"\A05-" + machineCode + "-" + lineName + "-DEBURR設備性能檢測數值記錄表.xls";
            }
            //FOR 雷射孔微蝕#2
            else if (lineName == "雷射孔微蝕#2")
            {
                machineName = "雷射孔微蝕(2)線";
                openPath = dirPath + directoryAFiles[6].Name;
                writePath = dirPathNewFolder + @"\A06-" + machineCode + "-" + lineName + "-雷射孔微蝕設備性能檢測數值記錄表.xls";
            }
            //FOR VCP
            else if (lineName == "水5" || lineName == "水6")
            {
                machineName = "水平電鍍線(VCP)(" + lineName.Remove(0,1) + ")線";
                openPath = dirPath + directoryAFiles[1].Name;
                writePath = dirPathNewFolder + @"\A02-" + machineCode + "-" + lineName + "-水平電鍍線(VCP)設備性能檢測數值記錄表.xls";
            }
            //FOR PTH
            else if (lineName == "PTH#4" || lineName == "PTH#5" || lineName == "PTH#6")
            {
                machineName = "水平PTH(" + lineName.Remove(0, 4) + ")線";
                openPath = dirPath + directoryAFiles[3].Name;
                writePath = dirPathNewFolder + @"\A041-" + machineCode + "-" + lineName + "-PTH設備性能檢測數值記錄表.xls";

                openPath2 = dirPath + directoryAFiles[4].Name;
                writePath2 = dirPathNewFolder + @"\A042-" + machineCode + "-" + lineName + "-PTH設備性能檢測數值記錄表.xls";
            }
            //FOR SVCP
            else
            {
                machineName = "水平電鍍線(SVCP)(" + lineName.Remove(0, 1) + ")線";
                openPath = dirPath + directoryAFiles[2].Name;
                writePath = dirPathNewFolder + @"\A03-" + machineCode + "-" + lineName + "-水平電鍍線(SVCP)設備性能檢測數值記錄表.xls";
            }
            FileStream readFile = new FileStream(openPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IWorkbook workBook = new HSSFWorkbook(readFile);
            ISheet workSheet = workBook.GetSheetAt(0);

            workSheet.GetRow(1).GetCell(0).SetCellValue("設備名稱:  " + machineName);
            workSheet.GetRow(1).GetCell(8).SetCellValue("檢測日期:" + executionDate[0].ToString("  yyyy   /    M    /   dd"));

            SetPrintStyle(workSheet);

            //file = new FileStream(writePath, FileMode.Create, FileAccess.Write);

            FileStream writeFile = new FileStream(dirPathNewFolder + @"\" + fileName[1] + ".xls", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            IWorkbook attachmentWorkBook = new HSSFWorkbook(writeFile);

            workSheet.CopyTo(attachmentWorkBook, lineName, true, true);

            attachmentWorkBook.Write(writeFile, true);
            Console.WriteLine(GetFileName(writeFile.Name) + "寫入成功");
            /*
            if (lineName == "PTH#4" || lineName == "PTH#5" || lineName == "PTH#6")
            {
                readFile = new FileStream(openPath2, FileMode.Open, FileAccess.Read);
                workBook = new HSSFWorkbook(readFile);
                workSheet = workBook.GetSheetAt(0);

                workSheet.GetRow(1).GetCell(0).SetCellValue("設備名稱:  " + machineName);
                workSheet.GetRow(1).GetCell(8).SetCellValue("檢測日期:" + executionDate[0].ToString("  yyyy   /    M    /   dd"));

                SetPrintStyle(workSheet);

                writeFile = new FileStream(writePath2, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                attachmentWorkBook.Write(writeFile, true);
                Console.WriteLine(GetFileName(writeFile.Name) + "寫入成功");
            }
            */
            workBook.Close();
            readFile.Close();
            writeFile.Close();
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
        static void DoForm_A07A08(string dirPath, string dirPathNewFolder, FileInfo[] directoryAFiles, List<DateTime> executionDate, string machineCode, string lineName)
        {
            //設定開啟及儲存的路徑跟檔名
            string openPath = "", writePath = "", writePath2 = "";
            //FOR PTH#5、PTH#6
            if (lineName == "水5" || lineName == "水6")
            {
                openPath = dirPath + directoryAFiles[7].Name;
                writePath = dirPathNewFolder + @"/A07-" + machineCode + "-" + lineName + "-PTH電流比對紀錄表.xls";
            }
            //FOR 奇數水平電鍍線(水7、水9、水11)
            else if (lineName == "水7" || lineName == "水9" || lineName == "水11")
            {
                openPath = dirPath + directoryAFiles[8].Name;
                writePath = dirPathNewFolder + @"/A08-" + machineCode + "-" + lineName + "-水平電鍍線電流比對紀錄表.xls";
            }
            //FOR 偶數水平電鍍線(水8、水10、水12)
            else if (lineName == "水8" || lineName == "水10" || lineName == "水12")
            {
                openPath = dirPath + directoryAFiles[8].Name;
                writePath = dirPathNewFolder + @"/A08-" + machineCode + "-" + lineName + "A-水平電鍍線電流比對紀錄表.xls";
                writePath2 = dirPathNewFolder + @"/A08-" + machineCode + "-" + lineName + "B-水平電鍍線電流比對紀錄表.xls";
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
            FileStream readFile = new FileStream(openPath, FileMode.Open, FileAccess.Read);
            IWorkbook workBook = new HSSFWorkbook(readFile);
            ISheet workSheet = workBook.GetSheetAt(0);
            readFile.Close();
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
            workBook.SetSheetName(0, lineName);
            FileStream writeFile = new FileStream(writePath, FileMode.Create, FileAccess.Write);
            workBook.Write(writeFile, true);
            Console.WriteLine(GetFileName(writeFile.Name) + "寫入成功");
            workBook.Close();
            writeFile.Close();
        }
        static void DoAppointmentMaintenanceFormExcelFile(string dirPath, string folderPath, DateTime date)
        {
            string month = date.ToString("MM");
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
                        FileStream readFile = new FileStream(dirPath + directoryFile.Name, FileMode.Open, FileAccess.Read);
                        IWorkbook workBook = new HSSFWorkbook(readFile);
                        ISheet workSheet = workBook.GetSheetAt(0);
                        readFile.Close();

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
                        FileStream writeFile = new FileStream(folderPath + @"\" + directoryFile.Name, FileMode.Create, FileAccess.Write);
                        workBook.Write(writeFile, true);
                        Console.WriteLine(GetFileName(writeFile.Name) + "寫入成功");
                        workBook.Close();
                        writeFile.Close();
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
        /// <summary>
        /// 复制sheet
        /// </summary>
        /// <param name="bjDt">sheet名集合</param>
        /// <param name="modelfilename">模板附件名</param>
        /// <param name="tpath">生成文件路径</param>
        /// <returns></returns>
        public HSSFWorkbook SheetCopy(DataTable bjDt, string modelfilename, out string tpath)
        {
            string templetfilepath = @"files\" + modelfilename + ".xls";//模版Excel

            tpath = @"files\download\" + modelfilename + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";//中介Excel，以它为中介来导出，避免直接使用模块Excel而改变模块的格式
            FileInfo ff = new FileInfo(tpath);
            if (ff.Exists)
            {
                ff.Delete();
            }
            FileStream fs = File.Create(tpath);//创建中间excel
            HSSFWorkbook x1 = new HSSFWorkbook();
            x1.Write(fs);
            fs.Close();
            FileStream fileRead = new FileStream(templetfilepath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook hssfworkbook = new HSSFWorkbook(fileRead);
            FileStream fileSave2 = new FileStream(tpath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook book2 = new HSSFWorkbook(fileSave2);
            HSSFWorkbook[] book = new HSSFWorkbook[2] { book2, hssfworkbook };
            HSSFSheet CPS = hssfworkbook.GetSheet("Sheet0") as HSSFSheet;//获得模板sheet
            string rsbh = bjDt.Rows[0]["name"].ToString();
            CPS.CopyTo(book2, rsbh, true, true);//将模板sheet复制到目标sheet
            HSSFSheet sheet = book2.GetSheet(bjDt.Rows[0]["name"].ToString()) as HSSFSheet;//获得当前sheet
            for (int i = 1; i < bjDt.Rows.Count; i++)
            {
                sheet.CopySheet(bjDt.Rows[i]["name"].ToString(), true);//将sheet复制到同一excel的其他sheet上
            }
            return book2;
        }
    }
    /*
    public static class NPOIExt
    {    /// 
             /// 跨工作薄Workbook复制工作表Sheet    /// 
             /// 源工作表Sheet
             /// 目标工作薄Workbook
             /// 目标工作表Sheet名
             /// 是否复制打印设置
        public static ISheet CrossCloneSheet(this ISheet sSheet, IWorkbook dWb, string dSheetName, bool clonePrintSetup)
        {
            ISheet dSheet;
            dSheetName = string.IsNullOrEmpty(dSheetName) ? sSheet.SheetName : dSheetName;
            dSheetName = (dWb.GetSheet(dSheetName) == null) ? dSheetName : dSheetName + "_拷贝";
            dSheet = dWb.GetSheet(dSheetName) ?? dWb.CreateSheet(dSheetName);
            CopySheet(sSheet, dSheet); if (clonePrintSetup)
                ClonePrintSetup(sSheet, dSheet);
            dWb.SetActiveSheet(dWb.GetSheetIndex(dSheet));  //当前Sheet作为下次打开默认Sheet
            return dSheet;
        }    /// 
                 /// 跨工作薄Workbook复制工作表Sheet    /// 
                 /// 源工作表Sheet
                 /// 目标工作薄Workbook
                 /// 目标工作表Sheet名
        public static ISheet CrossCloneSheet(this ISheet sSheet, IWorkbook dWb, string dSheetName)
        {
            bool clonePrintSetup = true; return CrossCloneSheet(sSheet, dWb, dSheetName, clonePrintSetup);
        }    /// 
                 /// 跨工作薄Workbook复制工作表Sheet    /// 
                 /// 源工作表Sheet
                 /// 目标工作薄Workbook
        public static ISheet CrossCloneSheet(this ISheet sSheet, IWorkbook dWb)
        {
            string dSheetName = sSheet.SheetName; bool clonePrintSetup = true; return CrossCloneSheet(sSheet, dWb, dSheetName, clonePrintSetup);
        }
        private static IFont FindFont(this IWorkbook dWb, IFont font, List<IFont> dFonts)
        {        //IFont dFont = dWb.FindFont(font.Boldweight, font.Color, (short)font.FontHeight, font.FontName, font.IsItalic, font.IsStrikeout, font.TypeOffset, font.Underline);
            IFont dFont = null; foreach (IFont currFont in dFonts)
            {            //if (currFont.Charset != font.Charset) continue;            //else            //if (currFont.Color != font.Color) continue;            //else
                if (currFont.FontName != font.FontName) continue; else if (currFont.FontHeight != font.FontHeight) continue; else if (currFont.IsBold != font.IsBold) continue; else if (currFont.IsItalic != font.IsItalic) continue; else if (currFont.IsStrikeout != font.IsStrikeout) continue; else if (currFont.Underline != font.Underline) continue; else if (currFont.TypeOffset != font.TypeOffset) continue; else { dFont = currFont; break; }
            }
            return dFont;
        }
        private static ICellStyle FindStyle(this IWorkbook dWb, IWorkbook sWb, ICellStyle style, List<ICellStyle> dCellStyles, List<IFont> dFonts)
        {
            ICellStyle dStyle = null; foreach (ICellStyle currStyle in dCellStyles)
            {
                if (currStyle.Alignment != style.Alignment) continue;
                else if (currStyle.VerticalAlignment != style.VerticalAlignment) continue;
                else if (currStyle.BorderTop != style.BorderTop) continue;
                else if (currStyle.BorderBottom != style.BorderBottom) continue;
                else if (currStyle.BorderLeft != style.BorderLeft) continue;
                else if (currStyle.BorderRight != style.BorderRight) continue;
                else if (currStyle.TopBorderColor != style.TopBorderColor) continue;
                else if (currStyle.BottomBorderColor != style.BottomBorderColor) continue;
                else if (currStyle.LeftBorderColor != style.LeftBorderColor) continue;
                else if (currStyle.RightBorderColor != style.RightBorderColor) continue;            //else if (currStyle.BorderDiagonal != style.BorderDiagonal) continue;            //else if (currStyle.BorderDiagonalColor != style.BorderDiagonalColor) continue;            //else if (currStyle.BorderDiagonalLineStyle != style.BorderDiagonalLineStyle) continue;            //else if (currStyle.FillBackgroundColor != style.FillBackgroundColor) continue;            //else if (currStyle.FillBackgroundColorColor != style.FillBackgroundColorColor) continue;            //else if (currStyle.FillForegroundColor != style.FillForegroundColor) continue;            //else if (currStyle.FillForegroundColorColor != style.FillForegroundColorColor) continue;            //else if (currStyle.FillPattern != style.FillPattern) continue;
                else if (currStyle.Indention != style.Indention) continue;
                else if (currStyle.IsHidden != style.IsHidden) continue;
                else if (currStyle.IsLocked != style.IsLocked) continue;
                else if (currStyle.Rotation != style.Rotation) continue;
                else if (currStyle.ShrinkToFit != style.ShrinkToFit) continue;
                else if (currStyle.WrapText != style.WrapText) continue;
                else if (!currStyle.GetDataFormatString().Equals(style.GetDataFormatString())) continue;
                else
                {
                    IFont sFont = sWb.GetFontAt(style.FontIndex);
                    IFont dFont = dWb.FindFont(sFont, dFonts);
                    if (dFont == null)
                        continue;
                    else
                    {
                        currStyle.SetFont(dFont);
                        dStyle = currStyle; break;
                    }
                }
            }
            return dStyle;
        }
        private static IFont CopyFont(this IFont dFont, IFont sFont, List<IFont> dFonts)
        {        //dFont.Charset = sFont.Charset;        //dFont.Color = sFont.Color;
            dFont.FontHeight = sFont.FontHeight;
            dFont.FontName = sFont.FontName;
            dFont.IsBold = sFont.IsBold;
            dFont.IsItalic = sFont.IsItalic;
            dFont.IsStrikeout = sFont.IsStrikeout;
            dFont.Underline = sFont.Underline;
            dFont.TypeOffset = sFont.TypeOffset;
            dFonts.Add(dFont); return dFont;
        }
        private static ICellStyle CopyStyle(this ICellStyle dCellStyle, ICellStyle sCellStyle, IWorkbook dWb, IWorkbook sWb, List<ICellStyle> dCellStyles, List<IFont> dFonts)
        {
            ICellStyle currCellStyle = dCellStyle;
            currCellStyle.Alignment = sCellStyle.Alignment;
            currCellStyle.VerticalAlignment = sCellStyle.VerticalAlignment;
            currCellStyle.BorderTop = sCellStyle.BorderTop;
            currCellStyle.BorderBottom = sCellStyle.BorderBottom;
            currCellStyle.BorderLeft = sCellStyle.BorderLeft;
            currCellStyle.BorderRight = sCellStyle.BorderRight;
            currCellStyle.TopBorderColor = sCellStyle.TopBorderColor;
            currCellStyle.LeftBorderColor = sCellStyle.LeftBorderColor;
            currCellStyle.RightBorderColor = sCellStyle.RightBorderColor;
            currCellStyle.BottomBorderColor = sCellStyle.BottomBorderColor;        //dCellStyle.BorderDiagonal = sCellStyle.BorderDiagonal;        //dCellStyle.BorderDiagonalColor = sCellStyle.BorderDiagonalColor;        //dCellStyle.BorderDiagonalLineStyle = sCellStyle.BorderDiagonalLineStyle;        //dCellStyle.FillBackgroundColor = sCellStyle.FillBackgroundColor;        //dCellStyle.FillForegroundColor = sCellStyle.FillForegroundColor;        //dCellStyle.FillPattern = sCellStyle.FillPattern;
            currCellStyle.Indention = sCellStyle.Indention;
            currCellStyle.IsHidden = sCellStyle.IsHidden;
            currCellStyle.IsLocked = sCellStyle.IsLocked;
            currCellStyle.Rotation = sCellStyle.Rotation;
            currCellStyle.ShrinkToFit = sCellStyle.ShrinkToFit;
            currCellStyle.WrapText = sCellStyle.WrapText;
            currCellStyle.DataFormat = dWb.CreateDataFormat().GetFormat(sWb.CreateDataFormat().GetFormat(sCellStyle.DataFormat));
            IFont sFont = sCellStyle.GetFont(sWb);
            IFont dFont = dWb.FindFont(sFont, dFonts) ?? dWb.CreateFont().CopyFont(sFont, dFonts);
            currCellStyle.SetFont(dFont);
            dCellStyles.Add(currCellStyle); return currCellStyle;
        }
        private static void CopySheet(ISheet sSheet, ISheet dSheet)
        {
            var maxColumnNum = 0;
            List<ICellStyle> dCellStyles = new List<ICellStyle>();
            List<IFont> dFonts = new List<IFont> { };
            MergerRegion(sSheet, dSheet); for (int i = sSheet.FirstRowNum; i <= sSheet.LastRowNum; i++)
            {
                IRow sRow = sSheet.GetRow(i);
                IRow dRow = dSheet.CreateRow(i); if (sRow != null)
                {
                    CopyRow(sRow, dRow, dCellStyles, dFonts); if (sRow.LastCellNum > maxColumnNum)
                        maxColumnNum = sRow.LastCellNum;
                }
            }
            for (int i = 0; i <= maxColumnNum; i++)
                dSheet.SetColumnWidth(i, sSheet.GetColumnWidth(i));
        }
        private static void CopyRow(IRow sRow, IRow dRow, List<ICellStyle> dCellStyles, List<IFont> dFonts)
        {
            dRow.Height = sRow.Height;
            ISheet sSheet = sRow.Sheet;
            ISheet dSheet = dRow.Sheet; for (int j = sRow.FirstCellNum; j <= sRow.LastCellNum; j++)
            {
                ICell sCell = sRow.GetCell(j);
                ICell dCell = dRow.GetCell(j); if (sCell != null)
                {
                    if (dCell == null)
                        dCell = dRow.CreateCell(j);
                    CopyCell(sCell, dCell, dCellStyles, dFonts);
                }
            }
        }
        private static void CopyCell(ICell sCell, ICell dCell, List<ICellStyle> dCellStyles, List<IFont> dFonts)
        {
            ICellStyle currCellStyle = dCell.Sheet.Workbook.FindStyle(sCell.Sheet.Workbook, sCell.CellStyle, dCellStyles, dFonts); if (currCellStyle == null)
                currCellStyle = dCell.Sheet.Workbook.CreateCellStyle().CopyStyle(sCell.CellStyle, dCell.Sheet.Workbook, sCell.Sheet.Workbook, dCellStyles, dFonts);
            dCell.CellStyle = currCellStyle; switch (sCell.CellType)
            {
                case CellType.String:
                    dCell.SetCellValue(sCell.StringCellValue); break;
                case CellType.Numeric:
                    dCell.SetCellValue(sCell.NumericCellValue); break;
                case CellType.Blank:
                    dCell.SetCellType(CellType.Blank); break;
                case CellType.Boolean:
                    dCell.SetCellValue(sCell.BooleanCellValue); break;
                case CellType.Error:
                    dCell.SetCellValue(sCell.ErrorCellValue); break;
                case CellType.Formula:
                    dCell.SetCellFormula(sCell.CellFormula); break;
                default: break;
            }
        }
        private static void MergerRegion(ISheet sSheet, ISheet dSheet)
        {
            int sheetMergerCount = sSheet.NumMergedRegions; for (int i = 0; i < sheetMergerCount; i++)
                dSheet.AddMergedRegion(sSheet.GetMergedRegion(i));
        }
        private static void ClonePrintSetup(ISheet sSheet, ISheet dSheet)
        {        //工作表Sheet页面打印设置
            dSheet.PrintSetup.Copies = 1;                               //打印份数
            dSheet.PrintSetup.PaperSize = sSheet.PrintSetup.PaperSize;  //纸张大小
            dSheet.PrintSetup.Landscape = sSheet.PrintSetup.Landscape;  //纸张方向：默认纵向false(横向true)
            dSheet.PrintSetup.Scale = sSheet.PrintSetup.Scale;          //缩放方式比例
            dSheet.PrintSetup.FitHeight = sSheet.PrintSetup.FitHeight;  //调整方式页高
            dSheet.PrintSetup.FitWidth = sSheet.PrintSetup.FitWidth;    //调整方式页宽
            dSheet.PrintSetup.FooterMargin = sSheet.PrintSetup.FooterMargin;
            dSheet.PrintSetup.HeaderMargin = sSheet.PrintSetup.HeaderMargin;        //页边距        dSheet.SetMargin(MarginType.TopMargin, sSheet.GetMargin(MarginType.TopMargin));
            dSheet.SetMargin(MarginType.BottomMargin, sSheet.GetMargin(MarginType.BottomMargin));
            dSheet.SetMargin(MarginType.LeftMargin, sSheet.GetMargin(MarginType.LeftMargin));
            dSheet.SetMargin(MarginType.RightMargin, sSheet.GetMargin(MarginType.RightMargin));
            dSheet.SetMargin(MarginType.HeaderMargin, sSheet.GetMargin(MarginType.HeaderMargin));
            dSheet.SetMargin(MarginType.FooterMargin, sSheet.GetMargin(MarginType.FooterMargin));        //页眉页脚
            dSheet.Header.Left = sSheet.Header.Left;
            dSheet.Header.Center = sSheet.Header.Center;
            dSheet.Header.Right = sSheet.Header.Right;
            dSheet.Footer.Left = sSheet.Footer.Left;
            dSheet.Footer.Center = sSheet.Footer.Center;
            dSheet.Footer.Right = sSheet.Footer.Right;        //工作表Sheet参数设置
            dSheet.IsPrintGridlines = sSheet.IsPrintGridlines;          //true: 打印整表网格线。不单独设置CellStyle时外框实线内框虚线。 false: 自己设置网格线
            dSheet.FitToPage = sSheet.FitToPage;                        //自适应页面
            dSheet.HorizontallyCenter = sSheet.HorizontallyCenter;      //打印页面为水平居中
            dSheet.VerticallyCenter = sSheet.VerticallyCenter;          //打印页面为垂直居中
            dSheet.RepeatingRows = sSheet.RepeatingRows;                //工作表顶端标题行范围    }
        }
    }
    */
}
