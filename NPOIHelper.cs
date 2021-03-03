using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using System.Linq;
using NPOI.SS.Util;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using log4net;

namespace ExcelInNOut
{
    public class NPOIHelper
    {
        private static ILog loger = LogManager.GetLogger(" myLogger ");

#region 從DataTable導出到excel文件中，支援xls和xlsx格式
#region 導出為xls文件內部方法
        ///  <summary> 
        ///從datatable中導出到excel
        ///  </summary> 
        ///  <param name="strFileName"> excel文件名</param> 
        ///  <param name="dtSource"> datatabe源數據</param> 
        ///  <param name="strHeaderText">表名</param> 
        ///  <param name="sheetnum"> sheet的編號</param> 
        ///  <returns></returns> 
        static MemoryStream ExportDT(String strFileName, DataTable dtSource, string strHeaderText, Dictionary<string, string> dir, int sheetnum)
        {
            //創建工作簿和sheet 
            IWorkbook workbook = new HSSFWorkbook();
            using (Stream writefile = new FileStream(strFileName, FileMode.OpenOrCreate, FileAccess.Read))
            {
                if (writefile.Length > 0 && sheetnum > 0)
                {
                    workbook = WorkbookFactory.Create(writefile);
                }
            }

            ISheet sheet = null;
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat(" yyyy-mm-dd ");
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            foreach (DataRow row in dtSource.Rows)
            {
#region 新建表，填充表頭，填充列頭，樣式
                if (rowIndex == 0)
                {
                    string sheetName = strHeaderText + (sheetnum == 0 ? "" : sheetnum.ToString());
                    if (workbook.GetSheetIndex(sheetName) >= 0)
                    {
                        workbook.RemoveSheetAt(workbook.GetSheetIndex(sheetName));
                    }
                    sheet = workbook.CreateSheet(sheetName);
#region 表頭及樣式
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        IRow headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = Horizo​​ntalAlignment.Center;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;

                        rowIndex = 1;
                    }
#endregion

#region 列頭及樣式

                    if (rowIndex == 1)
                    {
                        IRow headerRow = sheet.CreateRow(1); //第二行設置列名
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = Horizo​​ntalAlignment.Center;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        //寫入列標題
                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(dir[column.ColumnName]);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            //設置列寬
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256 * 2);
                        }
                        rowIndex = 2;
                    }
#endregion
                }
#endregion

#region 填充內容

                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);
                    string drValue = row[column].ToString();
                    switch (column.DataType.ToString())
                    {
                        case " System.String ": //字符串類型
                            double result;
                            if (isNumeric(drValue, out result))
                            {
                                //數字字符串
                                double.TryParse(drValue, out result);
                                newCell.SetCellValue(result);
                                break;
                            }
                            else
                            {
                                newCell.SetCellValue(drValue);
                                break;
                            }

                        case " System.DateTime ": //日期類型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle; //格式化顯示
                            break;
                        case " System.Boolean ": //布爾型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case " System.Int16 ": //整型
                        case " System.Int32 ":
                        case " System.Int64 ":
                        case " System.Byte ":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case " System.Decimal ": //浮點型
                        case " System.Double ":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case " System.DBNull ": //空值處理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue(drValue.ToString());
                            break;
                    }

                }
#endregion
                rowIndex++;
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                return ms;
            }

        }
    #endregion

#region 導出為xlsx文件內部方法
    ///  <summary> 
    ///從datatable中導出到excel
     ///  </summary> 
    ///  <param name="dtSource"> datatable數據源</param> 
    ///  <param name="strHeaderText">表名</param> 
    ///  <param name="fs">文件流</param> 
    ///  <param name="readfs">內存流</param> 
    ///  <param name="sheetnum"> sheet索引</param> 
    static void ExportDTI(DataTable dtSource, string strHeaderText, FileStream fs, MemoryStream readfs, Dictionary<string, string> dir, int sheetnum)
        {

            IWorkbook workbook = new XSSFWorkbook();
            if (readfs.Length > 0 && sheetnum > 0)
            {
                workbook = WorkbookFactory.Create(readfs);
            }
            ISheet sheet = null;
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat(" yyyy-mm-dd ");

            //取得列寬
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;

            foreach (DataRow row in dtSource.Rows)
            {
#region 新建表，填充表頭，填充列首，樣式

                if (rowIndex == 0)
                {
#region 表頭及樣式
                    {
                        string sheetName = strHeaderText + (sheetnum == 0 ? "" : sheetnum.ToString());
                        if (workbook.GetSheetIndex(sheetName) >= 0)
                        {
                            workbook.RemoveSheetAt(workbook.GetSheetIndex(sheetName));
                        }
                        sheet = workbook.CreateSheet(sheetName);
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        IRow headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = Horizo​​ntalAlignment.Center;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;
                    }
#endregion

#region 列首及樣式
                    {
                        IRow headerRow = sheet.CreateRow(1);
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = Horizo​​ntalAlignment.Center;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);


                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(dir[column.ColumnName]);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            //設置列寬
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256 * 2);
                        }
                    }

#endregion

                    rowIndex = 2;
                }
#endregion

#region 填充內容
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);
                    string drValue = row[column].ToString();
                    switch (column.DataType.ToString())
                    {
                        case " System.String ": //字符串類型
                            double result;
                            if (isNumeric(drValue, out result))
                            {

                                double.TryParse(drValue, out result);
                                newCell.SetCellValue(result);
                                break;
                            }
                            else
                            {
                                newCell.SetCellValue(drValue);
                                break;
                            }
                        case " System.DateTime ": //日期類型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle; //格式化顯示
                            break;
                        case " System.Boolean ": //布爾型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case " System.Int16 ": //整型
                        case " System.Int32 ":
                        case " System.Int64 ":
                        case " System.Byte ":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case " System.Decimal ": //浮點型
                        case " System.Double ":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case " System.DBNull ": //空值處理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue(drValue.ToString());
                            break;
                    }
                }
#endregion
                rowIndex++;
            }
            workbook.Write(fs);
            fs.Close();
        }
#endregion

#region 導出excel表格
        ///  <summary> 
        ///   DataTable導出到Excel文件，xls文件
        ///  </summary> 
        ///  <param name="dtSource">數據源</param> 
        ///  < param name="strHeaderText">表名</param> 
        ///  <param name="strFileName"> excel文件名</param> 
        ///  <param name="dir"> datatable和excel列名對應字典< /param> 
        ///  <param name="sheetRow">每個sheet存放的行數</param> 
        public static void ExportDTtoExcel(DataTable dtSource, string strHeaderText, string strFileName, Dictionary<string, string> dir, bool isNew, int sheetRow = 50000)
        {
            int currentSheetCount = GetSheetNumber(strFileName); //現有的頁數sheetnum 
            if (sheetRow <= 0)
            {
                sheetRow = dtSource.Rows.Count;
            }
            string[] temp = strFileName.Split('.');
            string fileExtens = temp[temp.Length - 1];
            int sheetCount = (int)Math.Ceiling((double)dtSource.Rows.Count / sheetRow); // sheet數目
            if (temp[temp.Length - 1] == " xls " && dtSource.Columns.Count < 256 && sheetRow < 65536)
            {
                if (isNew)
                {
                    currentSheetCount = 0;
                }
                for (int i = currentSheetCount; i < currentSheetCount + sheetCount; i++)
                {
                    DataTable pageDataTable = dtSource.Clone();
                    int hasRowCount = dtSource.Rows.Count - sheetRow * (i - currentSheetCount) < sheetRow ? dtSource.Rows.Count - sheetRow * (i - currentSheetCount) : sheetRow;
                    for (int j = 0; j < hasRowCount; j++)
                    {
                        pageDataTable.ImportRow(dtSource.Rows[(i - currentSheetCount) * sheetRow + j]);
                    }

                    using (MemoryStream ms = ExportDT(strFileName, pageDataTable, strHeaderText, dir, i))
                    {
                        using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                        {

                            byte[] data = ms.ToArray();
                            fs.Write(data, 0, data.Length);
                            fs.Flush();
                        }
                    }
                }
            }
            else
            {
                if (temp[temp.Length - 1] == " xls ")
                    strFileName = strFileName + " x ";
                if (isNew)
                {
                    currentSheetCount = 0;
                }
                for (int i = currentSheetCount; i < currentSheetCount + sheetCount; i++)
                {
                    DataTable pageDataTable = dtSource.Clone();
                    int hasRowCount = dtSource.Rows.Count - sheetRow * (i - currentSheetCount) < sheetRow ? dtSource.Rows.Count - sheetRow * (i - currentSheetCount) : sheetRow;
                    for (int j = 0; j < hasRowCount; j++)
                    {
                        pageDataTable.ImportRow(dtSource.Rows[(i - currentSheetCount) * sheetRow + j]);
                    }
                    FileStream readfs = new FileStream(strFileName, FileMode.OpenOrCreate, FileAccess.Read);
                    MemoryStream readfsm = new MemoryStream();
                    readfs.CopyTo(readfsm);
                    readfs.Close();
                    using (FileStream writefs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                    {

                        ExportDTI(pageDataTable, strHeaderText, writefs, readfsm, dir, i);
                    }
                    readfsm.Close();
                }
            }
        }
#endregion
#endregion

#region 從excel文件中將數據導出到datatable/datatable
        ///  <summary> 
        ///將製定sheet中的數據導出到datatable中
        ///  </summary> 
        ///  <param name="sheet">需要導出的sheet </param> 
        ///  <param name="HeaderRowIndex">列頭所在行號，-1表示沒有列頭</param> 
        ///  <param name="dir"> excel列名和DataTable列名的對應字典</param> 
        ///  <returns></returns> 
        static DataTable ImportDt(ISheet sheet, int HeaderRowIndex, Dictionary<string, string> dir)
        {
            DataTable table = new DataTable();
            IRow headerRow;
            int cellCount;
            try
            {
                //沒有標頭或者不需要表頭用excel列的序號（1,2,3..）作為DataTable的列名
                if (HeaderRowIndex < 0)
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = headerRow.LastCellNum;

                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        DataColumn column = new DataColumn(Convert.ToString(i));
                        table.Columns.Add(column);
                    }
                }
                //有表頭，使用表頭做為DataTable的列名
                else
                {
                    headerRow = sheet.GetRow(HeaderRowIndex);
                    cellCount = headerRow.LastCellNum;
                    for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                    {
                        //如果excel某一列列名不存在：以該列的序號作為Datatable的列名，如果DataTable中包含了這個序列為名的列，那麼列名為重複列名+序號
                        if (headerRow.GetCell(i) == null)
                        {
                            if (table.Columns.IndexOf(Convert.ToString(i)) > 0)
                            {
                                DataColumn column = new DataColumn(Convert.ToString("重複列名" + i));
                                table.Columns.Add(column);
                            }
                            else
                            {
                                DataColumn column = new DataColumn(Convert.ToString(i));
                                table.Columns.Add(column);
                            }

                        }
                        // excel中的某一列列名不為空，但是重複了：對應的Datatable列名為“重複列名+序號” 
                        else if (table.Columns.IndexOf(headerRow.GetCell(i).ToString()) > 0)
                        {
                            DataColumn column = new DataColumn(Convert.ToString("重複列名" + i));
                            table.Columns.Add(column);
                        }
                        else
                        //正常情況，列名存在且不重複：用excel中的列名作為datatable中對應的列名
                        {
                            string colName = dir.Where(s => s.Value == headerRow.GetCell(i).ToString()).First().Key;
                            DataColumn column = new DataColumn(colName);
                            table.Columns.Add(column);
                        }
                    }
                }
                int rowCount = sheet.LastRowNum;
                for (int i = (HeaderRowIndex + 1); i <= sheet.LastRowNum; i++) // excel行遍歷
                {
                    try
                    {
                        IRow row;
                        if (sheet.GetRow(i) == null) //如果excel有空行，則添加缺失的行
                        {
                            row = sheet.CreateRow(i);
                        }
                        else
                        {
                            row = sheet.GetRow(i);
                        }

                        DataRow dataRow = table.NewRow();

                        for (int j = row.FirstCellNum; j <= cellCount; j++) // excel列遍歷
                        {
                            try
                            {
                                if (row.GetCell(j) != null)
                                {
                                    switch (row.GetCell(j).CellType)
                                    {
                                        case CellType.String: //字符串
                                            string str = row.GetCell(j).StringCellValue;
                                            if (str != null && str.Length > 0)
                                            {
                                                dataRow[j] = str.ToString();
                                            }
                                            else
                                            {
                                                dataRow[j] = default(string);
                                            }
                                            break;
                                        case CellType.Numeric: //數字
                                            if (DateUtil.IsCellDateFormatted(row.GetCell(j))) //時間戳數字
                                            {
                                                dataRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                            }
                                            else
                                            {
                                                dataRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                            }
                                            break;
                                        case CellType.Boolean:
                                            dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                            break;
                                        case CellType.Error:
                                            dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                            break;
                                        case CellType.Formula: //公式
                                            switch (row.GetCell(j).CachedFormulaResultType)
                                            {
                                                case CellType.String:
                                                    string strFORMULA = row.GetCell(j).StringCellValue;
                                                    if (strFORMULA != null && strFORMULA.Length > 0)
                                                    {
                                                        dataRow[j] = strFORMULA.ToString();
                                                    }
                                                    else
                                                    {
                                                        dataRow[j] = null;
                                                    }
                                                    break;
                                                case CellType.Numeric:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).NumericCellValue);
                                                    break;
                                                case CellType.Boolean:
                                                    dataRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                                    break;
                                                case CellType.Error:
                                                    dataRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                                    break;
                                                default:
                                                    dataRow[j] = "";
                                                    break;
                                            }
                                            break;
                                        default:
                                            dataRow[j] = "";
                                            break;
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                loger.Error(exception.ToString());
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                    catch (Exception exception)
                    {
                        loger.Error(exception.ToString());
                    }
                }
            }
            catch (Exception exception)
            {
                loger.Error(exception.ToString());
            }
            return table;
        }

        ///  <summary> 
        ///讀取Excel文件特定名字sheet的內容到DataTable
        ///  </summary> 
        ///  <param name="strFileName"> excel文件路徑</param> 
        ///  <param name="sheet">需要導出的sheet </param> 
        ///  <param name="HeaderRowIndex">列頭所在行號，-1表示沒有列頭</param> 
        ///  <param name="dir "> excel列名和DataTable列名的對應字典</param> 
        ///  <returns></returns> 
        public static DataTable ImportExceltoDt(string strFileName, Dictionary<string, string> dir, string SheetName, int HeaderRowIndex = 1)
        {
            DataTable table = new DataTable();
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (file.Length > 0)
                {
                    IWorkbook wb = WorkbookFactory.Create(file);
                    ISheet isheet = wb.GetSheet(SheetName);
                    table = ImportDt(isheet, HeaderRowIndex, dir);
                    isheet = null;
                }
            }
            return table;
        }

        ///  <summary> 
        ///讀取Excel文件某一索引sheet的內容到DataTable
        ///  </summary> 
        ///  <param name="strFileName"> excel文件路徑</param> 
        ///  < param name="sheet">需要導出的sheet序號</param> 
        ///  <param name="HeaderRowIndex">列頭所在行號，-1表示沒有列頭</param> 
        ///  <param name= "dir"> excel列名和DataTable列名的對應字典</param> 
        ///  <returns></returns> 
        public static DataTable ImportExceltoDt(string strFileName, Dictionary<string, string> dir, int HeaderRowIndex = 1, int SheetIndex = 0)
        {
            DataTable table = new DataTable();
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (file.Length > 0)
                {
                    IWorkbook wb = WorkbookFactory.Create(file);
                    ISheet isheet = wb.GetSheetAt(SheetIndex);
                    table = ImportDt(isheet, HeaderRowIndex, dir);
                    isheet = null;
                }
            }
            return table;

        }
#endregion



        ///  <summary> 
        ///獲取excel文件的sheet數目
        ///  </summary> 
        ///  <param name="outputFile"></param> 
        ///  <returns></returns> 
        public static int GetSheetNumber(string outputFile)
        {
            int number = 0;
            using (FileStream readfile = new FileStream(outputFile, FileMode.OpenOrCreate, FileAccess.Read))
            {
                if (readfile.Length > 0)
                {
                    IWorkbook wb = WorkbookFactory.Create(readfile);
                    number = wb.NumberOfSheets;
                }
            }
            return number;
        }

        ///  <summary> 
        ///判斷內容是否是數字
        ///  </summary> 
        ///  <param name="message"></param> 
        ///  <param name="result"></param > 
        ///  <returns></returns> 
        public static bool isNumeric(String message, out double result)
        {
            Regex rex = new Regex(@" ^[-]?\d+[.]?\d*$ ");
            result = -1;
            if (rex.IsMatch(message))
            {
                result = double.Parse(message);
                return true;
            }
            else
                return false;
        }

        ///  <summary> 
        ///驗證導入的Excel是否有數據
        ///  </summary> 
        ///  <param name="excelFileStream"></param> 
        ///  <returns></returns> 
        public static bool HasData(Stream excelFileStream)
        {
            using (excelFileStream)
            {
                IWorkbook workBook = new HSSFWorkbook(excelFileStream);
                if (workBook.NumberOfSheets > 0)
                {
                    ISheet sheet = workBook.GetSheetAt(0);
                    return sheet.PhysicalNumberOfRows > 0;
                }
            }
            return false;
        }
    }
}
