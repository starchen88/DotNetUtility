using NPOI.HSSF.UserModel;//excel 2003
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;//excel 2007
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace DotNetUtility
{
    public static class NPOIHelper
    {
        #region Read Excel
        public static DataSet ExcelToDataset(string file)
        {
            var workbook = WorkbookFactory.Create(file);
            return WorkbookToDataset(workbook);
        }
        public static DataSet ExcelToDataset(Stream stream)
        {
            var workbook = WorkbookFactory.Create(stream);
            return WorkbookToDataset(workbook);
        }
        public static DataSet WorkbookToDataset(IWorkbook workbook)
        {
            DataSet result = new DataSet();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                var dt = SheetToDatatable(sheet);
                result.Tables.Add(dt);
            }
            return result;
        }
        #endregion
        #region Write Excel
        public static DataTable SheetToDatatable(ISheet sheet)
        {
            int rowIndex = 0;
            DataTable table = new DataTable(sheet.SheetName);
            var row = sheet.GetRow(rowIndex);
            for (int i = 0; i < row.LastCellNum; i++)
            {
                var cell = row.GetCell(i);
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                table.Columns.Add(cell.StringCellValue);
            }
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                row = sheet.GetRow(i);
                if (row == null || row.Cells.All(cell => cell == null || cell.CellType == CellType.Blank))
                {
                    break;
                }

                DataRow dr = table.NewRow();
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var cell = row.GetCell(j);
                    if (cell != null && cell.CellType != CellType.Blank)
                    {
                        object value;
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    value = cell.DateCellValue;
                                }
                                else
                                {
                                    value = cell.NumericCellValue;
                                }
                                break;

                            case CellType.Boolean:
                                value = cell.BooleanCellValue;
                                break;
                            case CellType.Formula:
                                value = cell.CellFormula;
                                break;
                            default:
                                value = cell.StringCellValue;
                                break;
                        }
                        dr[j] = value;
                    }
                }
                table.Rows.Add(dr);
            }
            return table;

        }
        public static IWorkbook DatasetToWorkbook(DataSet ds, ExcelFormat excelFormat)
        {
            IWorkbook workbook;
            switch (excelFormat)
            {
                case ExcelFormat.Xls:
                    workbook = new HSSFWorkbook();
                    break;
                case ExcelFormat.Xlsx:
                    workbook = new XSSFWorkbook();
                    break;
                default:
                    throw new NotSupportedException("不支持的ExcelFormat");
            }
            foreach (DataTable dt in ds.Tables)
            {
                var sheet = workbook.CreateSheet(dt.TableName);
                DataTableToSheet(dt, sheet);
            }
            return workbook;
        }
        public static void DataTableToSheet(DataTable dt, ISheet sheet)
        {
            var row = sheet.CreateRow(0);
            IDataFormat format = sheet.Workbook.CreateDataFormat();
            short dateFormat = format.GetFormat("yyyy-MM-dd");
            ICellStyle dateCellStyle = sheet.Workbook.CreateCellStyle();
            dateCellStyle.DataFormat = dateFormat;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var column = dt.Columns[i];
                var cell = row.CreateCell(i, CellType.String);
                cell.SetCellValue(column.ColumnName);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell;
                    var column = dt.Columns[j];
                    //cell.SetCellValue(dt.Rows[i][j].ToString());
                    CellType ct;

                    switch (Type.GetTypeCode(column.DataType))
                    {
                        case TypeCode.Byte:
                        case TypeCode.SByte:
                        case TypeCode.UInt16:
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            ct = CellType.Numeric;
                            cell = row.CreateCell(j, ct);
                            if (dt.Rows[i][j] != null && dt.Rows[i][j] != DBNull.Value)
                            {
                                cell.SetCellValue(Convert.ToDouble(dt.Rows[i][j]));
                            }

                            break;

                        case TypeCode.Empty:
                        case TypeCode.DBNull:
                            ct = CellType.Blank;
                            cell = row.CreateCell(j, ct);
                            break;
                        case TypeCode.Char:

                        case TypeCode.String:
                            ct = CellType.String;
                            cell = row.CreateCell(j, ct);
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                            break;
                        case TypeCode.DateTime:                            
                            ct = CellType.Numeric;
                            cell = row.CreateCell(j, ct);
                            if (dt.Rows[i][j] != null && dt.Rows[i][j] != DBNull.Value)
                            {
                                cell.SetCellValue(Convert.ToDateTime(dt.Rows[i][j]));
                            }
                            cell.CellStyle = dateCellStyle;
                            break;
                        case TypeCode.Boolean:
                            ct = CellType.Boolean;
                            cell = row.CreateCell(j, ct);
                            if (dt.Rows[i][j] != null && dt.Rows[i][j] != DBNull.Value)
                            {
                                cell.SetCellValue(Convert.ToBoolean(dt.Rows[i][j]));
                            }
                            break;
                        case TypeCode.Object:
                        default:
                            ct = CellType.String;
                            cell = row.CreateCell(j, ct);
                            cell.SetCellValue(Convert.ToString(dt.Rows[i][j]));
                            break;
                    }

                }
            }
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sheet.AutoSizeColumn(i);

            }
        }
        public static void DatasetToExcel(DataSet ds, string fileName)
        {
            var ext = Path.GetExtension(fileName).ToLower();
            ExcelFormat excelFormat;
            switch (ext)
            {
                case ".xls":
                    excelFormat = ExcelFormat.Xls;
                    break;
                case ".xlsx":
                    excelFormat = ExcelFormat.Xlsx;
                    break;
                default:
                    throw new NotSupportedException("不支持的excel文件类型");
            }
            DatasetToExcel(ds, fileName, excelFormat);
        }
        public static void DatasetToExcel(DataSet ds, string fileName, ExcelFormat excelFormat)
        {
            var ext = Path.GetExtension(fileName).ToLower();
            var workbook = DatasetToWorkbook(ds, excelFormat);
            using (var stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(stream);
            }
        }
        public static void DatasetToExcel(DataSet ds, Stream stream, ExcelFormat excelFormat)
        {
            var workbook = DatasetToWorkbook(ds, excelFormat);
            workbook.Write(stream);
        }
        #endregion
    }
    public enum ExcelFormat
    {
        /// <summary>
        /// Excel 97-2003格式
        /// </summary>
        Xls = 1997,
        /// <summary>
        /// Excel 2007格式
        /// </summary>
        Xlsx = 2007
    }
}
