using ExcelAbstraction.Entities;
using ExcelAbstraction.Helpers;
using ExcelAbstraction.Services;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record;
using NPOI.HSSF.Record.Aggregates;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ExcelAbstraction.NPOI
{
    
    public class ExcelService : IExcelService
    {
        public ExcelService()
        {
            this.Format = NumberFormatInfo.CurrentInfo;
        }
        
        public Stream AddAuthor(Stream stream, ExcelVersion excelVersion, string author)
        {
            return this.AddProperty(stream, excelVersion, delegate (SummaryInformation summaryInformation) {
                summaryInformation.Author = author;
            }, delegate (CoreProperties coreProperties) {
                coreProperties.Creator = author;
            });
        }
        
        private static void AddComment(ICell cell, string comment)
        {
            ISheet sheet = cell.Sheet;
            ICreationHelper creationHelper = sheet.Workbook.GetCreationHelper();
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = creationHelper.CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = (int) (cell.ColumnIndex + 3);
            anchor.Row1 = cell.RowIndex;
            anchor.Row2 = (int) (cell.RowIndex + 5);
            IComment comment2 = drawing.CreateCellComment(anchor);
            comment2.String = (creationHelper.CreateRichTextString(comment));
            cell.CellComment = comment2;
        }
        
        public Stream AddComments(Stream stream, ExcelVersion excelVersion, string comments)
        {
            return this.AddProperty(stream, excelVersion, delegate (SummaryInformation summaryInfo) {
                summaryInfo.Comments = comments;
            }, delegate (CoreProperties coreProperties) {
                coreProperties.Description = comments;
            });
        }
        
        private static void AddNames(IWorkbook workbook, ExcelVersion version, params NamedRange[] names)
        {
            foreach (NamedRange range in names)
            {
                IName name = workbook.CreateName();
                name.NameName = range.Name;
                name.RefersToFormula = ExcelHelper.RangeToString(range.Range, version);
            }
        }
        
        public void AddNames(object workbook, ExcelVersion version, params NamedRange[] names)
        {
            AddNames((IWorkbook) workbook, version, names);
        }
        
        private MemoryStream AddProperty(Stream stream, ExcelVersion excelVersion, Action<SummaryInformation> hssWorkbookAction, Action<CoreProperties> corePropertiesAction)
        {
            MemoryStream stream2 = new MemoryStream();
            if (excelVersion.Equals(ExcelVersion.Xls))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(stream);
                hssWorkbookAction(workbook.SummaryInformation);
                workbook.Write(stream2);
            }
            if (excelVersion.Equals(ExcelVersion.Xlsx))
            {
                XSSFWorkbook workbook2 = new XSSFWorkbook(stream);
                POIXMLProperties properties = workbook2.GetProperties();
                corePropertiesAction(properties.CoreProperties);
                workbook2.Write(stream2);
            }
            return new MemoryStream(stream2.ToArray());
        }
        
        private static void AddRows(ISheet sheet, params Row[] rows)
        {
            foreach (Row row in rows)
            {
                if (row != null)
                {
                    IRow row2 = sheet.CreateRow(row.Index);
                    foreach (Cell cell in row.Cells)
                    {
                        if (cell != null)
                        {
                            ICell cell2 = row2.CreateCell(cell.ColumnIndex);
                            if (!string.IsNullOrEmpty(cell.DataFormat))
                            {
                                IWorkbook workbook = sheet.Workbook;
                                IDataFormat format = workbook.CreateDataFormat();
                                cell2.CellStyle = workbook.CreateCellStyle();
                                cell2.CellStyle.DataFormat = format.GetFormat(cell.DataFormat);
                            }
                            if (!string.IsNullOrEmpty(cell.Comment))
                            {
                                AddComment(cell2, cell.Comment);
                            }
                            if (cell.Value != null)
                            {
                                cell2.SetCellValue(cell.Value);
                            }
                        }
                    }
                }
            }
        }
        
        public void AddRows(object workbook, string sheetName, params Row[] rows)
        {
            AddRows(((IWorkbook) workbook).GetSheet(sheetName), rows);
        }
        
        private static void AddToNames(ICollection<NamedRange> names, IWorkbook workbook)
        {
            string str;
            ExcelVersion xls;
            if (workbook is HSSFWorkbook)
            {
                str = "names";
                xls = ExcelVersion.Xls;
            }
            else
            {
                if (!(workbook is XSSFWorkbook))
                {
                    return;
                }
                str = "namedRanges";
                xls = ExcelVersion.Xlsx;
            }
            foreach (IName name in (IList) workbook.GetType().GetField(str, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly).GetValue(workbook))
            {
                if (name.RefersToFormula.Contains("://"))
                {
                    continue;
                }
                ExcelAbstraction.Entities.Range range = ExcelHelper.ParseRange(name.RefersToFormula, xls);
                if (range != null)
                {
                    NamedRange item = new NamedRange {
                        Name = name.NameName,
                        Range = range
                    };
                    names.Add(item);
                }
            }
        }
        
        private static void AddToValidations(ICollection<DataValidation> validations, HSSFSheet sheet, string[] names)
        {
            InternalSheet sheet2 = sheet.Sheet;
            DataValidityTable table = (DataValidityTable) sheet2.GetType().GetField("_dataValidityTable", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(sheet2);
            if (table != null)
            {
                foreach (DVRecord record in (IList) table.GetType().GetField("_validationList", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(table))
                {
                    Formula formula = (Formula) record.GetType().GetField("_formula1", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(record);
                    DataValidation item = new DataValidation {
                        Range = ExcelHelper.ParseRange(record.CellRangeAddress.CellRangeAddresses[0].FormatAsString(), ExcelVersion.Xls)
                    };
                    Ptg ptg = formula.Tokens[0];
                    NamePtg ptg2 = ptg as NamePtg;
                    if (ptg2 != null)
                    {
                        item.Type = DataValidationType.Formula;
                        item.Name = names.ElementAt<string>(ptg2.Index);
                    }
                    else
                    {
                        StringPtg ptg3 = ptg as StringPtg;
                        if (ptg3 == null)
                        {
                            continue;
                        }
                        item.Type = DataValidationType.List;
                        item.List = ptg3.Value.Split(new char[1]);
                    }
                    validations.Add(item);
                }
            }
        }
        
        private static void AddToValidations(ICollection<DataValidation> validations, ISheet sheet, string[] names)
        {
            HSSFSheet sheet2 = sheet as HSSFSheet;
            if (sheet2 != null)
            {
                AddToValidations(validations, sheet2, names);
            }
            else
            {
                XSSFSheet sheet3 = sheet as XSSFSheet;
                if (sheet3 != null)
                {
                    AddToValidations(validations, sheet3, names);
                }
            }
        }
        
        private static void AddToValidations(ICollection<DataValidation> validations, XSSFSheet sheet, string[] names)
        {
            CT_DataValidations dataValidations = ((CT_Worksheet) sheet.GetType().GetField("worksheet", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly).GetValue(sheet)).dataValidations;
            if (dataValidations != null)
            {
                foreach (CT_DataValidation validation in dataValidations.dataValidation)
                {
                    if (validation.formula1 == null)
                    {
                        continue;
                    }
                    ExcelAbstraction.Entities.Range range = ExcelHelper.ParseRange(validation.sqref, ExcelVersion.Xlsx);
                    if (range != null)
                    {
                        DataValidation item = new DataValidation {
                            Range = range
                        };
                        if (names.Contains<string>(validation.formula1))
                        {
                            item.Type = DataValidationType.Formula;
                            item.Name = validation.formula1;
                        }
                        else
                        {
                            item.Type = DataValidationType.List;
                            item.List = validation.formula1.Trim(new char[] { '"' }).Split(new char[] { ',' });
                        }
                        validations.Add(item);
                    }
                }
            }
        }
        
        private static void AddValidations(ISheet sheet, ExcelVersion version, params DataValidation[] validations)
        {
            IDataValidationHelper dataValidationHelper = sheet.GetDataValidationHelper();
            foreach (DataValidation validation in validations)
            {
                if (((validation.List == null) || (validation.List.Count == 0)) && (validation.Name == null))
                {
                    throw new InvalidOperationException("Validation is invalid");
                }
                IDataValidationConstraint constraint = (validation.Name != null) ? dataValidationHelper.CreateFormulaListConstraint(validation.Name) : dataValidationHelper.CreateExplicitListConstraint(validation.List.ToArray<string>());
                int? rowStart = validation.Range.RowStart;
                int? rowEnd = validation.Range.RowEnd;
                int? columnStart = validation.Range.ColumnStart;
                int? columnEnd = validation.Range.ColumnEnd;
                IDataValidation validation2 = dataValidationHelper.CreateValidation(constraint, new CellRangeAddressList((rowStart != null) ? rowStart.GetValueOrDefault() : 0, (rowEnd != null) ? rowEnd.GetValueOrDefault() : (ExcelHelper.GetRowMax(version) - 1), (columnStart != null) ? columnStart.GetValueOrDefault() : 0, (columnEnd != null) ? columnEnd.GetValueOrDefault() : (ExcelHelper.GetColumnMax(version) - 1)));
                sheet.AddValidationData(validation2);
            }
        }
        
        public void AddValidations(object workbook, string sheetName, ExcelVersion version, params DataValidation[] validations)
        {
            AddValidations(((IWorkbook) workbook).GetSheet(sheetName), version, validations);
        }
        
        private Cell CreateCell(ICell cell)
        {
            string str = null;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    str = cell.NumericCellValue.ToString(this.Format);
                    break;
                
                case CellType.String:
                    str = cell.StringCellValue;
                    break;
                
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Numeric:
                            double num1;
                            if (cell.CellFormula != "TODAY()")
                            {
                                num1 = cell.NumericCellValue;
                            }
                            else
                            {
                                num1 = DateTime.Today.ToOADate();
                            }
                            str = num1.ToString(this.Format);
                            break;
                        
                        case CellType.String:
                            str = cell.StringCellValue;
                            break;
                        
                        default:
                            break;
                    }
                    break;
                
                case CellType.Boolean:
                    str = cell.BooleanCellValue.ToString();
                    break;
                
                default:
                    break;
            }
            return new Cell(cell.RowIndex, cell.ColumnIndex, str, "", "");
        }
        
        private Row CreateRow(IRow row, int columns)
        {
            if (row == null)
            {
                return null;
            }
            List<Cell> cells = new List<Cell>();
            ICell[] cellArray = row.Cells.ToArray();
            int num = 0;
            for (int i = 0; i < columns; i++)
            {
                Cell item = null;
                if (((i - num) >= cellArray.Length) || (i != cellArray[i - num].ColumnIndex))
                {
                    num++;
                }
                else
                {
                    item = this.CreateCell(cellArray[i - num]);
                }
                cells.Add(item);
            }
            return new Row(row.RowNum, cells);
        }
        
        private Workbook CreateWorkbook(IWorkbook iWorkbook)
        {
            List<Worksheet> worksheets = new List<Worksheet>();
            Workbook workbook = new Workbook(worksheets);
            AddToNames(workbook.Names, iWorkbook);
            string[] names = (from name in workbook.Names select name.Name).ToArray<string>();
            for (int i = 0; i < iWorkbook.NumberOfSheets; i++)
            {
                ISheet sheetAt = iWorkbook.GetSheetAt(i);
                Worksheet item = this.CreateWorksheet(sheetAt, i);
                item.IsHidden = iWorkbook.IsSheetHidden(i);
                AddToValidations(item.Validations, sheetAt, names);
                worksheets.Add(item);
            }
            return workbook;
        }
        
        private static IWorkbook CreateWorkbook(Workbook workbook, ExcelVersion version)
        {
            IWorkbook workbook2;
            switch (version)
            {
                case ExcelVersion.Xls:
                    workbook2 = new HSSFWorkbook();
                    break;
                
                case ExcelVersion.Xlsx:
                    workbook2 = new XSSFWorkbook();
                    break;
                
                default:
                    throw new InvalidEnumArgumentException("version", (int) version, version.GetType());
            }
            AddNames(workbook2, version, workbook.Names.ToArray<NamedRange>());
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                ISheet sheet = workbook2.CreateSheet(worksheet.Name);
                AddValidations(sheet, version, worksheet.Validations.ToArray<DataValidation>());
                AddRows(sheet, worksheet.Rows.ToArray<Row>());
                if (worksheet.IsHidden)
                {
                    workbook2.SetSheetHidden(worksheet.Index, SheetState.Hidden);
                }
            }
            return workbook2;
        }
        
        private Worksheet CreateWorksheet(ISheet sheet, int index)
        {
            List<IRow> list = new List<IRow>();
            int maxColumns = 0;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow item = sheet.GetRow(i);
                if (item != null)
                {
                    maxColumns = Math.Max(maxColumns, item.LastCellNum);
                }
                list.Add(item);
            }
            return new Worksheet(sheet.SheetName, index, maxColumns, (from row in list select this.CreateRow(row, maxColumns)).ToArray<Row>());
        }
        
        public string GetAuthor(Stream stream, ExcelVersion excelVersion)
        {
            return this.GetProperty(stream, excelVersion, info => info.Author, properties => properties.Creator);
        }
        
        public string GetComments(Stream stream, ExcelVersion excelVersion)
        {
            return this.GetProperty(stream, excelVersion, info => info.Comments, properties => properties.Description);
        }
        
        private string GetProperty(Stream stream, ExcelVersion excelVersion, Func<SummaryInformation, string> hssfWorkbookAction, Func<CoreProperties, string> corePropertiesActions)
        {
            using (MemoryStream stream2 = new MemoryStream())
            {
                stream.CopyTo(stream2);
                stream.Position = 0L;
                stream2.Position = 0L;
                if (!excelVersion.Equals(ExcelVersion.Xls))
                {
                    if (excelVersion.Equals(ExcelVersion.Xlsx))
                    {
                        CoreProperties coreProperties = new XSSFWorkbook(stream2).GetProperties().CoreProperties;
                        return corePropertiesActions(coreProperties);
                    }
                }
                else
                {
                    SummaryInformation summaryInformation = new HSSFWorkbook(stream2).SummaryInformation;
                    return hssfWorkbookAction(summaryInformation);
                }
            }
            return string.Empty;
        }
        
        public object GetWorkbook(Stream stream)
        {
            return WorkbookFactory.Create(stream);
        }
        
        public object GetWorkbook(string path)
        {
            return WorkbookFactory.Create(path);
        }
        
        public Workbook ReadWorkbook(Stream stream)
        {
            return this.CreateWorkbook(WorkbookFactory.Create(stream));
        }
        
        public Workbook ReadWorkbook(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                return this.ReadWorkbook(stream);
            }
        }
        
        public void SaveWorkbook(object workbook, Stream stream)
        {
            ((IWorkbook) workbook).Write(stream);
        }
        
        public void SaveWorkbook(object workbook, string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            {
                this.SaveWorkbook(workbook, stream);
            }
        }
        
        public void WriteWorkbook(Workbook workbook, ExcelVersion version, Stream stream)
        {
            CreateWorkbook(workbook, version).Write(stream);
        }
        
        public void WriteWorkbook(Workbook workbook, ExcelVersion version, string path)
        {
            using (Stream stream = File.Create(path))
            {
                this.WriteWorkbook(workbook, version, stream);
            }
        }
        
        public IFormatProvider Format { get; set; }
    }
}
