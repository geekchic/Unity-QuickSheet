﻿///////////////////////////////////////////////////////////////////////////////
///
/// ExcelQuery.cs
///
/// (c)2014 Kim, Hyoun Woo
///
///////////////////////////////////////////////////////////////////////////////
using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System;
using System.Linq;
using System.ComponentModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UnityQuickSheet
{
    /// <summary>
    /// Query each of cell data from the given excel sheet and deserialize it to the ScriptableObject's data array.
    /// </summary>
    public class ExcelQuery
    {
        private readonly IWorkbook workbook = null;
        //private readonly ISheet sheet = null;
        private readonly XSSFTable table = null;
        private string filepath = string.Empty;

        /// <summary>
        /// Constructor.
        /// </summary>
        public ExcelQuery(string path, string sheetName = "")
        {
            try
            {
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    string extension = GetSuffix(path);

                    if (extension == "xls")
                        workbook = new HSSFWorkbook(fileStream);
                    else if (extension == "xlsx")
                    {
                    #if UNITY_EDITOR_OSX
                        throw new Exception("xlsx is not supported on OSX.");
                    #else
                        workbook = new XSSFWorkbook(fileStream);
                    #endif
                    }
                    else
                    {
                        throw new Exception("Wrong file.");
                    }


                    int numSheets = this.workbook.NumberOfSheets;
                    for (int i = 0; i < numSheets; i++)
                    {
                        var sheet = (XSSFSheet)this.workbook.GetSheetAt(i);
                        var found = sheet.GetTables().Find(x => x.Name == sheetName);
                        if (found != null)
                        {
                            table = found;
                        }
                    }

                    this.filepath = path;
                }
            }
            catch (Exception e)
            {
                Debug.LogError(e.Message);
            }
        }

        /// <summary>
        /// Determine whether the excel file is successfully read in or not.
        /// </summary>
        public bool IsValid()
        {
            if (this.workbook != null && this.table != null)
                return true;

            return false;
        }

        /// <summary>
        /// Retrieves file extension only from the given file path.
        /// </summary>
        static string GetSuffix(string path)
        {
            string ext = Path.GetExtension(path);
            string[] arg = ext.Split(new char[] { '.' });
            return arg[1];
        }

        string GetHeaderColumnName(int cellnum)
        {
            if (table.GetCTTable().tableColumns.count >= cellnum)
            {
                return string.Empty;
            }

            var column = table.GetCTTable().tableColumns.tableColumn[cellnum];
            return column.name;
        }

        /// <summary>
        /// Deserialize all the cell of the given sheet.
        ///
        /// NOTE:
        ///     The first row of a sheet is header column which is not the actual value
        ///     so it skips when it deserializes.
        /// </summary>
        public List<T> Deserialize<T>(int start = 1)
        {
            var t = typeof(T);
            PropertyInfo[] p = t.GetProperties();

            var result = new List<T>();

            var sheet = table.GetXSSFSheet();
            int startRow = table.GetStartCellReference().Row + 1;
            int endRow = table.GetEndCellReference().Row;
            int startCol = table.GetStartCellReference().Col;
            int endCol = table.GetEndCellReference().Col;

            for (int r = startRow; r <= endRow; r++)
            {
                IRow row = sheet.GetRow(r);

                var item = (T)Activator.CreateInstance(t);

                for (int c = startCol; c <= endCol; c++)
                {
                    var property = p[c - startCol];

                    ICell cell = row.GetCell(c);

                    try
                    {
                        var value = ConvertFrom(cell, property.PropertyType);
                        property.SetValue(item, value, null);
                    }
                    catch (Exception e)
                    {
                        string pos = string.Format("Row[{0}], Cell[{1}]", r - startRow, c - startCol);
                        Debug.LogError(string.Format("Excel File {0} Deserialize Exception: {1} at {2}", this.filepath, e.Message, pos));
                    }
                }

                result.Add(item);
            }

            return result;
        }

        /// <summary>
        /// Retrieves all sheet names.
        /// </summary>
        public string[] GetSheetNames()
        {
            List<string> sheetList = new List<string>();
            if (this.workbook != null)
            {
                int numSheets = this.workbook.NumberOfSheets;
                for (int i = 0; i < numSheets; i++)
                {
                    var sheet = (XSSFSheet)this.workbook.GetSheetAt(i);
                    foreach (var table in sheet.GetTables())
                    {
                        if (table.Name != table.GetCTTable().tableColumns.tableColumn[0].name)
                        {
                            sheetList.Add(table.Name);
                        }
                    }
                }
            }
            else
                Debug.LogError("Workbook is null. Did you forget to import excel file first?");

            return (sheetList.Count > 0) ? sheetList.ToArray() : null;
        }

        /// <summary>
        /// Retrieves all first columns(aka. header column) which are needed to determine each type of a cell.
        /// </summary>
        public string[] GetTitle(ref string error)
        {
            if (table == null)
            {
                error = string.Format(@"Empty row");
                return null;
            }

            List<string> result = new List<string>();

            for (int i = 0; i < table.GetCTTable().tableColumns.count; i++)
            {
                var cell = table.GetCTTable().tableColumns.tableColumn[i];
                if (cell == null)
                {
                    // null or empty column is found. Note column index starts from 0.
                    Debug.LogWarningFormat("Null or empty column is found at {0}.\n", i);
                    continue;
                }
                string value = cell.name;
                if (string.IsNullOrEmpty(value))
                {
                    // null or empty column is found. Note column index starts from 0.
                    Debug.LogWarningFormat("Null or empty column is found at {0}.", i);
                }
                else
                {
                    // column header is not an empty string, we check its validation later.
                    result.Add(value);
                }
            }

            return result.ToArray();
        }

        /// <summary>
        /// Convert type of cell value to its predefined type which is specified in the sheet's ScriptMachine setting file.
        /// </summary>
        protected object ConvertFrom(ICell cell, Type t)
        {
            object value = null;

            if (t == typeof(float) || t == typeof(double) || t == typeof(short) || t == typeof(int) || t == typeof(long))
            {
                if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    value = cell.NumericCellValue;
                }
                else if (cell.CellType == NPOI.SS.UserModel.CellType.String)
                {
                    //Get correct numeric value even the cell is string type but defined with a numeric type in a data class.
                    if (t == typeof(float))
                        value = Convert.ToSingle(cell.StringCellValue);
                    if (t == typeof(double))
                        value = Convert.ToDouble(cell.StringCellValue);
                    if (t == typeof(short))
                        value = Convert.ToInt16(cell.StringCellValue);
                    if (t == typeof(int))
                        value = Convert.ToInt32(cell.StringCellValue);
                    if (t == typeof(long))
                        value = Convert.ToInt64(cell.StringCellValue);
                }
                else if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
                {
                    // Get value even if cell is a formula
                    if (t == typeof(float))
                        value = Convert.ToSingle(cell.NumericCellValue);
                    if (t == typeof(double))
                        value = Convert.ToDouble(cell.NumericCellValue);
                    if (t == typeof(short))
                        value = Convert.ToInt16(cell.NumericCellValue);
                    if (t == typeof(int))
                        value = Convert.ToInt32(cell.NumericCellValue);
                    if (t == typeof(long))
                        value = Convert.ToInt64(cell.NumericCellValue);
                }
            }
            else if (t == typeof(string) || t.IsArray)
            {
                // HACK: handles the case that a cell contains numeric value
                //       but a member field in a data class is defined as string type.
                //       e.g. string s = "123"
                if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                    value = cell.NumericCellValue;
                else
                    value = cell.StringCellValue;
            }
            else if (t == typeof(bool))
                value = cell.BooleanCellValue;

            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                var nc = new NullableConverter(t);
                return nc.ConvertFrom(value);
            }

            if (t.IsEnum)
            {
                // for enum type, first get value by string then convert it to enum.
                value = cell.StringCellValue;
                return Enum.Parse(t, value.ToString(), true);
            }
            else if (t.IsArray)
            {
                if (t.GetElementType() == typeof(float))
                    return ConvertExt.ToSingleArray((string)value);

                if (t.GetElementType() == typeof(double))
                    return ConvertExt.ToDoubleArray((string)value);

                if (t.GetElementType() == typeof(short))
                    return ConvertExt.ToInt16Array((string)value);

                if (t.GetElementType() == typeof(int))
                    return ConvertExt.ToInt32Array((string)value);

                if (t.GetElementType() == typeof(long))
                    return ConvertExt.ToInt64Array((string)value);

                if (t.GetElementType() == typeof(string))
                    return ConvertExt.ToStringArray((string)value);
            }

            // for all other types, convert its corresponding type.
            return Convert.ChangeType(value, t);
        }

        public XSSFTable[] GetEnumTables()
        {
            List<XSSFTable> tableList = new List<XSSFTable>();
            if (this.workbook != null)
            {
                int numSheets = this.workbook.NumberOfSheets;
                for (int i = 0; i < numSheets; i++)
                {
                    var sheet = (XSSFSheet)this.workbook.GetSheetAt(i);
                    foreach (var table in sheet.GetTables())
                    {
                        if (table.Name == table.GetCTTable().tableColumns.tableColumn[0].name)
                        {
                            tableList.Add(table);
                        }
                    }
                }
            }
            else
                Debug.LogError("Workbook is null. Did you forget to import excel file first?");
            return (tableList.Count > 0) ? tableList.ToArray() : null;
        }
    }
}
