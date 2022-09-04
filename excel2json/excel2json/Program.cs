using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace excel2json
{
    class Program
    {
        static void Main(string[] args)
        {
            Start();
        }

        static void Start()
        {
            Console.WriteLine("please write or drag excel file into the window");
            string[] files = WatingWriteSourcePath();
            foreach (var item in files)
            {
                Export(item);
            }
            Start();
        }

        static void Export(string file)
        {

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(file, false))
            {
                Dictionary<string, string> map = new Dictionary<string, string>();
                foreach (var item in spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>())
                {
                    map.Add(item.Name, item.Id);
                }

                if (!map.TryGetValue("字段说明", out string id))
                {
                    Console.WriteLine("error:not find table :字段说明");
                    return;
                }

                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(id);
                Dictionary<string, TableInfo> fields = GetTableInfo(worksheetPart, spreadsheetDocument.WorkbookPart.SharedStringTablePart);
                Dictionary<string, DataTable> tables = new Dictionary<string, DataTable>();
                foreach (KeyValuePair<string, string> item in map)
                {
                    if (item.Key == "字段说明" || !fields.TryGetValue(item.Key, out TableInfo table))
                    {
                        continue;
                    }
                    worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(item.Value);
                    DataTable dataTable = GetDataTable(worksheetPart, spreadsheetDocument.WorkbookPart.SharedStringTablePart, table);
                    tables.Add(table.name, dataTable);
                }

                //TODO export to lua script code
                foreach (var item in fields.Values)
                {
                    StringBuilder builder = new StringBuilder();
                    builder.AppendLine($"---@class {item.name} {item.des}");
                    builder.AppendLine(String.Format("local {0} = {{}}", item.name));
                    builder.AppendLine(String.Format("local this = {0}", item.name));
                    foreach (var field in item.fields.Values)
                    {
                        if (field.type.IsArray)
                        {
                            builder.AppendLine("---@type table " + field.des);
                        }
                        else
                        if (field.type == typeof(string))
                        {
                            builder.AppendLine("---@type string " + field.des);
                        }
                        else
                        if (field.type == typeof(bool))
                        {
                            builder.AppendLine("---@type boolean " + field.des);
                        }
                        else
                        {
                            builder.AppendLine("---@type number " + field.des);
                        }
                        builder.AppendLine(String.Format("this.{0} = nil", field.name));
                    }
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "script");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    File.WriteAllText(Path.Combine(path, item.name + ".lua"), builder.ToString());
                }

                //TODO ecport to json file
                foreach (var item in tables)
                {
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(item.Value);
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "data");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    File.WriteAllText(Path.Combine(path, item.Key + ".ini"), json);
                }

                Console.WriteLine("export excel:" + file + " completion");
            }
        }
        class FieldInfo
        {
            public string name;
            public Type type;
            public string des;
        }

        class TableInfo
        {
            public string name;
            public string des;
            public Dictionary<string, FieldInfo> fields = new Dictionary<string, FieldInfo>();
        }
        static Dictionary<string, TableInfo> GetTableInfo(WorksheetPart worksheetPart, SharedStringTablePart sharedStringTablePart)
        {
            Dictionary<string, TableInfo> tables = new Dictionary<string, TableInfo>();
            SheetData _Sheet1data = worksheetPart.Worksheet.Elements<SheetData>().First();
            TableInfo current = null;
            foreach (Row row in _Sheet1data.Elements<Row>())
            {
                List<Cell> cells = row.Elements<Cell>().ToList();
                Cell firstCell = cells.FirstOrDefault();
                if (cells.Where(x => string.IsNullOrEmpty(GetCellData(x, sharedStringTablePart))).Count() <= 0)
                {
                    continue;
                }
                if (firstCell != null && GetCellData(firstCell, sharedStringTablePart) == "表名")
                {
                    current = new TableInfo();
                    cells = row.Elements<Cell>().ToList();
                    current.name = GetCellData(cells[1], sharedStringTablePart);
                    current.des = GetCellData(cells[2], sharedStringTablePart);
                    tables.Add(current.name, current);
                    continue;
                }
                if (current == null)
                {
                    continue;
                }
                string fieldName = GetCellData(cells[0], sharedStringTablePart);
                if (string.IsNullOrEmpty(fieldName))
                {
                    continue;
                }
                FieldInfo info = new FieldInfo();
                info.name = fieldName;
                info.type = Type.GetType(GetCellData(cells[1], sharedStringTablePart));
                info.des = GetCellData(cells[2], sharedStringTablePart);
                current.fields.Add(info.name, info);
            }
            return tables;
        }

        static string GetCellData(Cell cell, SharedStringTablePart sharedStringTablePart)
        {
            if (cell == null || cell.CellValue == null)
            {
                return string.Empty;
            }

            string text = cell.CellValue.Text;
            //判断是不是在SharedStringTable中
            if (cell.DataType != null)
            {
                var _xmlpart = sharedStringTablePart.SharedStringTable.ElementAt(Convert.ToInt32(cell.CellValue.Text));
                text = _xmlpart.FirstChild.InnerText;
            }
            return text;
        }

        static DataTable GetDataTable(WorksheetPart worksheetPart, SharedStringTablePart sharedStringTablePart, TableInfo fiedles)
        {
            if (fiedles == null || fiedles.fields == null || fiedles.fields.Count <= 0)
            {
                return null;
            }
            DataTable dttable = new DataTable(fiedles.name);
            foreach (var item in fiedles.fields.Values)
            {
                dttable.Columns.Add(new DataColumn(item.name, item.type));
            }

            //sheet页中的内容
            SheetData _Sheet1data = worksheetPart.Worksheet.Elements<SheetData>().First();
            List<Row> rows = _Sheet1data.Elements<Row>().ToList();
            for (int i = 3; i < rows.Count; i++)
            {
                List<Cell> cells = rows[i].Elements<Cell>().ToList();
                List<object> datas = new List<object>();
                for (int j = 0; j < cells.Count; j++)
                {
                    string data = GetCellData(cells[j], sharedStringTablePart);
                    if (data.Contains(","))
                    {
                        string[] strings = data.Split(',');
                        Type arguments = GetArrayElementType(dttable.Columns[j].DataType);
                        ArrayList arrayList = new ArrayList();
                        foreach (var item in strings)
                        {
                            arrayList.Add(Convert.ChangeType(item, arguments));
                        }
                        datas.Add(arrayList.ToArray(arguments));
                    }
                    else
                    {
                        datas.Add(Convert.ChangeType(data, dttable.Columns[j].DataType));
                    }
                }
                dttable.Rows.Add(datas.ToArray());
            }
            return dttable;
        }

        public static Type GetArrayElementType(Type t)
        {
            if (!t.IsArray) return null;

            string tName = t.FullName.Replace("[]", string.Empty);

            Type elType = t.Assembly.GetType(tName);

            return elType;
        }

        static string[] WatingWriteSourcePath()
        {
            string path = Console.ReadLine();
            if (string.IsNullOrEmpty(path))
            {
                return WatingWriteSourcePath();
            }
            List<string> strings = new List<string>();
            string[] parts = path.Split('"');
            foreach (var item in parts)
            {
                if (string.IsNullOrEmpty(item))
                {
                    continue;
                }
                if (string.IsNullOrEmpty(Path.GetExtension(item)))
                {
                    strings.AddRange(GetFiles(item));
                    continue;
                }
                strings.Add(item);
            }
            return strings.ToArray();
        }

        static string[] GetFiles(string path)
        {
            string[] files = Directory.GetFiles(path, "*.xlsm", SearchOption.AllDirectories);
            return files;
        }
    }
}