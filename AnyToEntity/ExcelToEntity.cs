using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AnyToEntity
{
    public class ExcelToEntity
    {
        public IEnumerable<T> Read<T>(Stream stream, string name) where T : new()
        {
            using (var doc = SpreadsheetDocument.Open(stream, false))
            {
                var workbookPart = doc.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(w => w.Name == name);
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet?.Id);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Elements<Row>().ToList();

                var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                var sst = sstpart.SharedStringTable;

                var properties = typeof(T).GetProperties();

                for (var i = 1; i < rows.Count; i++)
                {
                    var expando = new Dictionary<string, object>();
                    var cell = rows[i].Elements<Cell>().ToList();
                    for (var j = 0; j < cell.Count; j++)
                    {
                        var c = cell[j];
                        var propiedad = properties[j];
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var cellIndex = int.Parse(c.CellValue.Text);
                            var value = sst.ChildElements[cellIndex].InnerText;
                            expando[propiedad.Name] = value;
                        }
                        else if (c.CellValue != null)
                        {
                            expando[propiedad.Name] = c.CellValue.Text;
                        }
                    }
                    yield return GetObject<T>(expando);
                }
            }
        }

        private T GetObject<T>(Dictionary<string, object> dict) where T : new()
        {
            var type = typeof(T);
            var obj = new T();

            foreach (var kv in dict)
            {
                type.GetProperty(kv.Key)?.SetValue(obj, kv.Value);
            }
            return obj;
        }
    }
}
