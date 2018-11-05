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
            var dt = ReadFile(stream, name);
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var dtRow = dt.Rows[i];
                if (!dtRow.ItemArray.All(x => string.IsNullOrWhiteSpace(x.ToString())))
                {
                    yield return GetEntity<T>(dtRow);
                }
            }
        }

        private DataTable ReadFile(Stream stream, string name)
        {
            var dt = new DataTable();
            using (var doc = SpreadsheetDocument.Open(stream, false))
            {
                var workbookPart = doc.WorkbookPart;
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(w => w.Name == name);
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet?.Id);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var rows = sheetData.Elements<Row>().ToArray();

                var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                var sst = sstpart.SharedStringTable;

                var headers = rows[0].Elements<Cell>().ToArray();
                for (var i = 0; i < headers.Length; i++)
                {
                    var cell = headers[i];
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                    {
                        var cellIndex = Convert.ToInt32(cell.CellValue.Text);
                        var value = sst.ChildElements[cellIndex].InnerText.Replace('-', ' ').Replace(" ", string.Empty);
                        dt.Columns.Add(value);
                    }
                }

                for (var i = 1; i < rows.Length; i++)
                {
                    var dtRow = dt.NewRow();
                    var cell = rows[i].Elements<Cell>().ToArray();
                    for (var j = 0; j < cell.Length; j++)
                    {
                        var c = cell[j];
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var cellIndex = int.Parse(c.CellValue.Text);
                            var value = sst.ChildElements[cellIndex].InnerText;
                            dtRow[j] = value;
                        }
                        else if (c.CellValue != null)
                        {
                            dtRow[j] = c.CellValue.Text;
                        }
                    }
                    dt.Rows.Add(dtRow);
                }
            }

            return dt;
        }

        private T GetEntity<T>(DataRow row) where T : new()
        {
            var entity = new T();
            var properties = typeof(T).GetProperties();

            for (var index = 0; index < properties.Length; index++)
            {
                var property = properties[index];
                var propertyName = property.Name;

                object v;
                try
                {
                    v = row[propertyName];
                }
                catch (Exception)
                {
                    continue;
                }

                try
                {
                    var valor = v.GetType() == property.PropertyType
                        ? v
                        : Convert.ChangeType(row[propertyName], property.PropertyType);

                    property.SetValue(entity, valor);
                }
                catch (Exception)
                {
                    throw new InvalidCastException();
                }
            }

            return entity;
        }
    }
}
