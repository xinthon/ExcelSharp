using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelSharp
{
    public sealed class ExcelSharp
    {
        public ExcelSharp() { ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; }

        public IEnumerable<T> Import<T>(string filePath) where T : class
        {
            List<T> importedData = new List<T>();

            // Read the Excel file
            using(ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if(worksheet is not ExcelWorksheet excelWorksheet)
                    throw new Exception("Worksheet not set");

                // Get the column headers
                Dictionary<int, string?> columnHeaders = GetColumnHeaders<T>(excelWorksheet);

                // Read the data rows
                int rowCount = 2;
                while(worksheet.Cells[rowCount, 1].Value != null)
                {
                    T model = CreateModel<T>(worksheet, columnHeaders, rowCount);
                    importedData.Add(model);
                    rowCount++;
                }
            }

            return importedData;
        }

        public void Export<Model>(IEnumerable<Model> models, string filePath)
        {
            try
            {
                using(ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Set the column headers based on the Model attributes
                    this.SetColumnHeaders<Model>(worksheet);

                    // Populate the data rows
                    int rowCount = 2;
                    foreach(var model in models)
                    {
                        this.SetDataRow<Model>(worksheet, model, rowCount);
                        rowCount++;
                    }

                    // Save the Excel file
                    FileInfo excelFile = new FileInfo(filePath);
                    package.SaveAs(excelFile);
                }
            } catch(Exception ex)
            {
            }
        }

        private Dictionary<int, string?> GetColumnHeaders<Model>(ExcelWorksheet worksheet)
        {
            var modelProperties = typeof(Model).GetProperties().Select(p => this.GetColumnName(p)).ToArray();

            Dictionary<int, string?> columnHeaders = new Dictionary<int, string?>();

            int columnCount = 1;
            while(worksheet.Cells[1, columnCount].Value != null)
            {
                string? columnName = worksheet.Cells[1, columnCount].Value.ToString();

                if(!modelProperties?.Any(e => e.columnName == columnName) == true)
                {
                    columnCount++;
                    continue;
                }

                columnHeaders.Add(columnCount, columnName);
                columnCount++;
            }
            return columnHeaders;
        }

        private T CreateModel<T>(ExcelWorksheet worksheet, Dictionary<int, string?> columnHeaders, int rowCount)
            where T : class
        {
            T model = Activator.CreateInstance<T>();

            foreach(var property in typeof(T).GetProperties())
            {
                (string columnName, int width) = GetColumnName(property);

                if(columnHeaders.ContainsValue(columnName))
                {
                    int columnIndex = columnHeaders.FirstOrDefault(x => x.Value == columnName).Key;
                    var cellValue = worksheet.Cells[rowCount, columnIndex].Value;

                    if(cellValue != null)
                    {
                        property.SetValue(model, Convert.ChangeType(cellValue, property.PropertyType));
                    }
                }
            }
            return model;
        }

        private (string columnName, int width) GetColumnName(PropertyInfo property)
        {
            var column = (columnName: (string)property.Name, width: (int)40);

            if(property.IsDefined(typeof(DisplayAttribute), true))
            {
                var displayAttribute = property.GetCustomAttributes(typeof(DisplayAttribute), true).FirstOrDefault() as DisplayAttribute;

                column.columnName = displayAttribute.Name;
            }

            if(property.IsDefined(typeof(WidthAttribute), true))
            {
                var widthAttribute = property.GetCustomAttributes(typeof(WidthAttribute), true).FirstOrDefault() as WidthAttribute;

                column.width = widthAttribute.Width;
            }

            return column;
        }

        private void SetColumnHeaders<T>(ExcelWorksheet worksheet)
        {
            int columnCount = 1;
            var properties = typeof(T).GetProperties();

            foreach(var property in properties)
            {
                var (columnName, width) = GetColumnName(property);
                worksheet.Cells[1, columnCount].Value = columnName;
                worksheet.Columns[1, columnCount].Width = width;
                columnCount++;
            }
        }

        private void SetDataRow<Model>(ExcelWorksheet worksheet, Model model, int rowCount)
        {
            int columnCount = 1;
            foreach(var property in typeof(Model).GetProperties())
            {
                var value = property.GetValue(model);
                worksheet.Cells[rowCount, columnCount].Value = value;
                columnCount++;
            }
        }
    }
}
