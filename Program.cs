using System.Xml.Linq;
using ClosedXML.Excel;
using System.Globalization; // <--- Importante para manejar los puntos decimales

string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
string outputDirectory = Path.Combine(baseDirectory, "out");

if (!Directory.Exists(outputDirectory)) Directory.CreateDirectory(outputDirectory);

XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";
string[] xmlFiles = Directory.GetFiles(baseDirectory, "*.xml");

foreach (var xmlPath in xmlFiles)
{
    string fileName = Path.GetFileNameWithoutExtension(xmlPath);
    string outputPath = Path.Combine(outputDirectory, $"{fileName}.xlsx");

    try
    {
        XDocument doc = XDocument.Load(xmlPath);
        using (var workbook = new XLWorkbook())
        {
            var worksheets = doc.Descendants(ss + "Worksheet");

            foreach (var wsNode in worksheets)
            {
                string sheetName = wsNode.Attribute(ss + "Name")?.Value ?? "Sheet";
                var worksheet = workbook.Worksheets.Add(sheetName);

                var rows = wsNode.Descendants(ss + "Row");
                int currentRow = 1;

                foreach (var rowNode in rows)
                {
                    var cells = rowNode.Descendants(ss + "Cell");
                    int currentCol = 1;

                    foreach (var cellNode in cells)
                    {
                        var dataNode = cellNode.Element(ss + "Data");
                        if (dataNode != null)
                        {
                            string value = dataNode.Value;
                            string type = dataNode.Attribute(ss + "Type")?.Value ?? "String";

                            var currentCell = worksheet.Cell(currentRow, currentCol);

                            // TRATAMIENTO DE NÚMEROS CON PUNTO DECIMAL
                            if (type == "Number")
                            {
                                // Usamos InvariantCulture para que entienda que el "." es decimal
                                if (double.TryParse(value, CultureInfo.InvariantCulture, out double num))
                                {
                                    currentCell.Value = num;
                                    // Opcional: Forzar a que Excel muestre 2 decimales
                                    currentCell.Style.NumberFormat.Format = "#,##0.00";
                                }
                            }
                            else if (type == "DateTime" && DateTime.TryParse(value, out DateTime dt))
                            {
                                currentCell.Value = dt;
                                currentCell.Style.DateFormat.Format = "dd/mm/yyyy";
                            }
                            else
                            {
                                currentCell.Value = value;
                            }
                        }
                        currentCol++;
                    }
                    currentRow++;
                }
                worksheet.Columns().AdjustToContents();
            }
            workbook.SaveAs(outputPath);
            Console.WriteLine($"[OK] {fileName}.xlsx generado con éxito.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[ERROR] {fileName}: {ex.Message}");
    }
}