using ClosedXML.Excel;

namespace EzyInvoiceMaker;

public class ExcelProcessor
{
    public List<string> GetUniqueValues(string filePath, string sheetName, int columnIndex)
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet == null) throw new Exception($"La feuille '{sheetName}' est introuvable.");

            // Vérifier si l'index de colonne est correct
            Console.WriteLine($"Recherche des valeurs uniques dans la colonne {columnIndex}");

            // Récupérer les valeurs affichées plutôt que les valeurs brutes
            var uniqueValues = worksheet.RowsUsed()
                .Skip(1) // Ignorer l'en-tête
                .Select(row =>
                {
                    var cell = row.Cell(columnIndex);
                    var displayValue = cell.GetFormattedString();
                    return displayValue;
                })
                .Distinct()
                .ToList();

            Console.WriteLine($"Valeurs uniques trouvées : {string.Join(", ", uniqueValues)}");
            return uniqueValues;
        }
    }

    public void CreateExcelFile(string destinationPath, IEnumerable<IXLRow> rows, string invoiceMonth, string invoiceYear, string trigram = "TRIGRAMMEINTROUVABLE")
    {
        Console.WriteLine($"Création du fichier pour la valeur : {trigram}");
        using (var destinationWorkbook = new XLWorkbook())
        {
            var destinationWorksheet = destinationWorkbook.Worksheets.Add($"Facture_{invoiceMonth}_{invoiceYear}");

            var rowsList = rows.ToList();
            IXLRow headerRow = rowsList.FirstOrDefault();
            IXLRow lastRow = rowsList.LastOrDefault();

            foreach (var row in rowsList)
            {
                int currentRowIdx = row == headerRow ? 1 : destinationWorksheet.LastRowUsed().RowNumber() + 1;
                int columnIdx = 1;

                foreach (var cell in row.Cells())
                {
                    destinationWorksheet.Cell(currentRowIdx, columnIdx).Value = cell.Value;
                    
                    
                    if (columnIdx >= 26 && columnIdx <= 31 && cell.Value.IsNumber)
                    {
                        // Format monétaire euro pour els colonnes 26 à 31 qui sont les montants a payer
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.NumberFormat.Format = "#,##0.00 €";
                    }
                    // On bloque a deux chiffres après la virgule pour le reste
                    else if (cell.Value.IsNumber)
                    {
                        double numValue = cell.Value.GetNumber();
                        if (Math.Abs(numValue - Math.Round(numValue)) < double.Epsilon)
                        {
                            // nb entier
                            destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.NumberFormat.Format = "0";
                        }
                        else
                        {
                            // nb reel
                            destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.NumberFormat.Format = "0.00";
                        }
                    }

                    // Style pour l'en-tête
                    if (row == headerRow)
                    {
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Fill.BackgroundColor = XLColor.FromHtml("#277aa1");
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Font.FontColor = XLColor.White;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Font.Bold = true;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    }
                    else // style du reste 
                    {
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                        if (currentRowIdx % 2 == 0)
                            destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Fill.BackgroundColor = XLColor.FromHtml("#ffffff");
                        else
                            destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Fill.BackgroundColor = XLColor.FromHtml("#b8b8b8");

                        if (row == lastRow)
                            destinationWorksheet.Cell(currentRowIdx, columnIdx).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    }
                    columnIdx++;
                }
            }

            destinationWorksheet.Columns().AdjustToContents();

            string destinationFilePath;
            if (rowsList[1].Cell(1).Value.ToString() == "OK")
                destinationFilePath = Path.Combine(destinationPath, $"Facture_Ezytail_{trigram}_{invoiceMonth}_{invoiceYear}.xlsx");
            else
                destinationFilePath = Path.Combine(destinationPath, $"ResumeKO_{trigram}_{invoiceMonth}_{invoiceYear}.xlsx");

            destinationWorkbook.SaveAs(destinationFilePath);
            Console.WriteLine($"Fichier créé : {destinationFilePath}");
        }
    }
}