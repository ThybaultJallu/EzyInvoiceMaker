using ClosedXML.Excel;

class Program
{
    static void Main(string[] args)
    {
        string sourceFolderPath; // Dossier contenant le fichier source
        string destinationFolderPath = "ExcelsExports/"; // Dossier de destination

        
#if DEBUG
        string sourceFilePath = "C:\\Users\\Thybault JALLU\\Documents\\CsharpTests\\EzyInvoiceMaker\\FactureGlobal\\Classeur57.xlsx";
        Console.WriteLine($"Mode Debug : Chemin exact du fichier source : {sourceFilePath}");
#else
        sourceFolderPath = "ExcelSource/"; // Dossier contenant le fichier source

        
        Console.WriteLine("Début du programme...");
        Console.WriteLine($"Dossier source : {sourceFolderPath}");
        Console.WriteLine($"Dossier de destination : {destinationFolderPath}");

        // Récupère tous les fichiers Excel dans le dossier source
        var excelFiles = Directory.GetFiles(sourceFolderPath, "*.xlsx");

        if (excelFiles.Length == 0)
        {
            Console.WriteLine("Aucun fichier Excel trouvé dans le dossier source.");
            return;
        }
        else if (excelFiles.Length > 1)
        {
            Console.WriteLine("Trop de fichiers Excel dans le dossier source. Veuillez n'en laisser qu'un.");
            return;
        }
        string sourceFilePath = excelFiles[0];
#endif
        
        
        Console.WriteLine($"Fichier source sélectionné : {sourceFilePath}");

        using (var sourceWorkbook = new XLWorkbook(sourceFilePath))
        {
            Console.WriteLine("Fichier source chargé avec succès.");

            // Récupère la feuille "CALCUL JUSTIF VENTE" uniquement car le fichier sera toujours fait de plusieurs feuilles
            var sourceWorksheet = sourceWorkbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, "CALCUL JUSTIF VENTE", StringComparison.OrdinalIgnoreCase));
            if (sourceWorksheet == null)
            {
                Console.WriteLine("La feuille 'CALCUL JUSTIF VENTE' est introuvable.");
                return;
            }

            Console.WriteLine("Traitement de la feuille : CALCUL JUSTIF VENTE");

            // Récupère toutes les valeurs uniques de la colonne ou on a les trigrammes
            var uniqueValues = sourceWorksheet
                .Column(6) // Colonne F
                .CellsUsed()
                .Skip(1) // Ignore l'en-tête
                .Select(cell => cell.Value.ToString())
                .Distinct()
                .ToList();

            Console.WriteLine($"Valeurs uniques trouvées dans la colonne F : {string.Join(", ", uniqueValues)}");

            // Crée un fichier Excel pour chaque tri unique
            foreach (var value in uniqueValues)
            {
                Console.WriteLine($"Création du fichier pour la valeur : {value}");

                using (var destinationWorkbook = new XLWorkbook())
                {
                    var destinationWorksheet = destinationWorkbook.Worksheets.Add("Résultats");

                    // Un peu de style parce que c'est jolie et que ça fait "pro"
                    var headerRow = sourceWorksheet.Row(1);
                    for (int col = 1; col <= headerRow.LastCellUsed().Address.ColumnNumber; col++)
                    {
                        var cell = destinationWorksheet.Cell(1, col);
                        cell.Value = headerRow.Cell(col).Value;

                        // Style de l'en-tête
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#277aa1");
                        cell.Style.Font.FontColor = XLColor.White;
                        cell.Style.Font.Bold = true;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Ajout de bordures
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    // Filtre les lignes pour chaque trigramme 
                    var rows = sourceWorksheet.RowsUsed()
                        .Where(row => row.RowNumber() > 1 && row.Cell(6).Value.ToString() == value);

                    var lastRow = rows.Last();

                    int destinationRow = 2; //  l'en-tête
                    foreach (var row in rows)
                    {
                        for (int col = 1; col <= row.LastCellUsed().Address.ColumnNumber; col++)
                        {
                            var cell = destinationWorksheet.Cell(destinationRow, col);
                            cell.Value = row.Cell(col).Value;

                            // Style des lignes alternées pour faire plus lisible 
                            cell.Style.Fill.BackgroundColor = destinationRow % 2 == 0
                                ? XLColor.FromHtml("#ffffff")
                                : XLColor.FromHtml("#b8b8b8");
                            cell.Style.Font.FontColor = XLColor.Black;

                            // Centrage des valeurs parce qu'on le peut
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                            // Ajout de bordures pour la lisibilite
                            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            cell.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                            // Ajout de la bordure inférieure pour la dernière ligne warn perf faible essayer d'init la lastrow ca devrais regler le soucis ??
                            if (row == lastRow)
                            {
                                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            }
                        }
                        destinationRow++;
                    }

                    // Ajuste la largeur des colonnes
                    destinationWorksheet.Columns().AdjustToContents();
                    
                    string destinationFilePath;
                    string invoiceMonth = rows.FirstOrDefault().Cell(4).Value.ToString();
                    string invoiceyear = rows.FirstOrDefault().Cell(3).Value.ToString();
                    if (rows.FirstOrDefault().Cell(1).Value.ToString() == "OK")
                    {
                        destinationFilePath = Path.Combine(destinationFolderPath, $"Facture_Ezytail_{value}_{invoiceMonth}_{invoiceyear}.xlsx");
                        destinationWorkbook.SaveAs(destinationFilePath);
                    }
                    else
                    {
                        destinationFilePath = Path.Combine(destinationFolderPath, $"ResumeKO_{value}_{invoiceMonth}_{invoiceyear}.xlsx");
                        destinationWorkbook.SaveAs(destinationFilePath);
                    }
                    Console.WriteLine($"Fichier créé : {destinationFilePath}");
                }
            }
        }

        Console.WriteLine("Fin du programme.");
#if DEBUG
#else
        Console.WriteLine("Appuyez sur une touche pour quitter.");
        Console.ReadKey();
#endif
    }
}