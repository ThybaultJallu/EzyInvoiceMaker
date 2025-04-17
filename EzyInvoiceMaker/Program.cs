using ClosedXML.Excel;

namespace EzyInvoiceMaker;
class Program
{
    static void Main(string[] args)
    {
        string sourceFolderPath = "ExcelSource/";
        string destinationFolderPath = "ExcelsExports/";
        
    #if DEBUG
        sourceFolderPath = "C:\\Users\\Thybault JALLU\\Documents\\Github\\EzyInvoiceMaker\\FactureGlobal\\ExcelSource";
        destinationFolderPath = "C:\\Users\\Thybault JALLU\\Documents\\Github\\EzyInvoiceMaker\\FactureGlobal\\ExcelsExports";
    #endif

        var fileManager = new FileManager();
        var excelProcessor = new ExcelProcessor();

        try
        {
            fileManager.EnsureDirectoryExists(destinationFolderPath);
            string sourceFilePath = fileManager.GetExcelSource(sourceFolderPath);
            Console.WriteLine($"Fichier source : {sourceFilePath}");
            
            // Saisie des données
            Console.WriteLine("Pour quelle annee souhaitez vous faire le traitement ?");
            //string invoiceYear = Console.ReadLine();
            string invoiceYear = "2025";
            Console.WriteLine("Pour quel mois souhaitez vous faire le traitement ?");
            //string invoiceMonth = Console.ReadLine();
            string invoiceMonth = "3";
            
            // Traitement en une seule lecture du fichier
            using (var excelWorkbook = new XLWorkbook(sourceFilePath))
            {
                var sourceWorksheet = excelWorkbook.Worksheet("CALCUL JUSTIF VENTE");
                var headerRow = sourceWorksheet.Row(1);
                var dataRows = sourceWorksheet.RowsUsed().Skip(1).ToList(); // Toutes les lignes sauf l'en-tête
                
                // Regrouper les lignes par trigramme en une seule passe comme 
                var groupedRows = dataRows
                    .GroupBy(row => row.Cell(6).Value.ToString())
                    .ToDictionary(g => g.Key, g => g.ToList());
                    
                // Ajouter l'en-tête à chaque groupe
                foreach (var group in groupedRows)
                {
                    var trigram = group.Key;
                    var rowsWithHeader = new List<IXLRow> { headerRow };
                    rowsWithHeader.AddRange(group.Value);
                    
                    Console.WriteLine($"Traitement pour la valeur : {trigram}");
                    excelProcessor.CreateExcelFile(destinationFolderPath, rowsWithHeader, invoiceMonth, invoiceYear, trigram);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erreur : {ex.Message}");
        }
    }
}