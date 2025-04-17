using System.IO;

namespace EzyInvoiceMaker;

public class FileManager
{
    public string GetExcelSource(string folderPath)
    {
        var files = Directory.GetFiles(folderPath, "Socle_Ventilation.xlsx");
        if (files.Length == 0)
            throw new Exception(
                "Aucun fichier Excel trouvé. Verifiez le nom du fichier source ce la doit etre \"Socle_Ventilation.xlsx\".");
        if (files.Length > 1) throw new Exception("Trop de fichiers Excel dans le dossier source.");
        return files[0];
    }



    public void EnsureDirectoryExists(string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }
    }
}