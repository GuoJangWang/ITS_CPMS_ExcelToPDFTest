// See https://aka.ms/new-console-template for more information


//https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Excel-Conversion/NET-Excel-New-method-of-Convert-Excel-to-PDF.html

//https://www.e-iceblue.com/Buy/Spire.PDF.html



using Spire.Xls;

Console.WriteLine("開始轉檔");

FileProcess.FileTransProcess();

Console.ReadKey();


public class Const
{
    public static string XlsxFileFullPath
    {
        get
        {
            return Path.Combine(Directory.GetCurrentDirectory(),"Files", "TableSchema.xlsx");
        }
    }

    public static string PDFName
    {
        get
        {
            return "PDFResult.pdf";
        }
    }

}


public class FileProcess
{
    public static bool ReadFileToStream()
    {
        try
        {

        }
        catch (Exception)
        {

            throw;
        }
        return true;
    }

    public static bool FileTransProcess()
    {
        bool result = true;
        try
        {
            Workbook workbook = new Workbook();

            workbook.LoadFromFile(Const.XlsxFileFullPath);

            workbook.ConverterSetting.SheetFitToPage = true;

            workbook.SaveToFile(Const.PDFName,FileFormat.PDF);
        }
        catch (Exception ex)
        {
            
            Console.WriteLine(ex); 
            result = false;
        }
        return result;
    }

    public static bool WriteFile()
    {
        try
        {

        }
        catch (Exception)
        {

            throw;
        }
        return false;
    }

}