// See https://aka.ms/new-console-template for more information

using System.Text;
using iTextSharp.text.pdf;
using OfficeOpenXml;

//非商業使用，如果沒加這行會被擋下來
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//取得excel路徑
var excelFilePath = "D:\\self-Practice\\BackEnd\\EncryptSample\\EncryptSample\\ExcelContainer\\NameAndPassword.xlsx";

//取得pdf路徑
var pdfFolderPath = "D:\\self-Practice\\BackEnd\\EncryptSample\\EncryptSample\\FileContainer";

var package = new ExcelPackage(new FileInfo(excelFilePath));

//假設Excel文件的第一個工作表包含姓名和密碼
var worksheet = package.Workbook.Worksheets[0];

//從第二行開始，(第一行是標題
for (var row = 2; row <= worksheet?.Dimension.End.Row; row++)
{
    //姓名
    var name = worksheet.Cells[row, 1].Value.ToString();

    //要設定的密碼
    var password = worksheet.Cells[row, 2].Value.ToString();

    //假設PDF文件名與姓名對應
    var pdfFilePath = Path.Combine(pdfFolderPath, $"{name}.pdf");

    //加密
    if (password != null)
    {
        EncryptPdf(pdfFilePath, pdfFilePath, password);
    }
}

static void EncryptPdf(string inputPdfPath, string outputPdfPath, string password)
{
    using (var reader = new PdfReader(inputPdfPath))
    {
        using (var fs = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
        {
            using (var stamper = new PdfStamper(reader, fs))
            {
                stamper.SetEncryption(
                    Encoding.ASCII.GetBytes(password), // 轉換密碼為位元組
                    Encoding.ASCII.GetBytes(password),
                    PdfWriter.ALLOW_PRINTING, // 允許列印
                    PdfWriter.ENCRYPTION_AES_128); // 使用AES 128位元加密
            }
        }
    }
}