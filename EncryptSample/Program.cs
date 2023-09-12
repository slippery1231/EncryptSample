// See https://aka.ms/new-console-template for more information

using EncryptSample;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

//非商業使用，如果沒加這行會被擋下來
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//取得config設定
var fileConfig = GetConfig();

//取得excel路徑
var excelFilePath = fileConfig.ExcelFilePath;

//取得pdf路徑
var pdfFolderPath = fileConfig.PdfFilePath;

var package = new ExcelPackage(new FileInfo(excelFilePath));
var worksheet = package.Workbook.Worksheets[0];

//excel內容會從第二行開始，(第一行是標題)
for (var row = 2; row <= worksheet?.Dimension.End.Row; row++)
{
    //姓名
    var name = worksheet.Cells[row, 1].Value.ToString();

    //要設定的密碼
    var password = worksheet.Cells[row, 2].Value.ToString();

    //PDF文件名與姓名對應
    var pdfFilePath = Path.Combine(pdfFolderPath, $"{name}.pdf");

    //output
    var outPutFilePath = fileConfig.OutputFilePath;
    var outPutFileName = Path.Combine(outPutFilePath, $"{name}.pdf");
    //加密
    if (password != null)
    {
        EncryptPdf(pdfFilePath, outPutFileName, password);
    }

    Console.WriteLine($"第{row - 1}筆已執行完成");
}

Console.WriteLine("執行成功，檔案已全部加密完成");
Console.WriteLine("請按任意鍵關閉視窗");
Console.ReadLine();

static void EncryptPdf(string inputPdfPath, string outputPdfPath, string password)
{
    var reader = new PdfReader(inputPdfPath);
    var fileStream = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write);
    PdfEncryptor.Encrypt(reader, fileStream, true, password, password, PdfWriter.ALLOW_PRINTING);

    reader.Close();
}

FileConfig GetConfig()
{
    var config = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
        .Build();
    var fConfig = new FileConfig();
    config.Bind("FileConfig", fConfig);
    return fConfig;
}