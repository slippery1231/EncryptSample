// See https://aka.ms/new-console-template for more information

using EncryptSample;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using Serilog;


Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .WriteTo.Console()
    .WriteTo.File("Log/log.txt", rollingInterval: RollingInterval.Day)
    .CreateLogger();

//非商業使用，如果沒加這行會被擋下來
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//取得config設定
var fileConfig = GetConfig();



//取得pdf路徑
var pdfFolderPath = fileConfig.PdfFilePath;

//處理excel
var excelFilePath = fileConfig.ExcelFilePath;
var package = new ExcelPackage(new FileInfo(excelFilePath));
var worksheet = package.Workbook.Worksheets[0];

//excel內容會從第二行開始，(第一行是標題)
for (var row = 2; row <= worksheet?.Dimension.End.Row; row++)
{
    Log.Information($"開始執行第{row - 1}筆");
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

    Log.Information($"第{row - 1}筆已執行完成");
}

Log.Information("執行成功，檔案已全部加密完成");
Log.Information("請按任意鍵關閉視窗");
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