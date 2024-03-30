using ClosedXML.Excel;
using DocumentReader;
using FluentValidation.Results;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using Path = System.IO.Path;
using ValidationResult = FluentValidation.Results.ValidationResult;

// Load config from json file
Config config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("config.json", false, true)
                    .Build()
                    .Get<Config>() ??
                new Config();

// Validate config
var validator = new ConfigValidator();
ValidationResult? validationResult = validator.Validate(config);
if (!validationResult.IsValid)
{
    AnsiConsole.MarkupLine("[bold red]Validation errors in config.json:[/]");
    foreach (ValidationFailure? error in validationResult.Errors)
    {
        AnsiConsole.MarkupLine($"[red]{error.ErrorMessage}[/]");
    }

    return;
}

var values = new List<(string name, string value)>();

string pdfPath = string.Empty;
#if DEBUG
pdfPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "China");
#else
pdfPath = ".";
#endif
string[] files = Directory.GetFiles(pdfPath, "*.pdf", SearchOption.TopDirectoryOnly);

AnsiConsole.Status()
    .Start(
        "Processing PDF files...",
        ctx =>
        {
            foreach (string file in files)
            {
                using var pdfDoc = new PdfDocument(new PdfReader(file));
                PdfPage page = pdfDoc.GetPage(1);
                foreach (Area configArea in config.Areas)
                {
                    string? text = GetTextFromLocation(
                        configArea.X,
                        configArea.Y,
                        configArea.Width,
                        configArea.Height,
                        page);
                    if (text == null)
                    {
                        AnsiConsole.MarkupLine($"[yellow]No text found for {configArea.Name}[/]");
                        continue;
                    }

                    values.Add((configArea.Name, text));
                }

                ctx.Spinner(Spinner.Known.Star);
                ctx.Status($"Processed: {Path.GetFileName(file)}");
            }
        });

using var workbook = new XLWorkbook();
IXLWorksheet? worksheet = workbook.Worksheets.Add("Sheet1");
var index = 1;

var groupedValues = values.GroupBy(x => x.name)
    .Select(x => new { name = x.Key, values = x.Select(y => y.value).ToList() })
    .ToList();
worksheet.Cell(1, index++).Value = "File Name";
foreach (var group in groupedValues)
{
    worksheet.Cell(1, index++).Value = group.name;
}

for (var i = 0; i < files.Length; i++)
{
    worksheet.Cell(i + 2, 1).Value = Path.GetFileName(files[i]);
    for (var j = 0; j < groupedValues.Count; j++)
    {
        worksheet.Cell(i + 2, j + 2).Value = groupedValues[j].values[i];
    }
}

var end = (char)('A' + groupedValues.Count);
IXLTable table = worksheet.Range($"A1:{end}{groupedValues.Count + 2}").CreateTable();
table.Theme = XLTableTheme.TableStyleLight15;
table.SetShowRowStripes(false);
table.SetShowColumnStripes(false);

worksheet.Columns().AdjustToContents();
var excelFileName = $"Excel-{DateTime.Now:hh-mm-ss}.xlsx";
workbook.SaveAs(excelFileName);
AnsiConsole.MarkupLine($"[green]Excel file generated: {excelFileName}[/]");

static string GetTextFromLocation(float x, float y, float width, float height, PdfPage page)
{
    float pdfPageHeight = page.GetPageSizeWithRotation().GetHeight();
    Rectangle rect = GetPointsRectangle(x, y, width, height, pdfPageHeight);
    var strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), new TextRegionEventFilter(rect));
    string extractedText = PdfTextExtractor.GetTextFromPage(page, strategy);
    return extractedText;
}

static Rectangle GetPointsRectangle(float x, float y, float width, float height, float pdfPageHeight)
{
    const float inchesToPoints = 72f; // 1 inch = 72 points

    // convert inches to points
    x *= inchesToPoints;
    y = pdfPageHeight - y * inchesToPoints - height * inchesToPoints; // convert to bottom-left corner
    width *= inchesToPoints;
    height *= inchesToPoints;

    return new Rectangle(x, y, width, height);
}
