using DocumentFormat.OpenXml;
using OfficeOpenXml;

namespace ispring_converter;

public class ExcelDocumentBuilder
{
    private readonly ExcelWorksheet sheet;
    private readonly LinkedList<(string, Func<Question, string>)> headers = new LinkedList<(string, Func<Question, string> function)>();
    private IEnumerable<Question> data;
    
    public ExcelDocumentBuilder(ExcelPackage package, string sheetTitle)
    {
        sheet = package
            .Workbook
            .Worksheets.Add(sheetTitle);
    }

    public ExcelDocumentBuilder AddColumn(string title, Func<Question, string> function)
    {
        headers.AddLast((title, function));
        return this;
    }

    public ExcelDocumentBuilder AddRange(IEnumerable<Question> items)
    {
        data = items.OrderBy(x=>x.Number);
        return this;
    }

    public void Build()
    {
        var columnNumber = 0;
        var rowNumber = 0;

        foreach (var header in headers)
        {
            var (title, function) = header;
            rowNumber++; columnNumber++;
            
            sheet.Cells[rowNumber, columnNumber].Value = title;
            
            foreach (var question in data)
            {
                rowNumber++;
                sheet.Cells[rowNumber, columnNumber].Value = function(question);
            }

            rowNumber = 0;
        }
    }
}