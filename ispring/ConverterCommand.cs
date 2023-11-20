using System.CommandLine;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeOpenXml;

namespace ispring_converter;

public class ConverterCommand : System.CommandLine.Command
{
    public ConverterCommand() : base("converter", "Word to Excel converter")
    {
        var keyOption = new Option<string>(
                "--file",
                "Select a word.docx file to convert"
            );
        
        var titleOption = new Option<string>(
            "--out",
            "Specify the folder to lay out the results"
        );
        
        this.AddOption(keyOption);
        this.AddOption(titleOption);
        
        this.SetHandler(async (string fileName, string outPutFolder) => InvokeCommand(fileName, outPutFolder), keyOption, titleOption);
    }

    private async Task InvokeCommand(string fileName, string title)
    {
        if (fileName == default)
        {
            const string error = "The filename was not specified";
            Console.WriteLine(error);
            throw new ArgumentException(error);
        }

        var workingPath = new WorkingPath(fileName);

        using var wordDocument = WordprocessingDocument.Open(workingPath.Value, false);
        var excelDocument = new ExcelPackage();

        Convert(wordDocument, excelDocument);

        //var documentsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var excelPath = Path.Combine(workingPath.Folder, Path.GetFileName("ispring.xlsx"));
        
        var excelFile = new FileInfo(excelPath);
        excelDocument.SaveAs(excelFile);

        Console.WriteLine("Storing excel file to Documents folder");
        
        Console.WriteLine("Done.");
    }
    private void Convert(WordprocessingDocument wordDocument, ExcelPackage excelDocument)
    {
        var parser = new WordDocumentParser(wordDocument);
        var converter = new XmlToDataConverter();
        var excelBuilder = new ExcelDocumentBuilder(excelDocument, "Some title of questionare");

        var models = parser.QuestionNodeGroups.Select(converter.Convert);

        Console.WriteLine($"{models.Count()} results were added");

        excelBuilder
            .AddColumn("Тип вопроса", (q) => q.IsMultiply ? "MR" : "MC")
            .AddColumn("Текст вопроса", (q) => q.Text)
            .AddColumn("Ответ 1", (q) => GetAnswer(q, 1))
            .AddColumn("Ответ 2", (q) => GetAnswer(q, 2))
            .AddColumn("Ответ 3", (q) => GetAnswer(q, 3))
            .AddColumn("Ответ 4", (q) => GetAnswer(q, 4))
            .AddColumn("Ответ 5", (q) => GetAnswer(q, 5))
            // .AddColumn("Ответ 6", (q) => GetAnswer(q, 6))
            // .AddColumn("Ответ 7", (q) => GetAnswer(q, 7))
            // .AddColumn("Ответ 8", (q) => GetAnswer(q, 8))
            // .AddColumn("Ответ 9", (q) => GetAnswer(q, 9))
            // .AddColumn("Ответ 10", (q) => GetAnswer(q, 10))
            .AddColumn("Сообщение, если верно", question => question.MessageIfCorrect)
            .AddColumn("Сообщение, если не верно", question => question.MessageIfIncorrect)
            //.AddColumn("Баллы", question => question.Points.ToString())
            .AddRange(models);

        excelBuilder.Build();

        string GetAnswer(Question question, int position)
        {
            var stringBuilder = new StringBuilder();
            if (question.Answers.Length < position) return "";
            var answer = question.Answers[position - 1];
            if (answer.IsValid) stringBuilder.Append('*');
            stringBuilder.Append(answer.Text);
            return stringBuilder.ToString();
        }
    }
}