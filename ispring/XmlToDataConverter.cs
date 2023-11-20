using System.Runtime.InteropServices.JavaScript;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ispring_converter;

public class XmlToDataConverter
{
    public Question Convert(IEnumerable<OpenXmlElement> nodes)
    {
        var elements = nodes.ToArray();
        var (number, variant, points, attempts) = ParseQuestionNode(GetTextFromParagraph(elements[0]));

        var feedbackTable = elements[3];
        var (correct, incorrect) = GetIfCorrectFeedBackString(feedbackTable);
        
        var answerTable = elements[2];
        var rows = answerTable.Where(x => x is TableRow).Skip(1);
        var answers = rows.Select(ParseAnswer).Select(CreateAnswer);
        
        var question = new Question(
            ushort.TryParse(number, out var nbr) ? nbr : default,
            GetTextFromParagraph(elements[1]),
            IsMultiple(variant),
            byte.TryParse(points, out var pts) ? pts : default,
            byte.TryParse(attempts, out var atms) ? atms : default,
            correct,
            incorrect,
            answers.ToArray()
        );
        return question;
    }

    private Answer CreateAnswer((string marker, string text) value)
    {
        var (marker, text) = value;
        var answer = new Answer(
            IsValid: marker == "V",
            text
        );
        return answer;
    }

    private (string, string) ParseAnswer(OpenXmlElement item)
    {
        if (item is TableRow == false) throw new ArgumentException("Item is not a TableRow");

        var element = item as TableRow;

        var isCorrect = element.First(x => x is TableCell).Last() is Paragraph correct ? correct.InnerText : "";
        var value = element.Last(x => x is TableCell).Last() is Paragraph text ? text.InnerText : "";

        return (isCorrect, value);
    }

    private (string, string) GetIfCorrectFeedBackString(OpenXmlElement table)
    {
        var rows = table.Where(x => x is TableRow).Skip(1);

        var isCorrect = rows.First().Last(x => x is TableCell).Last() is Paragraph correct ? correct.InnerText : "";
        var isInCorrect = rows.Last().Last(x => x is TableCell).Last() is Paragraph incorrect
            ? incorrect.InnerText
            : "";

        return (isCorrect, isInCorrect);
    }

    private static string GetTextFromParagraph(OpenXmlElement item)
    {
        if (item is Paragraph paragraph) return paragraph.InnerText;
        throw new ArgumentException("Элемент не является параграфом");
    }

    private (string,string,string,string) ParseQuestionNode(string str)
    {
        const string pattern = @"Вопрос (\d+)\. (.+), (\d+) баллов, (\d+) попытка";
        var match = Regex.Match(str, pattern);

        var number = match.Groups[1].Value;
        var variant = match.Groups[2].Value;
        var points = new string(match.Groups[3].Value);
        var attempts = match.Groups[4].Value.Normalize();

        var result = match.Success
            ? new ParseResult(number, variant, points, attempts)
            : default;

        //return result;
        return (number, variant, points, attempts);
    }

    private bool IsMultiple(string value)
    {
        var result = value switch
        {
            "Выбор нескольких ответов" => true,
            "Последовательность" => true,
            _ => false
        };

        return result;
    }

    private record ParseResult(string Number, string Variant, string Points, string Attempts);
}