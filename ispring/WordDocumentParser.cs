using System.Collections.Immutable;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ispring_converter;

public class WordDocumentParser
{
    private readonly ICollection<IEnumerable<OpenXmlElement>> batch = new List<IEnumerable<OpenXmlElement>>();
    public ImmutableArray<IEnumerable<OpenXmlElement>> QuestionNodeGroups => batch.ToImmutableArray();

    public WordDocumentParser(WordprocessingDocument document)
    {
        ParseXMLBlocks(document.MainDocumentPart.Document.Body);
    }

    private void ParseXMLBlocks(Body body)
    {
        var elements = body.ChildElements;
        
        for (var i = 0; i < elements.Count; i++)
        {
            if (IsQuestionParagraph(elements[i]) == false) continue;
            var chunk = ChunkArray(elements.Skip(i));
            batch.Add(chunk);
        }
    }

    private IEnumerable<OpenXmlElement> ChunkArray(IEnumerable<OpenXmlElement> array)
    {
        var elements = array.ToArray();
        
        yield return elements[0];
        yield return elements[1];
        
        var firstTable = elements.FirstOrDefault(x => x is Table);
        int index = Array.IndexOf(elements, firstTable);
        
        yield return elements[index];

        var secondTable = elements.Skip(index+1).FirstOrDefault(x => x is Table);
        index = Array.IndexOf(elements, secondTable);
        
        yield return elements[index];
    }
    
    private bool IsQuestionParagraph(OpenXmlElement element) => 
        element is Paragraph && element.InnerText.Contains("Вопрос", StringComparison.Ordinal);
}