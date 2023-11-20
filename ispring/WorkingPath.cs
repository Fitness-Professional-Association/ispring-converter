namespace ispring_converter;

public class WorkingPath
{
    private readonly string givenFileString;
    public string Value { get; private set; }
    public string Folder => Path.GetDirectoryName(Value);

    public WorkingPath(string givenFileString)
    {
        this.givenFileString = givenFileString ?? throw new ArgumentException("File was not specified.");
        GetWorkingDirectory();
    }

    private void GetWorkingDirectory()
    {
        Value = Path.IsPathRooted(givenFileString)
            ? givenFileString
            : Path.Combine(Directory.GetCurrentDirectory(), givenFileString);

        if (File.Exists(Value) != false) return;
        Console.WriteLine($"File was not found at {Value}");
        Console.WriteLine($"Trying to find the file in the Documents folder");

        var documentsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filePathInDocuments = Path.Combine(documentsDirectory, Path.GetFileName(Value));

        if (File.Exists(filePathInDocuments))
        {
            Value = filePathInDocuments;
            Console.WriteLine($"File was found at {filePathInDocuments}");
        }
        else
        {
            Console.WriteLine($"File was not found.");
        }
    }
}