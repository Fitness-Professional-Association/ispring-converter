using System.CommandLine;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using ispring_converter;
using OfficeOpenXml;


class Program
{
    public static async Task<int> Main(string[] args)
    {
        var rootCommand = new RootCommand("ISpring FPA helper ");
    
        var converterCommand = new ConverterCommand();

        rootCommand.AddCommand(converterCommand);

        return await rootCommand.InvokeAsync(args);
    }
}