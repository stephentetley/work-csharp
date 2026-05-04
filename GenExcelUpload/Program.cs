using System;
using System.CommandLine;
using System.CommandLine.Parsing;

class Program
{
    static int Main(string[] args) 
    {

        RootCommand rootCommand = new("GenFileUpload") 
        {
            new Option<string>("--uploader_template_file")
            {
                Description = "Path to Excel Uploader template file",
                Required = true
            },
            new Option<string>("--database_file")
            {
                Description = "The DuckDB database file",
                Required = true
            },
            new Option<string>("--output_xlsx_file")
            {
                Description = "Output file positional format string (applied to batch number)",
                Required = true
            },
            new Option<string>("--excel_uploader_output_type")
            {
                Description = "Output type: floc_create, equi_create, equi_change",
                Required = true
            },
            new Option<string>("--title")
            {
                Description = "Change request title positional format string (applied to batch number and date.today())",
                Required = true
            },
        };
        rootCommand.SetAction(parseResult =>
        {
            string? uploaderTemplateFile = parseResult.GetValue<string>("--uploader_template_file");
            string? databaseFile = parseResult.GetValue<string>("--database_file");
            string? outputXlsxFile = parseResult.GetValue<string>("--output_xlsx_file");
            string? excelUploaderOutputType = parseResult.GetValue<string>("--excel_uploader_output_type");
            string? title = parseResult.GetValue<string>("--title");
            Console.WriteLine($"GenExcelUpload:");
            Console.WriteLine($"{uploaderTemplateFile}");
            Console.WriteLine($"{databaseFile}");
            Console.WriteLine($"{outputXlsxFile}");
            Console.WriteLine($"{excelUploaderOutputType}");
            Console.WriteLine($"{title}");
            return 0;
        });

        ParseResult parseResult = rootCommand.Parse(args);
        return parseResult.Invoke();

        // if (parseResult.Errors.Count == 0)
        // {
        //     var sourceFile = parseResult.GetValue(fileName);
        //     Console.WriteLine($"GenExcelUpload {sourceFile}");
        //     return 0;
        // }
        // foreach (ParseError parseError in parseResult.Errors)
        // {
        //     Console.Error.WriteLine(parseError.Message);
        // }
        // return 1;
        
    }
}