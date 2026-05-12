// Copyright 2026 Stephen Tetley

// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


using System.CommandLine;
using DuckDB.NET.Data;

namespace SptUtils.GenExcelUpload
{

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
                    Description = "Output file format string (C#) applied to batch number",
                    Required = true
                },
                new Option<string>("--excel_uploader_output_type")
                {
                    Description = "Output type: floc_create, equi_create, equi_change",
                    Required = true
                },
                new Option<string>("--title")
                {
                    Description = "Change request title positional (SQL/DuckDb) format string applied to batch number and date.today()",
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

                // TODO validate databaseFile
                string connstr = $"DataSource = {databaseFile};ACCESS_MODE=READ_WRITE;";
                using var conn = new DuckDBConnection(connstr);
                conn.Open();

                switch (excelUploaderOutputType)
                {
                    case "floc_create":
                        var flocMake = new FlocCreate(conn);
                        flocMake.InsertTitleFormatString(title ?? "bad");
                        flocMake.WriteFlocCreateUpload(uploaderTemplateFile ?? "bad", outputXlsxFile ?? "bad");
                        return 0;
                    
                    case "equi_create":
                        var equiMake = new EquiCreate(conn);
                        equiMake.InsertTitleFormatString(title ?? "bad");
                        equiMake.WriteEquiCreateUpload(uploaderTemplateFile ?? "bad", outputXlsxFile ?? "bad");
                        return 0;
                    
                    default:
                        Console.WriteLine($"Unrecognized excel_uploader_output_type: {excelUploaderOutputType}");
                        return 1;
                }
            });

            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
           
        }
    }

}
