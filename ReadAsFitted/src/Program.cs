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
using System.Text.Json;
using ClosedXML.Excel;

namespace SptUtils.ReadAsFitted 
{

    class Program
    {
        static int Main(string[] args) 
        {

            RootCommand rootCommand = new("ReadAsFitted") 
            {
                new Option<string>("--as_fitted_file")
                {
                    Description = "Path to As Fitted xlsx file",
                    Required = true
                },
                new Option<string>("--output_json_file")
                {
                    Description = "Path to output file (json)",
                    Required = true
                },                
            };
            rootCommand.SetAction(parseResult =>
            {
                string? asFitted = parseResult.GetValue<string>("--as_fitted_file");
                string? outputJsonFile = parseResult.GetValue<string>("--output_json_file");
                return ProcessWorkbook(asFitted, outputJsonFile);
            });
            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
        }

        private static int ProcessWorkbook(string? asFitted, string? outputJsonFile)
        {
            if (File.Exists(asFitted) && outputJsonFile != null)
            {
                var options = new JsonWriterOptions { Indented = true };
                using var workbook = new XLWorkbook(asFitted);
                using var stream = File.Create(outputJsonFile);
                using var writer = new Utf8JsonWriter(stream, options);
                var fileName = Path.GetFileName(asFitted);
                writer.WriteStartArray();
                foreach (var worksheet in workbook.Worksheets)
                {
                    var reader = new ReadTestSheet(fileName, worksheet);

                    if (reader.IsTestSheet())
                    {
                        Console.WriteLine("Sheet: " + worksheet.Name + " " + reader.IsTestSheet());
                        var tests = reader.ParseTestSheet();
                        foreach(var test in tests) test.WriteJson(writer);
                    }
                }
                Console.WriteLine("Source: " + asFitted);
                writer.WriteEndArray();
                return 0;
            } else {
                    Console.WriteLine("Invalid arguments");
                    return 1;
            }
        }
    }
}


