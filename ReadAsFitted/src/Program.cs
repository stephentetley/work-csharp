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
using Microsoft.Extensions.FileSystemGlobbing;

namespace SptUtils.ReadAsFitted 
{

    class Program
    {
        static int Main(string[] args) 
        {

            RootCommand rootCommand = new("ReadAsFitted") 
            {
                new Option<string>("--as_fitted_folder")
                {
                    Description = "Path to folder with As Fitted xlsx files",
                    Required = true
                },              
            };
            rootCommand.SetAction(parseResult =>
            {
                string? asFittedFolder = parseResult.GetValue<string>("--as_fitted_folder");
                return ProcessWorkbooks(asFittedFolder);
            });
            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
        }

        private static int ProcessWorkbooks(string? asFittedFolder)
        {
            if (Directory.Exists(asFittedFolder)) {
                int ans = 0;
                Matcher matcher = new();
                matcher.AddInclude("*.xlsx");
                Console.WriteLine("Source: " + asFittedFolder);

                var matchingFiles = matcher.GetResultsInFullPath(asFittedFolder);
                foreach (var asFitted in matchingFiles)
                {
                    var res = ProcessWorkbook(asFitted);
                    ans += res;
                }
                return (ans > 1) ? 1 : 0;
            } 
            else
            {
                return 1;
            }

        }

        private static int ProcessWorkbook(string? asFittedPath)
        {
            if (File.Exists(asFittedPath))
            {
                var outputJsonFile = Path.ChangeExtension(asFittedPath, ".json");
                var options = new JsonWriterOptions { Indented = true };
                using var workbook = new XLWorkbook(asFittedPath);
                using var stream = File.Create(outputJsonFile);
                using var writer = new Utf8JsonWriter(stream, options);
                var fileName = Path.GetFileName(asFittedPath);
                var checklistDate = ReadGeneralChecklist.GetChecklistDate(workbook);
                writer.WriteStartArray();
                foreach (var worksheet in workbook.Worksheets)
                {
                    
                    var tests = ReadTestSheet.ParseTestSheet(fileName, checklistDate, worksheet);
                    if (tests.Count > 0) Console.WriteLine("Sheet: " + worksheet.Name);
                    foreach(var test in tests) test.WriteJson(writer);
                    
                }
                writer.WriteEndArray();
                Console.WriteLine("Wrote: " + outputJsonFile);
                return 0;
            } else {
                    Console.WriteLine("Invalid arguments");
                    return 1;
            }
        }
    }
}


