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
using Microsoft.Extensions.Configuration;


using SptUtils.FlocDelta.Base;

namespace SptUtils.FlocDelta
{

    class Program
    {

        static int Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();
            
            var appSettings = builder.GetSection("AppSettings").Get<AppSettings>();

            RootCommand rootCommand = new("FlocDelta") 
            {
                new Option<string>("--worklist_xlsx")
                {
                    Description = "The input worklist, an Excel *.xlsx file",
                    Required = true
                },
                new Option<string>("--sqlite_output_file")
                {
                    Description = "The SQLite database file to be created",
                    Required = true
                },

            };            
            rootCommand.SetAction(parseResult =>
            {

                if (appSettings != null 
                        && parseResult.GetValue<string>("--worklist_xlsx") is string worklistXlsx
                        && parseResult.GetValue<string>("--sqlite_output_file") is string sqliteOutput)
                {
                    Console.WriteLine($"FlocDelta:");
                    Console.WriteLine($"{sqliteOutput}");

                    RunFlocDelta.GenerateSQLiteDB(appSettings, worklistXlsx, sqliteOutput);

                    return 0;
                }
                else
                {
                    return 1;
                }
                
            });

            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
        }
    }
}

