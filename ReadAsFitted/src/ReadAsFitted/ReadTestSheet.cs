// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System.Net.WebSockets;
using ClosedXML.Excel;

namespace SptUtils.ReadAsFitted 
{

    public class ReadTestSheet
    {
        IXLWorksheet sheet;
        string fileName;

        public ReadTestSheet(string xlsxName, IXLWorksheet ws)
        {
            sheet = ws;
            fileName = xlsxName;
        }

        public bool isTestSheet()
        {
            var value = sheet.Cell("F1").GetString();
            return (value == "TEST SHEET");

        }

        public void readTest()
        {
            var header = readHeader();
            var test1 = readCircuit("C", header);
            Console.WriteLine(test1);
        }

        private TestHeader readHeader()
        {
            var tabName = sheet.Name;
            var siteName = sheet.Cell("B3").GetString();
            var dbOrPanelNumber = sheet.Cell("E3").GetString();
            var testDate = sheet.Cell("J3").GetString();
            var sheetNumber = sheet.Cell("K3").GetString();
            var aibRef = sheet.Cell("B5").GetString();
            var location = sheet.Cell("B5").GetString();

            return new TestHeader
            (
                FileName: fileName, 
                TabName: tabName, 
                SiteName: siteName, 
                DbOrPanelNumber: dbOrPanelNumber, 
                TestDate: testDate, 
                SheetNumber: sheetNumber, 
                AibRef: aibRef, 
                Location: location
            );
        }

        // columns ["C" .. "K"]
        private AsFittedCircuit? readCircuit(string col, TestHeader header)
        {
            var cableNum = sheet.Cell(10, col).GetString();
            var fedFrom = sheet.Cell(11, col).GetString();
            var circuitRefAndPhase = sheet.Cell(12, col).GetString();
            var circuitDescription = sheet.Cell(13, col).GetString();
            var circuitType = sheet.Cell(14, col).GetString();

            if (cableNum == "" && fedFrom == "" && circuitRefAndPhase == "")
            {
                return null;
            } 
            else
            {
                return new AsFittedCircuit
                (
                    FileName: header.FileName, 
                    TabName: header.TabName, 
                    SiteName: header.SiteName,
                    DbOrPanelNumber: header.DbOrPanelNumber, 
                    TestDate: header.TestDate, 
                    SheetNumber: header.SheetNumber, 
                    AibRef: header.AibRef,
                    Location: header.Location, 
                    CableNum: cableNum, 
                    FedFrom: fedFrom, 
                    CircuitRefAndPhase: circuitRefAndPhase,
                    CircuitDescription: circuitDescription,
                    CircuitType: circuitType
                );
            }
        }

    }
}

