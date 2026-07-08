// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


using ClosedXML.Excel;

namespace SptUtils.ReadAsFitted.Base
{

    public static class ReadTestSheet
    {

        public static List<AsFittedCircuit> ParseTestSheet(string fileName, DateOnly? checklistDate, IXLWorksheet ws)
        {
            var isTestSheet = ws.Cell("F1").GetString() == "TEST SHEET";

            var isVersionZero = ws.Cell("B2").GetString() == "Site Name:" && ws.Cell("B5").GetString() == "CIRCUIT/CABLE DETAILS";

            var circuits = new List<AsFittedCircuit>();
            string[] columns = { "C", "D", "E", "F", "G", "H", "I", "J", "K"};
            if (isTestSheet)
            {
                var header = isVersionZero ? ReadHeaderZero(fileName, checklistDate, ws) : ReadHeader(fileName, checklistDate, ws);
                foreach (var column in columns) 
                {
                    var test = isVersionZero ? ReadCircuitZero(column, header, ws) : ReadCircuit(column, header, ws);
                    if (test != null) circuits.Add(test);
                }
            }
            return circuits;
        }

        

        private static TestHeader ReadHeader(string fileName, DateOnly? checklistDate, IXLWorksheet ws)
        {
            var tabName = ws.Name;
            var siteName = ws.Cell("B3").GetString();
            var dbOrPanelNumber = ws.Cell("E3").GetString();
            var testDate = ws.Cell("J3").GetString();
            var sheetNumber = ws.Cell("K3").GetString();
            var aibRef = ws.Cell("B5").GetString();
            var hasTpOrSp = ws.Cell("E4").GetString() == "TP / SP";
            var tpOrSp = hasTpOrSp ? ws.Cell("E5").GetString() : "";
            var locationCell = hasTpOrSp ? "F5" : "E5";
            var location = hasTpOrSp ? ws.Cell(locationCell).GetString() : ws.Cell("E5").GetString();
            var dbOrPanelIncomerDetails = ws.Cell("B7").GetString();

            return new TestHeader
            (
                FileName: fileName, 
                TabName: tabName, 
                ChecklistDate: checklistDate,
                SiteName: siteName, 
                DbOrPanelNumber: dbOrPanelNumber, 
                TestDate: testDate, 
                SheetNumber: sheetNumber, 
                AibRef: aibRef, 
                TpOrSp: tpOrSp,
                Location: location,
                DbOrPanelIncomerDetails: dbOrPanelIncomerDetails
            );
        }

        private static TestHeader ReadHeaderZero(string fileName, DateOnly? checklistDate, IXLWorksheet ws)
        {
            var tabName = ws.Name;
            var siteName = ws.Cell("B3").GetString();
            var aibRef = ws.Cell("E5").GetString();
            var location = ws.Cell("F3").GetString();
            var testDate = ws.Cell("J3").GetString();
            var sheetNumber = ws.Cell("K3").GetString();
            var dbOrPanelNumber = "";
            var tpOrSp = "";
            var dbOrPanelIncomerDetails = "";

            return new TestHeader
            (
                FileName: fileName, 
                TabName: tabName, 
                ChecklistDate: checklistDate,
                SiteName: siteName, 
                DbOrPanelNumber: dbOrPanelNumber, 
                TestDate: testDate, 
                SheetNumber: sheetNumber, 
                AibRef: aibRef, 
                TpOrSp: tpOrSp,
                Location: location,
                DbOrPanelIncomerDetails: dbOrPanelIncomerDetails
            );
        }

        // columns ["C" .. "K"]
        private static AsFittedCircuit? ReadCircuit(string col, TestHeader header, IXLWorksheet ws)
        {
            var cableNum = ws.Cell(10, col).GetString();
            var fedFrom = ws.Cell(11, col).GetString();
            var circuitRefAndPhase = ws.Cell(12, col).GetString();
            
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
                    ChecklistYear: header.ChecklistDate.HasValue ? header.ChecklistDate.Value.Year : 1970,
                    SiteName: header.SiteName,
                    DbOrPanelNumber: header.DbOrPanelNumber, 
                    HeaderTestDate: header.TestDate, 
                    SheetNumber: header.SheetNumber, 
                    AibRef: header.AibRef,
                    TpOrSp: header.TpOrSp,
                    Location: header.Location, 
                    DbOrPanelIncomerDetails: header.DbOrPanelIncomerDetails,
                    CircuitColumn: col,
                    CableNum: cableNum, 
                    FedFrom: fedFrom, 
                    CircuitRefAndPhase: circuitRefAndPhase,
                    CircuitDescription: ws.Cell(13, col).GetString(),
                    CircuitType: ws.Cell(14, col).GetString(),
                    CableType: ws.Cell(15, col).GetString(),
                    InstallationMethod: ws.Cell(16, col).GetString(), 
                    CableLength: ws.Cell(17, col).GetString(),
                    NumOfCoresCSA: ws.Cell(18, col).GetString(),
                    CircuitBreakerOrFuseRating: ws.Cell(26, col).GetString(),
                    CircuitBreakerBSAndTypeNum: ws.Cell(27, col).GetString(),
                    CircuitBreakerManufacturerAndRefNum: ws.Cell(28, col).GetString(),
                    RCDManufacturerAndType: ws.Cell(31, col).GetString(),
                    Load: ws.Cell(34, col).GetString(),
                    RatingKW: ws.Cell(35, col).GetString(),
                    FullLoadCurrentA: ws.Cell(36, col).GetString(),
                    CircuitVoltageV: ws.Cell(58, col).GetString(),
                    CircuitCurrentA: ws.Cell(59, col).GetString(),
                    TestDate: ws.Cell(60, col).GetString(),
                    Comments: ws.Cell(61, col).GetString()
                );
            }
        }

        private static AsFittedCircuit? ReadCircuitZero(string col, TestHeader header, IXLWorksheet ws)
        {
            var cableNum = ws.Cell(6, col).GetString();
            var fedFrom = ws.Cell(7, col).GetString();
            var circuitRefAndPhase = ws.Cell(8, col).GetString();
            
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
                    ChecklistYear: header.ChecklistDate.HasValue ? header.ChecklistDate.Value.Year : 1970,
                    SiteName: header.SiteName,
                    DbOrPanelNumber: header.DbOrPanelNumber, 
                    HeaderTestDate: header.TestDate, 
                    SheetNumber: header.SheetNumber, 
                    AibRef: header.AibRef,
                    TpOrSp: header.TpOrSp,
                    Location: header.Location, 
                    DbOrPanelIncomerDetails: header.DbOrPanelIncomerDetails,
                    CircuitColumn: col,
                    CableNum: cableNum, 
                    FedFrom: fedFrom, 
                    CircuitRefAndPhase: circuitRefAndPhase,
                    CircuitDescription: ws.Cell(9, col).GetString(),
                    CircuitType: ws.Cell(10, col).GetString(),
                    CableType: ws.Cell(11, col).GetString(),
                    InstallationMethod: ws.Cell(12, col).GetString(), 
                    CableLength: ws.Cell(13, col).GetString(),
                    NumOfCoresCSA: ws.Cell(14, col).GetString(),
                    CircuitBreakerOrFuseRating: ws.Cell(22, col).GetString(),
                    CircuitBreakerBSAndTypeNum: ws.Cell(23, col).GetString(),
                    CircuitBreakerManufacturerAndRefNum: ws.Cell(24, col).GetString(),
                    RCDManufacturerAndType: ws.Cell(27, col).GetString(),
                    Load: ws.Cell(30, col).GetString(),
                    RatingKW: ws.Cell(31, col).GetString(),
                    FullLoadCurrentA: ws.Cell(32, col).GetString(),
                    CircuitVoltageV: ws.Cell(56, col).GetString(),
                    CircuitCurrentA: ws.Cell(57, col).GetString(),
                    TestDate: ws.Cell(58, col).GetString(),
                    Comments: ws.Cell(59, col).GetString()
                );
            }
        }

    }
}

