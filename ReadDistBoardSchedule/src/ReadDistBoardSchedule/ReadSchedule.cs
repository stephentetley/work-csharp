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

namespace SptUtils.ReadDistBoardSchedule 
{

    public static class ReadSchedule
    {

        public static DistBoardSchedule ParseDBSchedule(string fileName, IXLWorksheet ws)
        {
            var projectReference = ws.Cell("E5").GetString();
            var DbReference = ws.Cell("E7").GetString() + "-" + ws.Cell("F7").GetString();

            return new DistBoardSchedule
            (
                FileName: fileName,
                TabName: ws.Name,
                ProjectReference: projectReference,
                DbReference: DbReference,
                SupplyCableRef: ws.Cell("E14").GetString(),
                NumberOfWays: null,
                FedFrom: "",
                ProtectiveDeviceA: null,
                Phase: "",
                FaultCurrentkA: null,
                Circuits: []
            );
        }

        // private Header ReadHeader()
        // {
        //     var tabName = sheet.Name;
        //     var siteName = sheet.Cell("B3").GetString();
        //     var dbOrPanelNumber = sheet.Cell("E3").GetString();
        //     var testDate = sheet.Cell("J3").GetString();
        //     var sheetNumber = sheet.Cell("K3").GetString();
        //     var aibRef = sheet.Cell("B5").GetString();
        //     var tpOrSp = sheet.Cell("E5").GetString();
        //     var location = sheet.Cell("F5").GetString();
        //     var dbOrPanelIncomerDetails = sheet.Cell("B7").GetString();

        //     return new Header
        //     (
        //         FileName: fileName, 
        //         TabName: tabName, 
        //         SiteName: siteName, 
        //         DbOrPanelNumber: dbOrPanelNumber, 
        //         TestDate: testDate, 
        //         SheetNumber: sheetNumber, 
        //         AibRef: aibRef, 
        //         TpOrSp: tpOrSp,
        //         Location: location,
        //         DbOrPanelIncomerDetails: dbOrPanelIncomerDetails
        //     );
        // }

        // // columns ["C" .. "K"]
        // private DbSchedule? ReadSchedule(string col, Header header)
        // {
        //     var cableNum = sheet.Cell(10, col).GetString();
        //     var fedFrom = sheet.Cell(11, col).GetString();
        //     var circuitRefAndPhase = sheet.Cell(12, col).GetString();
            
        //     if (cableNum == "" && fedFrom == "" && circuitRefAndPhase == "")
        //     {
        //         return null;
        //     } 
        //     else
        //     {
        //         return new DbSchedule
        //         (
        //             FileName: header.FileName, 
        //             TabName: header.TabName, 
        //             SiteName: header.SiteName,
        //             DbOrPanelNumber: header.DbOrPanelNumber, 
        //             HeaderTestDate: header.TestDate, 
        //             SheetNumber: header.SheetNumber, 
        //             AibRef: header.AibRef,
        //             TpOrSp: header.TpOrSp,
        //             Location: header.Location, 
        //             DbOrPanelIncomerDetails: header.DbOrPanelIncomerDetails,
        //             CableNum: cableNum, 
        //             FedFrom: fedFrom, 
        //             CircuitRefAndPhase: circuitRefAndPhase,
        //             CircuitDescription: sheet.Cell(13, col).GetString(),
        //             CircuitType: sheet.Cell(14, col).GetString(),
        //             CableType: sheet.Cell(15, col).GetString(),
        //             InstallationMethod: sheet.Cell(16, col).GetString(), 
        //             CableLength: sheet.Cell(17, col).GetString(),
        //             NumOfCoresCSA: sheet.Cell(18, col).GetString(),
        //             CircuitBreakerOrFuseRating: sheet.Cell(26, col).GetString(),
        //             CircuitBreakerBSAndTypeNum: sheet.Cell(27, col).GetString(),
        //             CircuitBreakerManufacturerAndRefNum: sheet.Cell(28, col).GetString(),
        //             RCDManufacturerAndType: sheet.Cell(31, col).GetString(),
        //             Load: sheet.Cell(34, col).GetString(),
        //             RatingKW: sheet.Cell(35, col).GetString(),
        //             FullLoadCurrentA: sheet.Cell(36, col).GetString(),
        //             CircuitVoltageV: sheet.Cell(58, col).GetString(),
        //             CircuitCurrentA: sheet.Cell(59, col).GetString(),
        //             TestDate: sheet.Cell(60, col).GetString(),
        //             Comments: sheet.Cell(61, col).GetString()
        //         );
        //     }
        // }

    }
}

