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

        public static List<DistBoardCircuit> ParseDBSchedule(string fileName, IXLWorksheet ws)
        {
            // var projectReference = ws.Cell("E5").GetString();
            // var dbReference = ws.Cell("E7").GetString() + "-" + ws.Cell("F7").GetString();

            // return new DistBoardSchedule
            // (
            //     FileName: fileName,
            //     TabName: ws.Name,
            //     ProjectReference: projectReference,
            //     DbReference: dbReference,
            //     SupplyCableRef: ws.Cell("E14").GetString(),
            //     NumberOfWays: null,
            //     FedFrom: "",
            //     ProtectiveDeviceA: null,
            //     Phase: "",
            //     FaultCurrentkA: null,
            //     Circuits: []
            // );

            return [];
        }

        private static CircuitHeader ReadHeader(string fileName, IXLWorksheet ws)
        {
            var dbReference = ws.Cell("E7").GetString() + "-" + ws.Cell("F7").GetString();
            int? getInt(string cellAddr)
            {
                if (ws.Cell(cellAddr).TryGetValue<int>(out var cellValue))
                {
                    return cellValue;
                }
                else
                {
                    return null;
                }
            };
            double? getDouble(string cellAddr)
            {
                if (ws.Cell(cellAddr).TryGetValue<double>(out var cellValue))
                {
                    return cellValue;
                }
                else
                {
                    return null;
                }
            };

            return new CircuitHeader
            (
                FileName: fileName,
                TabName: ws.Name,
                ProjectReference: ws.Cell("E5").GetString(),
                DbReference: dbReference,
                SupplyCableRef: ws.Cell("E14").GetString(),
                NumberOfWays: getInt("E18"),
                FedFrom: ws.Cell("J14").GetString(),
                ProtectiveDeviceA: getDouble("J16"),
                DistBoardPhase: ws.Cell("J18").GetString(),
                FaultCurrentkA: getDouble("J20")
            );

        }

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

