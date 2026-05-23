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
            int getLastRowNumber()
            {

                var lastRowUsed = ws.LastRowUsed();

                return lastRowUsed?.RowNumber() ?? 0;
            };
        
            var header = ReadHeader(fileName, ws);
            var circuits = new List<DistBoardCircuit>();
            for (int i = 26; i < getLastRowNumber(); i++)
            {
                var row = ws.Row(i);
                var circuit = ReadCircuit(header, row);

                if (circuit != null)
                {
                    circuits.Add(circuit);
                }

            };

            return circuits;

        }
        
        
        private static int? ReadInt(IXLCell? cell)
        {
            if (cell != null && cell.TryGetValue<int>(out var cellValue))
            {
                return cellValue;
            }
            else
            {
                return null;
            }
        }

        private static double? ReadDouble(IXLCell? cell)
        {
            if (cell != null && cell.TryGetValue<double>(out var cellValue))
            {
                return cellValue;
            }
            else
            {
                return null;
            }
        }

        private static CircuitHeader ReadHeader(string fileName, IXLWorksheet ws)
        {
            var dbReference = ws.Cell("E7").GetString() + "-" + ws.Cell("F7").GetString();
        
            return new CircuitHeader
            (
                FileName: fileName,
                TabName: ws.Name,
                ProjectReference: ws.Cell("E5").GetString(),
                DbReference: dbReference,
                SupplyCableRef: ws.Cell("E14").GetString(),
                NumberOfWays: ReadInt(ws.Cell("E18")),
                FedFrom: ws.Cell("J14").GetString(),
                ProtectiveDeviceA: ReadDouble(ws.Cell("J16")),
                DistBoardPhase: ws.Cell("J18").GetString(),
                FaultCurrentkA: ReadDouble(ws.Cell("J20"))
            );

        }

        private static DistBoardCircuit? ReadCircuit(CircuitHeader header, IXLRow row)
        {
            var way = ReadInt(row.Cell("C"));
            var phase = row.Cell("D").GetString();
            var loadReference = row.Cell("F").GetString();
            if (loadReference == "") loadReference = row.Cell("G").GetString();

            if(way.HasValue && phase != "")
            {
                return new DistBoardCircuit
                (
                    FileName: header.FileName, 
                    TabName: header.TabName, 
                    ProjectReference: header.ProjectReference,
                    DbReference: header.DbReference,
                    SupplyCableRef: header.SupplyCableRef,
                    NumberOfWays: header.NumberOfWays,
                    FedFrom: header.FedFrom,
                    ProtectiveDeviceA: header.ProtectiveDeviceA,
                    DistBoardPhase: header.DistBoardPhase,
                    FaultCurrentkA: header.FaultCurrentkA,
                    Way: way.Value,
                    Phase: phase,
                    LoadReference: loadReference,
                    ProtectiveInA: ReadDouble(row.Cell("H")),
                    DeviceIrA: ReadDouble(row.Cell("I")),
                    RCDmA: ReadDouble(row.Cell("J")),
                    ConductorLine: ReadDouble(row.Cell("K")),
                    ConductorCPC: ReadDouble(row.Cell("L"))
                );
            } 
            else
            {
                return null;
            }
        }

    }
}

