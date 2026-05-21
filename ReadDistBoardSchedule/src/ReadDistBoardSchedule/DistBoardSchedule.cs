// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.



using System.Text.Json;

namespace SptUtils.ReadDistBoardSchedule 
{
    public record DistBoardSchedule
    (
        string FileName, 
        string TabName,
        string ProjectReference,
        string DbReference,
        string SupplyCableRef,
        int? NumberOfWays,
        string FedFrom,
        double? ProtectiveDeviceA,
        string Phase,
        double? FaultCurrentkA,
        List<DistBoardCircuit> Circuits
    )
    {
        public void WriteJson(Utf8JsonWriter writer)
        {
            writer.WriteStartObject();
            writer.WriteString("file_name", FileName);
            writer.WriteString("sheet_name", TabName);
            writer.WriteString("project_reference", ProjectReference);
            writer.WriteString("db_reference", DbReference);
            writer.WriteString("supply_cable_reference", SupplyCableRef);
            if (NumberOfWays.HasValue) writer.WriteNumber("number_of_ways", NumberOfWays.Value); else writer.WriteNull("number_of_ways");
            if (ProtectiveDeviceA.HasValue) writer.WriteNumber("protective_device_a", ProtectiveDeviceA.Value); else writer.WriteNull("protective_device_a");
            writer.WriteString("phase", Phase);
            if (FaultCurrentkA.HasValue) writer.WriteNumber("fault_current_ka", FaultCurrentkA.Value); else writer.WriteNull("fault_current_ka");
            
            // TODO do we want nesting or denormalization?    
            // writer.WritePropertyName("circuits");
            // writer.WriteStartArray();
            // foreach (var item in Circuits)
            // {
            //     item.WriteJson(writer);
            // }
            // writer.WriteEndArray();
            writer.WriteEndObject();
        }
    };
    
}
