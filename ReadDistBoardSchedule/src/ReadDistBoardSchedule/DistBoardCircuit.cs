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
    public record DistBoardCircuit
    (
        int Way, 
        string Phase,
        double? ProtectiveInA,
        double? DeviceIrA,
        double? RCDmA,
        double? ConductorLine,
        double? ConductorCPC
    )
    {
        public void WriteJson(Utf8JsonWriter writer)
        {
            writer.WriteStartObject();
            writer.WriteNumber("way", Way);
            writer.WriteString("phase", Phase);
            if (ProtectiveInA.HasValue) writer.WriteNumber("protective_in_a", ProtectiveInA.Value); else writer.WriteNull("protective_in_a");
            if (DeviceIrA.HasValue) writer.WriteNumber("device_ir_a", DeviceIrA.Value); else writer.WriteNull("device_ir_a");
            if (RCDmA.HasValue) writer.WriteNumber("rcd_ma", RCDmA.Value); else writer.WriteNull("rcd_ma");
            if (ConductorLine.HasValue) writer.WriteNumber("conductor_line", ConductorLine.Value); else writer.WriteNull("conductor_line");
            if (ConductorCPC.HasValue) writer.WriteNumber("conductor_cpc", ConductorCPC.Value); else writer.WriteNull("conductor_cpc");
            writer.WriteEndObject();
        }
    };
    
}
