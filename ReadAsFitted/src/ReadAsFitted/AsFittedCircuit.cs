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

namespace SptUtils.ReadAsFitted 
{
    public record AsFittedCircuit
    (
        string FileName, 
        string TabName,
        string SiteName,
        string DbOrPanelNumber,
        string TestDate,
        string SheetNumber,
        string AibRef,
        string Location,
        string CableNum,
        string FedFrom,
        string CircuitRefAndPhase,
        string CircuitDescription,
        string CircuitType
    )
    {
        public void WriteJson(Utf8JsonWriter writer)
        {
            writer.WriteStartObject();
            writer.WriteString("file_name", FileName);
            writer.WriteString("sheet_name", TabName);
            writer.WriteEndObject();
        }
    };
    
}
