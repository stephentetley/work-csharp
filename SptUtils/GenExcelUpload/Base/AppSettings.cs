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


namespace SptUtils.GenExcelUpload.Base 
{

    public class AppSettings
    {
        public required string TemplateFolder { get; set; }
        public required string FlocCreateTemplate { get; set; }
        public required string FlocChangeTemplate { get; set; }
        public required string EquiCreateTemplate { get; set; }
        public required string EquiChangeTemplate { get; set; }

        public string FlocCreateTemplatePath => Path.Combine(TemplateFolder, FlocCreateTemplate);
        public string FlocChangeTemplatePath => Path.Combine(TemplateFolder, FlocChangeTemplate);
        public string EquiCreateTemplatePath => Path.Combine(TemplateFolder, EquiCreateTemplate);
        public string EquiChangeTemplatePath => Path.Combine(TemplateFolder, EquiChangeTemplate);
    }

}
