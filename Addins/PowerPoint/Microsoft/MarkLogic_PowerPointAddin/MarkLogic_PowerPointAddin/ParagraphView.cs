/*Copyright 2009-2011 MarkLogic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 * 
 * ParagraphView.cs - used by addShape(). 
 * JSON paragraphs deserialized as object of this class.
 * 
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarkLogic_PowerPointAddin
{
    [Serializable]
    public class ParagraphView
    {
        public ParagraphView() { }
        public List<string> paragraphAlignment { get; set; }
        public List<string> paragraphBulletType { get; set; }
        public List<string[]> runs { get; set; }
        // public List<string> runFontSize { get; set; }
        // public List<string> runFontRGB { get; set; }
        // public List<string> runFontName { get; set; }
        // public List<string[]> values { get; set; }
    }
}
