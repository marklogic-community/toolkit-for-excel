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
 * TKEvents.cs - event handling.
 * Events caught and signals sent to functions in MarkLogicPowerPointEventSupport.js
 * 
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MarkLogic_PowerPointAddin
{
    [Serializable]
    public class PictureFormatView
    {
        public PictureFormatView() { }

        public string brightness { get; set; }
        public string colorType { get; set; }
        public string contrast { get; set; }
        public string cropBottom { get; set; }
        public string cropLeft { get; set; }
        public string cropTop { get; set; }
        public string cropRight { get; set; }
        public string transparencyColor { get; set; }
        public string transparencyBackground { get; set; }
    }


	       
        
}
