xquery version "1.0-ml";
(:
Copyright 2008-2010 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
:)

declare namespace w    ="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
declare namespace pkg="http://schemas.microsoft.com/office/2006/xmlPackage";
import module namespace ooxml="http://marklogic.com/openxml" at "/MarkLogic/openxml/word-processing-ml-support.xqy";

(:Check for if folder or ends in .xml, currently assume need folder of .docx extracted:)
let $q := xdmp:get-request-field("qry")
return xdmp:quote(ooxml:get-directory-package($q))
