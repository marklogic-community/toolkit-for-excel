xquery version "1.0-ml";
import module namespace excel= "http://marklogic.com/openxml/excel" 
       at "/MarkLogic/openxml/spreadsheet-ml-support.xqy";

declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";

xdmp:save("C:\test-output.xml",
<et:tests xmlns:et="http://marklogic.com/openxml/excel/tests">
<et:test num="1">
   {excel:a1-column("D23")}
</et:test>
<et:test num="2">
   {excel:a1-row("D23")}
</et:test>
<et:test num="3" id="1">
   {excel:cell("A1",123) }
</et:test>
<et:test num="3" id="2">
   {excel:cell("A3",(),"SUM(A1:A2)")}
</et:test>
<et:test num="3" id="3">
   {excel:cell("A3",12,"SUM(A1:A2)")}
</et:test>
<et:test num="3" id="4">
   {excel:cell("A3",39999,(),0)}
</et:test>
<et:test num="3" id="5">
   {excel:cell("A3","FOO",(),())}
</et:test>
<et:test num="4">
   {
   let $ss := <ms:sst>
                <ms:si><ms:t>Name</ms:t></ms:si>
                <ms:si><ms:t>Description</ms:t></ms:si>
              </ms:sst>
   let $cels := (<ms:c>
                   <ms:v>1</ms:v>
                 </ms:c>,
                 <ms:c t="s">
                   <ms:v>1</ms:v>
                 </ms:c>)
   return excel:cell-string-value ($cels,$ss)
   }
</et:test>
<et:test num="5">
   {excel:column-width((15,25))}
</et:test>
<et:test num="6">
   {excel:content-types(3)}
</et:test>
<et:test num="7">
   {
   let $vals :=(1,2,3,"TEST")
   return excel:create-row($vals)
   }
</et:test>
<et:test num="8">
{
let $map := map:map()
let $put := (map:put($map, "RequestID",45683),
             map:put($map, "Customer","Oslo"))

let $keys := ("Customer","Address","RequestID")
return excel:create-row($map,$keys)
}
</et:test>
<et:test num="9">
   {
   let $ws:= (<ms:worksheet>
                <ms:sheetData>
                 <ms:row>
                   <ms:c t="s">
                   <ms:v>1</ms:v>
                 </ms:c>
                </ms:row>
               </ms:sheetData>
              </ms:worksheet>)
   return xdmp:zip-manifest(excel:create-simple-xlsx($ws))
   }
</et:test>
<et:test num="10" id="1">
   {
   let $xml := <catalog>
                <item>
                 <product>beach ball</product>
                 <sku>123123</sku>
                </item>
                <item>
                 <product>swim fins</product>
                 <sku>444444</sku>
                </item>
                <item>
                 <product>scuba glasses</product>
                 <sku>888</sku>
                </item>
               </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,"15",fn:true())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"),xdmp:zip-get($package,"xl/tables/table1.xml"))
}
</et:test>
<et:test num="10" id="2">
   {
   let $xml := <catalog>
                <item>
                 <product>beach ball</product>
                 <sku>123123</sku>
                </item>
                <item>
                 <product>swim fins</product>
                 <sku>444444</sku>
                </item>
                <item>
                 <product>scuba glasses</product>
                 <sku>888</sku>
                </item>
               </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,(),fn:true())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"),xdmp:zip-get($package,"xl/tables/table1.xml"))
   }
</et:test>
<et:test num="10" id="3">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,(),())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"))
   }
</et:test>
<et:test num="10" id="4">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml)
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"))
   }
</et:test>
<et:test num="10" id="5">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,"15",fn:false())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"))
   }
</et:test>
<et:test num="10" id="6">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,"15",fn:false(),fn:false())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"))
   }
</et:test>
<et:test num="10" id="7">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,"15",fn:true(),fn:false())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"),xdmp:zip-get($package,"xl/tables/table1.xml"))
   }
</et:test>
<et:test num="10" id="8">
   {
   let $xml := <catalog>
                 <item>
                    <product>beach ball</product>
                    <sku>123123</sku>
                 </item>
                 <item>
                    <product>swim fins</product>
                    <sku>444444</sku>
                 </item>
                 <item>
                    <product>scuba glasses</product>
                    <sku>888</sku>
                 </item>
                </catalog>
   let $package := excel:create-xlsx-from-xml-table($xml,"15",fn:false(),fn:true())
   return (xdmp:zip-manifest($package),xdmp:zip-get($package,"xl/worksheets/sheet1.xml"),xdmp:zip-get($package,"xl/tables/table1.xml"))
   }
</et:test>
<et:test num="11">
   { excel:directory-to-filename("/Default_xlsx_parts/")}
</et:test>
<et:test num="12">
   { excel:directory-uris("/Default_xlsx_parts/")}
</et:test>
<et:test num="13">
   { excel:get-mimetype("Default.xlsx")}
</et:test>
<et:test num="14">
   { excel:gregorian-to-julian(xs:date("2009-04-06"))}
</et:test>
<et:test num="15">
   { excel:julian-to-gregorian(39909)}
</et:test>
<et:test num="16">
   {
   let $ss := fn:doc(excel:sharedstring-uri("/Default_xlsx_parts/"))/node()

   let $ws:= (<ms:worksheet>
                <ms:sheetData>
                 <ms:row>
                   <ms:c t="s">
                   <ms:v>1</ms:v>
                 </ms:c>
                </ms:row>
               </ms:sheetData>
              </ms:worksheet>)

   return excel:map-shared-strings($ws,$ss)
   }
</et:test>
<et:test num="17">
   { excel:package-rels()}
</et:test>
<et:test num="18">
   { excel:r1c1-to-a1(1,2378)}
</et:test>
<et:test num="19">
   { excel:a1-to-r1c1("CJL1")}
</et:test>
<et:test num="20">
   { 
   let $cell1 := excel:cell("A3",32999,(),0)
   let $cell2 := excel:cell("B3",123)
   let $cell3 := excel:cell("C3","Foo")
   return excel:row(($cell1,$cell2,$cell3))
   }
</et:test>
<et:test num="21" id="1">
   { 
   let $cell1 := excel:cell("A1","foo") 
   let $cell2 := excel:cell("B3",123)
   let $cell3 := excel:cell("A5",456)
   let $cell4 := excel:cell("AA5",456)

   let $worksheet :=
   <ms:worksheet>	
      <ms:sheetData>
         <ms:row r="1">
            <ms:c r="A1"><ms:v>1</ms:v></ms:c>
         </ms:row>
         <ms:row r="5">
            <ms:c r="C5"><ms:v>1</ms:v></ms:c>
         </ms:row>
       </ms:sheetData>
   </ms:worksheet>
   return excel:set-cells($worksheet, ($cell1,$cell2,$cell3,$cell4))
   }
</et:test>
<et:test num="21" id="2">
   {
   let $cel1 := excel:cell("G23","bar")
   let $worksheet :=
   <ms:worksheet>	
      <ms:sheetData>
      </ms:sheetData>
   </ms:worksheet>
   return excel:set-cells($worksheet, ($cel1))
   }
</et:test>
<et:test num="22">
   { excel:sharedstring-uri("/Default_xlsx_parts/")}
</et:test>
<et:test num="23">
   { excel:sheet-uris("/Default_xlsx_parts/")}
</et:test>
<et:test num="24" id="1">
   { excel:table(1,"A1:C3",("Heading1","Heading2","Heading3"))}
</et:test>
<et:test num="24" id="2">
   { excel:table(1,"A1:C3",("Heading1","Heading2","Heading3"),fn:false(),fn:true())}
</et:test>
<et:test num="25">
   { excel:workbook(3)}
</et:test>
<et:test num="26">
   { excel:workbook-rels(3)}
</et:test>
<et:test num="27">
   {
   let $workbook := fn:doc("/Default_xlsx_parts/xl/workbook.xml")/node()
   return
     excel:workbook-sheet-names($workbook)
   }
</et:test>
<et:test num="28" id="1">
   {
   let $cells := ((excel:cell("A1",1), 
                   excel:cell("B1",2), 
                   excel:cell("C1",3)))
   let $row := excel:row($cells)
   return excel:worksheet($row)
   }
</et:test>
<et:test num="28" id="2">
   {
   let $cells := ((excel:cell("A1",1), 
                   excel:cell("B1",2), 
                   excel:cell("C1",3)))
   let $row := excel:row($cells)
   let $colwidths := excel:column-width((25,25,25))
   return excel:worksheet($row,$colwidths)
   }
</et:test>
<et:test num="28" id="3">
   {
   let $cells := ((excel:cell("A1",1), 
                   excel:cell("B1",2), 
                   excel:cell("C1",3)))
   let $row := excel:row($cells)
   let $colwidths := excel:column-width((25,25,25))
   return excel:worksheet($row,$colwidths,2)
   }
</et:test>
<et:test num="29">
   { excel:worksheet-rels(2,2)}
</et:test>
<et:test num="30">
   {
   let $worksheets:= (<ms:worksheet>
                    <ms:sheetData>
                     <ms:row>
                       <ms:c>
                        <ms:v>1</ms:v>
                       </ms:c>
                     </ms:row>
                    </ms:sheetData>
                  </ms:worksheet>)

   let $ws-count := fn:count($worksheets)
   let $content-types := excel:content-types($ws-count,0)
   let $workbook := excel:workbook($ws-count)
   let $rels :=  excel:package-rels()
   let $workbookrels :=  excel:workbook-rels($ws-count)
   let $package := excel:xlsx-package($content-types, $workbook, $rels, $workbookrels, $worksheets)

   return xdmp:zip-manifest($package)

   }
</et:test>
<et:test num="31">
{
let $uris := ("/Default_xlsx_parts/xl/workbook.xml",
              "/Default_xlsx_parts/xl/worksheets/Sheet1.xml",
              "/Default_xlsx_parts/[Content_Types].xml",
              "/Default_xlsx_parts/_rels/.rels")

return excel:xlsx-manifest("/Default_xlsx_parts/",$uris)
}
</et:test>
</et:tests>
)
