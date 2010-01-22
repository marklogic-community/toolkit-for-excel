xquery version "1.0-ml";
(: Copyright 2009 Mark Logic Corporation

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

module namespace  excel = "http://marklogic.com/openxml/excel";

declare namespace ms    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace pr    = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip   = "xdmp:zip";

declare default element namespace  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

(: import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy"; ? :)

(:version 1.0-2:)

declare function excel:error($message as xs:string)
{
     fn:error(xs:QName("EXCEL-ERROR"),$message)
};

declare function excel:get-mimetype(
  $filename as xs:string
) as xs:string?
{
    xdmp:uri-content-type($filename)	
};

declare function excel:directory-uris(
  $directory as xs:string
) as xs:string*
{
    cts:uris("","document",cts:directory-query($directory,"infinity"))
};

declare function excel:directory-uris(
  $directory     as xs:string, 
  $includesheets as xs:boolean
) as xs:string*
{
    if($includesheets eq xs:boolean("true")) then
        cts:uris("","document",cts:directory-query($directory,"infinity"))
    else
        let $uris := cts:uris("","document",cts:directory-query($directory,"infinity"))
        let $finaluris :=  
                         for $uri in $uris
                         let $u := $uri
                         where  fn:not(fn:matches($uri, "sheet\d+\.xml$"))
                         return $u
        return $finaluris
     
};

declare function excel:sheet-uris(
  $directory as xs:string
) as xs:string*
{
    cts:uri-match(fn:concat($directory,"*sheet*.xml"))
};

declare function excel:workbook-sheet-names(
  $workbook as element(ms:workbook)
) as xs:string*
{
    $workbook/ms:sheets/ms:sheet/@name
};

declare function excel:sharedstring-uri(
  $directory as xs:string
)  
{
    cts:uri-match(fn:concat($directory,"*sharedStrings.xml"))
};

declare function excel:cell-string-value(
  $cells          as element(ms:c)*,
  $shared-strings as element(ms:sst)
) as xs:string*
{ 
    for $c in $cells
    return
      if ( $c/@t="s" ) then  (: using fn:string(.) instead of /text() to account for empty string :)
            $shared-strings/ms:si[fn:data($c/ms:v)+1]/ms:t/fn:string(.)
      else
	    $c/ms:v/text()
};

(: we have a convert function for this, but not sure we want to import. See Open XML Extract pipeline for details. :)
(: look at convert , spaces in names? check for dangerous chars mapped away:)
declare function excel:directory-to-filename(
  $directory as xs:string
) as xs:string
{
    let $after:= fn:substring($directory, 2)
    let $name := fn:substring-before($after,"_parts") 
    let $filename := fn:replace($name,"_",".")
    return $filename
};

declare function excel:xlsx-manifest(
  $directory as xs:string, 
  $uris      as xs:string*
) as element(zip:parts)
{
    <parts xmlns="xdmp:zip"> 
    {
      for $i in $uris
      let $dir := fn:substring-after($i,$directory)
      let $part :=  <part>{$dir}</part>
      return $part
    }
    </parts>
};

(: ====MAP SHARED================================================================================================= :)
(: Currently only maps SharedStrings, update for formulas? (within ws) or separate function ? :)

declare function excel:map-shared-strings(
  $sheet          as element(ms:worksheet), 
  $shared-strings as element(ms:sst)  
)as element(ms:worksheet)
{
    (: for $sheet in $sheets ?, check function mapping :)
    let $shared := fn:data($shared-strings//ms:t)
    let $rows := for $row at $d in $sheet//ms:row
                 let $cells  :=  for $cell at $e in $row/ms:c
                                 let $c := if(fn:data($cell/@t) eq "s") 
                                           then 
                                             element ms:c { $cell/@* except $cell/@t, attribute t{"inlineStr"}, element ms:is { element ms:t { $shared[($cell/ms:v+1 cast as xs:integer)] } } }  
                                           else
                                             $cell
                                                          
                                  return $c
                     
                 return element ms:row{ $row/@*, $cells }
                               
    let $worksheet   :=  $sheet/* except ( $sheet/ms:sheetData, $sheet/ms:tableParts, $sheet/ms:pageMargins, $sheet/ms:pageSetup, $sheet/ms:drawing)
    let $page-setup :=  $sheet/ms:pageSetup 
    let $table-parts :=  $sheet/ms:tableParts
    let $sheet-data  :=  $sheet/ms:sheetData
    let $drawing := $sheet/ms:drawing
    let $ws := element ms:worksheet {  $sheet/ms:worksheet/@*, $worksheet, element ms:sheetData{ $sheet-data/@*, $rows },  $page-setup ,$table-parts, $drawing  }
    return $ws

};

(: ============================================================================================================== :)
(: Simple Validation used by table generation :)

declare function excel:validate-child(
  $seq as node()*
) as xs:boolean
{
    fn:count(fn:distinct-values($seq/fn:local-name(child::*[1]))) eq 1 
};

(: ============================================================================================================== :)
(:   for future, pass function a (), determine file by parent element, assume given in right order :)

declare function excel:create-simple-xlsx(
  $worksheets as element(ms:worksheet)*
) as binary()
{
    let $ws-count := fn:count($worksheets)
    let $content-types := excel:content-types($ws-count,0)
    let $workbook := excel:workbook($ws-count)
    let $rels :=  excel:package-rels()
    let $workbookrels :=  excel:workbook-rels($ws-count)
    let $package := excel:xlsx-package($content-types, $workbook, $rels, $workbookrels, $worksheets)
    return $package

};

declare function excel:xlsx-package(
  $content-types as element(types:Types),
  $workbook      as element(ms:workbook),
  $rels          as element(pr:Relationships),
  $workbookrels  as element(pr:Relationships),
  $sheets        as element(ms:worksheet)*
) as binary()
{
   excel:xlsx-package($content-types, $workbook, $rels, $workbookrels, $sheets, (), ())
};

declare function excel:xlsx-package(
  $content-types as element(types:Types),
  $workbook      as element(ms:workbook),
  $rels          as element(pr:Relationships),
  $workbookrels  as element(pr:Relationships),
  $sheets        as element(ms:worksheet)*,
  $worksheetrels as element(pr:Relationships)*,
  $table         as element(ms:table)*
) as binary()
{

    let $return := 
    if(   (fn:empty($worksheetrels) and fn:empty($table)) or
                          fn:not(fn:empty($worksheetrels)) and fn:not(fn:empty($table))) then   
       let $manifest := <parts xmlns="xdmp:zip">
			<part>[Content_Types].xml</part>
		 	<part>xl/workbook.xml</part>
		        <part>_rels/.rels</part>
			<part>xl/_rels/workbook.xml.rels</part>
                        {
                          for $i at $d in 1 to fn:count($sheets)
                          let $sheet-name := fn:concat("xl/worksheets/sheet",$d,".xml")
			  return <part>{$sheet-name}</part>
                        }
                        { 
                          for $i at $d in 1 to fn:count($worksheetrels)
                          let $sheet-rel-name := fn:concat("xl/worksheets/_rels/sheet", $d,".xml.rels")
		   	  return <part>{$sheet-rel-name}</part>
                        }
                        {
                          for $i at $d in 1 to fn:count($table)
                          let $table-name :=  fn:concat("xl/tables/table",$d,".xml")
                          return <part>{$table-name}</part>

                        }
		     </parts>
       let $parts := ($content-types, $workbook, $rels, $workbookrels, $sheets ,$worksheetrels,$table) 
       return
         xdmp:zip-create($manifest, $parts)

    else excel:error("unable to create .xlsx package; $worksheetrels and $table must either both have values, or both be empty. You can't pass one without passing the other.")

    return $return 
};

(: a couple of ways to create-rows, these don't use excel:row constructor; 
   the creat-row functions don't specify row/@r or cell/@r 
   this is acceptable for opening in excel, order determines layout in sheet
:)
declare function excel:create-row(
  $values as xs:anyAtomicType*
) as element(ms:row)
{ 
    <ms:row>
    {
           for $val at $v in $values  
           return if($val castable as xs:integer or $val castable as xs:double) then     
                       <ms:c><ms:v>{$val}</ms:v></ms:c>
                  else
                       <ms:c t="inlineStr"> 
                           <ms:is>
                               <ms:t>{$val}</ms:t>
                           </ms:is>
                       </ms:c>
    }
    </ms:row>
};

declare function excel:create-row(
  $map as map:map, 
  $keys as xs:string*
) as element(ms:row)
{

    let $cel-vals := for $i at $d in $keys
                  let $val := map:get($map,$i)
                  let $return := if(fn:empty($val)) then "" else $val (: if empty, still create cell, string ? :)
                  return $return
    return excel:create-row($cel-vals)
};

(: dates are stored as a julian number with an @r for id which indicates display format (applies style) :)
declare function excel:cell(
  $a1-ref as xs:string, 
  $value  as xs:anyAtomicType?
) as element(ms:c) 
{
    excel:cell($a1-ref,$value,(),())              
};

declare function excel:cell(
  $a1-ref  as xs:string, 
  $value   as xs:anyAtomicType?, 
  $formula as xs:string?
) as element(ms:c)
{
    excel:cell($a1-ref,$value,$formula,())
};

declare function excel:cell(
  $a1-ref  as xs:string, 
  $value   as xs:anyAtomicType?, 
  $formula as xs:string?,
  $date-id as xs:integer?
) as element(ms:c)
{
    if($value castable as xs:integer or fn:empty($value)) then     
        let $formula := if(fn:not(fn:empty($formula))) then 
                            <ms:f>{$formula}</ms:f> 
                        else ()
        let $value := if(fn:not($value eq 0) and fn:not(fn:empty($value)))then
                       <ms:v>{$value}</ms:v>
                      else ()
        return    
          element ms:c{ attribute r{$a1-ref}, 
                        if(fn:empty($date-id)) then () 
                        else attribute s{$date-id},
                        $formula, $value }
    else
              <ms:c r={$a1-ref} t="inlineStr"> 
                    <ms:is>
                        <ms:t>{$value}</ms:t>
                    </ms:is>
              </ms:c>
};

declare function excel:row(
  $cells as element(ms:c)+
) as element(ms:row)
{
    let $ids := fn:count(fn:distinct-values(excel:a1-row($cells/@r)))
    let $return := if($ids = 1) then
                                  let $ordcells := for $i in $cells 
                                                   order by excel:col-letter-to-idx(excel:a1-column($i/@r)) ascending
                                                   return $i
                                  return
                                   <ms:row r={excel:a1-row($cells[1]/@r)}>{$ordcells}</ms:row>
                   else 
                      excel:error("All cells are not in the same row.  Unable to create row.")
    return $return
};

declare function excel:col-letter-to-idx(
  $letter as xs:string
) as xs:integer
{
 let $result := if(fn:string-length($letter) = 1) then 
                      fn:string-to-codepoints($letter) - 64
                   else if(fn:string-length($letter) =2) then 
                      let $ref := (fn:string-to-codepoints($letter)- 64)
                      let $Nref := ($ref[1] * 26) + $ref[2] 
                      return $Nref
                   else let $ref := (fn:string-to-codepoints($letter)- 64)
                      let $Nref := ($ref[1] * 702)
                      let $NNref :=   $Nref+ (($ref[2] * 26) + $ref[3])
                      return $NNref
 return $result
};

declare function excel:a1-to-r1c1(
  $a1notation as xs:string
) as xs:string
{
    let $col-index := excel:col-letter-to-idx(excel:a1-column($a1notation))
    let $row-index := excel:a1-row($a1notation)
    let $return := if(($row-index gt 1048756) or ($col-index gt 16384)) then
                       excel:error("The row and/or column index is beyond the limits of what Excel allows.")
                   else fn:concat("R",$row-index,"C",$col-index)
    return $return
};

declare function excel:r1c1-to-a1(
  $row-index as xs:integer, 
  $col-index as xs:integer
) as xs:string 
{

(:excel has limits of 16384 for columns and 1048756 for rows :)
(:these are limits for 2007 only, increased from 2003 :)
(:if out of possible range, return error :)

 let $return := if(($row-index gt 1048756) or ($col-index gt 16384)) then
                      excel:error("The row and/or column index is beyond the limits of what Excel allows.")
                else 

    (:do a simple check for row/index  or simple type:) 
    (: the first columns, A-Z in excel, can be referred to numerically as 1-26 :)
    (: columns in excel, progress from A-Z, to AA-ZZ, to AAA-WID :)
    (: AA-ZZ has 676 possible combinations :)
    (: so 26+676=702; gt 702, we know the column has 3 letters :)
    (: first-letter here checks for a possible 3 letter column reference :)
    let $first-letter :=   
          if($col-index >= 703) then (:idiv mod :)
                let $newcol := fn:floor($col-index div 702)
                let $flcheck := $col-index - fn:floor($newcol* 702)
                let $l := $col-index - ($newcol* 702)
                let $first-letter := 
                         if($flcheck = $l and $flcheck mod 702 = 0) then 
                             fn:codepoints-to-string($newcol+63)
                         else
                             fn:codepoints-to-string($newcol+64)
                return $first-letter
          else ""
                   
    (:now that we have the first column, need to take the delta to calculate the index for 
      double/single-letter column references :)
    let $ucol := 
          if($col-index >= 703) then
                let $delta := $col-index mod 702 
                let $ncol := 
                         if($delta <= 26 )then 
                             $delta + 26 
                         else $delta   
                return $ncol
          else 
                $col-index      

    let $coldiv := fn:floor($ucol div 26 )
    let $letter := 
          (:obscure case, I'm probably off by 1 somewhere, will check later :)
          if($ucol <= 26 and $col-index > 703) then 
                "ZZ" (: off by one somewhere :)
          (: check for single column :)
          else if($ucol <= 26) then                  
                fn:codepoints-to-string($ucol+64)
          (:check for two-letter column reference :)
          else
                let $coldiv := fn:floor($ucol div 26)
                let $coldiv2 := $ucol div 26
                let $first-letter := 
                         if($coldiv2 eq $coldiv )then 
                             fn:codepoints-to-string($coldiv+63)  
                         else fn:codepoints-to-string($coldiv+64)

                let $next-letter-check := $ucol - fn:floor($coldiv2 * 26)
                let $next-letter := $ucol - ($coldiv * 26)
                let $final := 
                         if($next-letter-check = $next-letter and $next-letter-check mod 26 = 0) then 
                             fn:codepoints-to-string($next-letter-check +90)   
                         else fn:codepoints-to-string($next-letter+64) 
                return fn:concat($first-letter,$final)  

    (:concat the column letters with the row index for A1 notation :)
    return fn:concat($first-letter,$letter,$row-index)  
    
 return $return 
};

declare function excel:column-width(
  $widths as xs:integer+
) as element(ms:cols)
{
    <ms:cols>
    { 
        for $i at $d in $widths
        return <ms:col min="{$d}" max="{$d}" width="{$i}" customWidth="1"/>
    }
    </ms:cols>
};

(:option tbl-count parameter:)
declare function excel:content-types(
  $worksheet-count as xs:integer
) as element(types:Types)
{
    let $content-types := 
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        {
           for $i in 1 to $worksheet-count
           let $sheet-name := fn:concat("/xl/worksheets/sheet", $i )
           return
	     <Override PartName={$sheet-name} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        }
       </Types>
    return $content-types
};

declare function excel:content-types(
  $worksheet-count as xs:integer,
  $tbl-count as xs:integer
) as element(types:Types)
{
    let $content-types := 
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        {
           for $i in 1 to $worksheet-count
           let $sheet-name := fn:concat("/xl/worksheets/sheet", $i )
           return
	     <Override PartName={$sheet-name} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        }
        {
            for $j in 1 to $tbl-count
            let $table-name :=  fn:concat("/xl/tables/table", $j,".xml" )
            return
                <Override PartName={$table-name} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
        }
       </Types>
    return $content-types
};

declare function excel:workbook(
  $worksheet-count as xs:integer
) as element(ms:workbook)
{
    let $workbook := 
       <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
        {
          for $i at $d in 1 to $worksheet-count
          let $sheet-name := fn:concat("Sheet", $d )
          let $rId := fn:concat("rId",$d)
          return <sheet name={$sheet-name} sheetId={$d} r:id={$rId} />
        }
        </sheets>
       </workbook>
    return $workbook
};

declare function excel:package-rels() as element(pr:Relationships)
{
    let $rels :=
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
       </Relationships>
    return $rels
};

declare function excel:workbook-rels(
  $worksheet-count as xs:integer
) as element(pr:Relationships)
{
    let $workbookrels :=
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        {
          for $i at $d in 1 to $worksheet-count (: d redundant, STAMP OUT LET!! :)
          let $target := fn:concat("worksheets/sheet", $d,".xml")
          let $rId := fn:concat("rId",$d) 
	  return <Relationship Id={$rId} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target={$target}/>
        }
       </Relationships>
    return $workbookrels
};

(: will manifest as sheet1.xml.rels, sheet2.xml.rels, ... sheetN.xml.rels in .xlsx pkg:)
declare function excel:worksheet-rels(
  $start-ind as xs:integer,
  $tbl-count as xs:integer
) as element(pr:Relationships)
{
    let $worksheetrels:=
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        {
          for $i in 1 to $tbl-count
          let $target := fn:concat("../tables/table",($start-ind + $i - 1),".xml")
          let $id := fn:concat("rId",$i)
          return
            <Relationship Id={$id} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target={$target}/>
        }
       </Relationships> 
    return $worksheetrels
};

(: next go round pass in worksheet too; higher-level functions could simplify table mgmt :)
declare function excel:table(
  $table-number as xs:integer,
  $tablerange   as xs:string, 
  $column-names as xs:string+
) as element(ms:table)
{
    excel:table($table-number,$tablerange,$column-names,(),())
};

declare function excel:table(
  $table-number as xs:integer,
  $tablerange   as xs:string, 
  $column-names as xs:string+,
  $auto-filter  as xs:boolean? 
) as element(ms:table)
{
    excel:table($table-number,$tablerange,$column-names,$auto-filter,())
};

declare function excel:table(
  $table-number as xs:integer,
  $tablerange   as xs:string, 
  $column-names as xs:string+,
  $auto-filter  as xs:boolean?, 
  $style        as xs:boolean?
) as element(ms:table)
{
(: todo: verify range compatible with number of columns :)
(: todo: check disp-name; manifests as selectable named range in excel; @name might be required to map to table1.xml, etc. where @displayName differs for label in Excel? :)

    let $disp-name := fn:concat("Table",$table-number)
    let $id := $table-number

    let $column-count := fn:count($column-names)
    let $table :=
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id={$id} name={$disp-name} displayName={$disp-name} ref="{$tablerange}"  totalsRowShown="0" >
         {
             if(fn:empty($auto-filter) or $auto-filter) then
             <autoFilter ref="{$tablerange}"/>
             else 
              () 
         }
             <tableColumns count={$column-count}> 
         {
               for $i at $d in $column-names
               return <tableColumn id={$d} name={$i}/>
         }
             </tableColumns>
         {
         if($style) then 
             <tableStyleInfo name="TableStyleMedium10" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
         else ()
         }
      </table>
    return $table
};

(: optional col-widths, tbl-count parameters :)
declare function excel:worksheet(
  $rows as element(ms:row)*
) as element(ms:worksheet)
{
    excel:worksheet($rows,(),())
};

declare function excel:worksheet(
  $rows      as element(ms:row)*,
  $colwidths as element(ms:cols)*
) as element(ms:worksheet)
{
    excel:worksheet($rows,$colwidths,())
};

declare function excel:worksheet(
  $rows      as element(ms:row)*,
  $colwidths as element(ms:cols)?,
  $tbl-count as xs:integer?
) as element(ms:worksheet)
{
    let $sheet := <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      {
                       if(fn:not(fn:empty($colwidths))) then
                           $colwidths
                       else
                           ()
                      } 
                      <sheetData>
                      { 
                        ($rows) 
                      }
                      </sheetData>
                      {
                        if(fn:not(fn:empty($tbl-count)) and $tbl-count  gt 0) then
                          <tableParts count={$tbl-count}>
                           {
                            for $i in 1 to $tbl-count
                            let $id := fn:concat("rId",$i)
                            return 
                                <tablePart r:id={$id} />
                           }
                          </tableParts>
                        else () 
                      }
                  </worksheet>
    return $sheet
};

(:utility functions to grab Column or Row from "A1" notation :)
declare function excel:a1-row(
  $a1 as xs:string
) as xs:integer
{
      xs:integer(fn:replace($a1,("[A-Z]+"),""))
};

declare function excel:a1-column(
  $a1 as xs:string 
)as xs:string
{
       fn:replace($a1,("\d+"),"")
};

(:     functions related set-cells       :)
declare function excel:passthru-workbook($x as node(), $newSheetData as element(ms:sheetData)) as node()*
{
   for $i in $x/node() return excel:wb-set-sheetdata($i,$newSheetData)
};

declare function excel:wb-set-sheetdata($x as node(), $newSheetData as element(ms:sheetData)) as node()*
{
 
   typeswitch($x)
     case text() return $x
     case document-node() return document {$x/@*,excel:passthru-workbook($x,$newSheetData)}
     case element(ms:sheetData) return  $newSheetData
     case element() return  element{fn:name($x)} {$x/@*,excel:passthru-workbook($x,$newSheetData)}
     default return $x

};

declare function excel:set-cells(
  $v_sheet as element(ms:worksheet), 
  $cells as element(ms:c)*
) as element(ms:worksheet)
{   
    let $sheetDataTst := $v_sheet//ms:sheetData
    let $sheet := if(fn:empty($sheetDataTst/*)) then 
                        let $row := excel:row(excel:cell("A1",""))
                        let $tmpSheet := excel:worksheet($row)
                        return excel:wb-set-sheetdata($v_sheet, $tmpSheet//ms:sheetData) 
                 else $v_sheet

    let $sheetData := $sheet/ms:sheetData

    let $cel-rows := fn:distinct-values(excel:a1-row($cells/@r))
    let $orig-rows := $sheetData/ms:row
   
    (:determine missng rows, stub out rows that don't exist so we can loop through rows and add cells accordingly :) 
    let $mssing := 
                    for $i in $cel-rows
                    let $x := if(fn:exists($orig-rows[@r=$i])) then () else 
                              let $a1 := fn:concat("A",$i)
                              return excel:row(excel:cell($a1,""))
                    return $x

    let $both := (($orig-rows,$mssing))
   
    (:want to keep rows in order :) 
    let $ord-rows := for $i in $both
                     order by xs:integer($i/@r) ascending  
                     return $i
                               
    let $new-rows :=  for $r in $ord-rows
                      let $cells-to-add := for $i in $cells
                                           let $c := $i where (excel:a1-row($i/@r) = $r/@r)
                                           return $c
                      let $orig-cells := $r/ms:c

                      (:now must replace any cells that previous exist:)
                      (: want to eliminate cells that we're replacing, so update cells to only return delta, then merge the two sequences:)

                      let $upd-cells := for $i in $orig-cells
                                        let $x := if(fn:exists($cells-to-add[@r=$i/@r])) 
                                        then () else $i
                                        return $x
                      
                      (:now that we have all cells, order, return row:)
                      let $all-cells2 := ($cells-to-add, $upd-cells)
                      let $all-ord-cells := for $ce in $all-cells2
                                            order by excel:col-letter-to-idx(excel:a1-column($ce/@r)) ascending
                                            return $ce


                      return  element ms:row{ $r/@* , ($r/* except $r/ms:c), $all-ord-cells}

    (:set the updated rows in sheetData, set sheetData in worksheet, return worksheet:)
    let $newSheetData := element ms:sheetData { $sheetData/@*, $new-rows }
    let $final-sheet := excel:wb-set-sheetdata($sheet, $newSheetData)
    return $final-sheet
                    
};

(: calendar conversion functions   :)
declare function excel:julian-to-gregorian(
  $excel-julian-day as xs:integer
) as xs:date
{
   (: formula from http://quasar.as.utexas.edu/BillInfo/JulianDatesG.html :)
   (: adapted for excel :)
   (: won't calculate for years < 400 :)
   let $JD :=  $excel-julian-day - 2 +2415020.5 
   let $Z :=  $JD+0.5
   let $W := fn:floor(($Z - 1867216.25) div 36524.25)
   let $X := fn:floor($W div 4)
   let $A := $Z + 1 + $W - $X
   let $B := $A+1524
   let $C := fn:floor(( $B - 122.1) div 365.25)
   let $D := fn:floor(365.25 * $C)
   let $E := fn:floor(($B - $D) div 30.6001)
   let $F := fn:floor(30.6001 * $E)
   let $day  := $B - $D - $F
   let $month := if($E < 13.5) then ($E - 1) else ($E - 13)
   let $year := if($E <=2) then $C - 4715 else $C - 4716
   let $finday := if(fn:string-length($day cast as xs:string) eq 1) then fn:concat("0",$day) else $day
   let $finmonth := if(fn:string-length($month cast as xs:string) eq 1) then fn:concat("0",$month) else $month
   let $findate := fn:concat($year,"-", $finmonth, "-",$finday)

   (: return  ($day, $month, $year) :)
(: return as xs:date :)
   return   xs:date($findate)
};

(:function that's inverse of above :)
declare function excel:gregorian-to-julian(
  $date   as xs:date
) as xs:integer
{
   let $year := fn:year-from-date($date)
   let $month := fn:month-from-date($date)
   let $day := fn:day-from-date($date)
   (: formula from http://quasar.as.utexas.edu/BillInfo/JulianDatesG.html :)
   (: adapted for excel :)
   let $NY := if($year <= 2) then $year + 1 else $year
   let $NM := if($year<=2) then $month + 12 else $month
   let $A := fn:floor($NY div 100)
   let $B :=  fn:floor($A div 4)
   let $C := (2 - $A + $B)
   let $E := fn:floor(365.25 * ( $NY + 4716))
   let $F := fn:floor(30.6001 * ($NM + 1))
   let $NJD := $C + $day + $E + $F - 1524.5 - 2415020.5 + 2 
   return $NJD
};

(:rename width; optional colcustwidth and tabstyle params :)

declare function excel:create-xlsx-from-xml-table(
$originalxml as node()
) as binary()?
{
    create-xlsx-from-xml-table($originalxml,(),())
};

declare function excel:create-xlsx-from-xml-table(
$originalxml as node(),
$colcustwidth as xs:string?
) as binary()?
{
    create-xlsx-from-xml-table($originalxml,$colcustwidth,()) 
};

declare function excel:create-xlsx-from-xml-table(
$originalxml as node(),
$colcustwidth as xs:string?,
$auto-filter as xs:boolean?
) as binary()?
{
    create-xlsx-from-xml-table($originalxml,$colcustwidth, $auto-filter, ()) 
};


declare function excel:create-xlsx-from-xml-table(
$originalxml as node(),
$colcustwidth as xs:string?,
$auto-filter as xs:boolean?,
$tabstyle as xs:boolean?
) as binary()?
{
   let $worksheetrows := $originalxml/child::*

(: order of rows?  :)
(: remove at clauses :)
   let $allrows := for $r in $worksheetrows
             let $rowhdrs := $r//child::*
             let $validrownames :=
                    for $i in $rowhdrs 
                    let $rowhdrname := fn:local-name($i)
                    return $rowhdrname
             return $validrownames
   let $headerrows :=  fn:distinct-values($allrows) (:could wrap around flower above :)
   let $columncount := fn:count($headerrows)

   let $headers := excel:create-row($headerrows)

   let $rowvalues := for $i in $worksheetrows
                  let $map := map:map()
                  let $return := for $x at $z in $i/child::*
                                 let $put := map:put($map, fn:local-name($x),$x/text())
                                 return $put 
                  return $map

   let $rows := for $i in $rowvalues
             return excel:create-row($i,$headerrows) 

   let $rowcount := fn:count($rows)
            
   let $content-types := excel:content-types(1,1)
   let $workbook := excel:workbook(1)
   let $rels :=  excel:package-rels()
   let $workbookrels :=  excel:workbook-rels(1)

   let $tablerange := fn:concat("A1:",excel:r1c1-to-a1($rowcount+1,$columncount))

   let $tablexml :=  if($tabstyle or $auto-filter) then excel:table(1,$tablerange, $headerrows, $auto-filter, $tabstyle) 
                     else ()
                      

   let $worksheetrels := if($tabstyle or $auto-filter) then  excel:worksheet-rels(1,1)
                         else ()

   let $sheet-col-widths := if(fn:not(fn:empty($colcustwidth))) then
                                 for $i in 1 to $columncount
                                 return $colcustwidth cast as xs:integer 
                            else
                                 ()
   let $colwidths := if(fn:empty($sheet-col-widths)) then ()
                     else excel:column-width($sheet-col-widths) 

   let $tbl-count := if($tabstyle or $auto-filter) then 1 else ()
   let $sheet1 := excel:worksheet(($headers,$rows), $colwidths, $tbl-count) 

   let $package := excel:xlsx-package($content-types, $workbook, $rels, $workbookrels, $sheet1, $worksheetrels, $tablexml)
   return $package
};
