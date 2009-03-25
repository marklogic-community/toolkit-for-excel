xquery version "1.0-ml";
(: Copyright 2008 Mark Logic Corporation

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
module namespace excel = "http://marklogic.com/openxml/excel";
declare namespace ms="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
declare namespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
declare namespace zip = "xdmp:zip";

declare default element namespace  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";


(: import module "http://marklogic.com/openxml" at "/MarkLogic/openxml/package.xqy"; :)
declare function excel:get-mimetype(
  $filename as xs:string
) as xs:string?
{
           xdmp:uri-content-type($filename)	
};

(: mimetypes.xml - xdmp:get-mimetype ?:)
(:  xdmp:uri-content-type() :)
(:doesn't exist in other package :)
(:declare function excel:get-mimetype(
  $filename as xs:string
) as xs:string?
{
	let $mime-type:= 
        if(fn:contains($filename,"docx")) then 
               	"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else if(fn:contains($filename,"pptx")) then
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        else if(fn:contains($filename,"xlsx")) then
               	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else if(fn:contains($filename,"docm")) then
               	"application/vnd.ms-word.document.macroEnabled.12"
        else if(fn:contains($filename,"dotm")) then
                "application/vnd.ms-word.template.macroEnabled.12"
        else if(fn:contains($filename,"dotx")) then
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
        else if(fn:contains($filename,"ppsm")) then
               "application/vnd.ms-powerpoint.slideshow.macroEnabled.12"
        else if(fn:contains($filename,"ppsx")) then
              "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
        else if(fn:contains($filename,"pptm")) then
              "application/vnd.ms-powerpoint.presentation.macroEnabled.12"
        else if(fn:contains($filename,"xlsb")) then
              "application/vnd.ms-excel.sheet.binary.macroEnabled.12"
        else if(fn:contains($filename,"xlsm")) then
              "application/vnd.ms-excel.sheet.macroEnabled.12"
        else if(fn:contains($filename,".xps")) then
              "application/vnd.ms-xpsdocument"
        else ()

  	return $mime-type
};
:)

(: this returns uris from directory, without directory, so they can be used in open xml package manifest :)
(: we have a function in package.xqy ooxml:package-uris(package as binary), which returns the uris from a package, not the dir) :)
declare function excel:directory-uris(
  $directory as xs:string
) as xs:string*
{
   (: get rid of let return if possible :)
    let $uris := cts:uris("","document",cts:directory-query($directory,"infinity"))
    return $uris
};

(:not sure about these next two, useful, but theres also a sheet1.xml.rels in there for some sheets, assumes we modify the sheet, but rest of pkg stays the same:)

(:look at uri-match //cts:uri-match? :)
declare function excel:directory-uris(
  $directory as xs:string, 
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

    let $uris := cts:uris("","document",cts:directory-query($directory,"infinity"))
    let $finaluris :=  
                  for $uri in $uris
                  let $u := $uri
                  where  fn:matches($uri, "sheet\d+\.xml$")
                  return $u
    return $finaluris
    
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

declare function excel:get-manifest(
  $directory as xs:string, 
  $uris as xs:string*) 
as element(zip:parts)
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

(: ============================================================================================================== :)
declare function excel:map-shared-strings(
  $sheet as node()*, 
  $shared-strings as xs:string*  (: pass in node  explicit element type and return type :)
) 
{
                  (: for $sheet in $sheets :)
                  let $rows := for $row at $d in $sheet//ms:row
                               let $cells  :=  for $cell at $e in $row/ms:c
                                               let $c := if(fn:data($cell/@t) eq "s") 
                                               then 
                                                 element ms:c { $cell/@* except $cell/@t, attribute t{"inlineStr"}, element ms:is { element ms:t { $shared-strings[($cell/ms:v+1 cast as xs:integer)] } } }  
                                               else
                                                    $cell
                                                    (: assumes cel is an int, have to account for any cell type :)
                                                        (: element ms:c { $cell/@*, $cell/ms:f, element ms:v { fn:data($cell/ms:v)} } :)
                                                           
                                               return $c
                     
                              return element ms:row{ $row/@*, $cells }
                               
                  let $worksheet   :=  $sheet/ms:worksheet/* except ( $sheet/ms:worksheet/ms:sheetData, $sheet/ms:worksheet/ms:tableParts, $sheet/ms:worksheet/ms:pageMargins, $sheet/ms:worksheet/ms:pageSetup)
                  let $page-setup :=  $sheet/ms:worksheet/ms:pageSetup 
                  let $table-parts :=  $sheet/ms:worksheet/ms:tableParts
                  let $sheet-data  :=  $sheet/ms:worksheet/ms:sheetData
(: get rid of lets , use element contstructor for start, or pass in directlly :)
                  let $ws := element ms:worksheet {  $sheet/ms:worksheet/@*, $worksheet, element ms:sheetData{ $sheet-data/@*, $rows },  $page-setup ,$table-parts  }
                  return $ws
(: return $rows :)

};

(: ============================================================================================================== :)
(:srs work here!:)
declare function excel:validate-child(
  $seq as node()*
) as xs:boolean
{
    let $original := ($seq )(: except () why original? , collapse away , if/else not required, just return count($disctinct-child) eq 1) :)
    let $children :=  $original/fn:local-name(child::*[1])
    let $distinct-child := fn:distinct-values(if($children eq ""(: how is this possible:) ) then () else $children)
    let $child-count := fn:count($distinct-child) (: fn:count(fn:normalize-space(text{$distinct-child})) :)
    let $result := if($child-count eq 1) then xs:boolean("true")  else xs:boolean("false")
    return $result
};
(: ============================================================================================================== :)
(:   feed a (), tell by xml type, assume give in right order :)
declare function excel:generate-simple-xl-pkg(
  $content-types as node(),
  $workbook as node(),
  $rels as node(),
  $workbookrels as node(),
  $sheets as node()*
) as binary()
{
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
		    </parts>
    let $parts := ($content-types, $workbook, $rels, $workbookrels, $sheets) 
    return
 	xdmp:zip-create($manifest, $parts)
};

(: assumes one table , could become node* , be precise about types:)
declare function excel:generate-simple-xl-pkg(
    $content-types as node(),
    $workbook as node(),
    $rels as node(),
    $workbookrels as node(),
    $sheets as node()*,
    $worksheetrels as node(),
    $table as node()
) as binary()
{
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
			<part>xl/worksheets/_rels/sheet1.xml.rels</part>
			<part>xl/tables/table1.xml</part>
		     </parts>
    let $parts := ($content-types, $workbook, $rels, $workbookrels, $sheets ,$worksheetrels,$table) 
    return
         xdmp:zip-create($manifest, $parts)
};

declare function excel:create-row(
  $values as xs:anyAtomicType*
) as element(ms:row)
{ 
    <ms:row>
    {
           for $val at $v in $values  (: check for dates :)
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

    let $rows := for $i at $d in $keys
                 let $val := map:get($map,$i)
                 let $return := if(fn:empty($val)) then "" else $val (: if empty, still create cell, string ? :)
                 return $return
    return excel:create-row($rows)
};

(:check for dates, also, overload function to include formulas, other children of ms:c :)
declare function excel:cell($a1-ref as xs:string, $value as xs:anyAtomicType)
{
    if($value castable as xs:integer) then     
              <ms:c r={$a1-ref}><ms:v>{$value}</ms:v></ms:c>
    else
              <ms:c r={$a1-ref} t="inlineStr"> 
                    <ms:is>
                        <ms:t>{$value}</ms:t>
                    </ms:is>
              </ms:c>
};


(: when adding cell to worksheet:
    1. if cell exists in worksheet, just replace
    2. next look for row
        a. if row exists, create cell at appropriate position (index based on column)
        b. if row dne, create row, add cell - done
    :)

declare function excel:a1-to-r1c1($a1notation as xs:string)
{
(: not sure if we need, probably, stubbing out :)
<foo/>
};
(: currently limited to 702 columns :)
(: base 26 arithmetic, want sequence numbers, each 0-25 (1-26), take number, repeatedly div mod til end :)(: mary function for hex , recursively mod til < 26 :)
declare function excel:r1c1-to-a1(
  $rowcount as xs:integer, 
  $colcount as xs:integer
) 
{
    let $coldiv := fn:floor($colcount div 26)
    let $letter := if($colcount <= 26)
                   then                  
                     fn:codepoints-to-string($colcount+64)
                 
                   else
                      let $coldiv := fn:floor($colcount div 26)
                      let $coldiv2 := $colcount div 26

                      let $first-letter := if($coldiv2 eq $coldiv )then 
                                              fn:codepoints-to-string($coldiv+63) 
                                           else fn:codepoints-to-string($coldiv+64)

                      let $next-letter-check := $colcount - ($coldiv2 * 26)
                      let $next-letter := $colcount - ($coldiv * 26)

                      let $final := if($next-letter-check eq 0) 
                                    then "Z" 
                                    else fn:codepoints-to-string($next-letter+64) 
                      return fn:concat($first-letter,$final)  

    return   fn:concat($letter,$rowcount)   
   
};

(: orig
declare function excel:r1c1-to-a1(
  $rowcount as xs:integer, 
  $colcount as xs:integer
) as xs:string
{
    let $coldiv := fn:floor($colcount div 26)
    let $colmod := $colcount mod 26
    let $letter := if($coldiv gt 0)
                   then                  
                     fn:codepoints-to-string($colcount+64)
                   else
                     fn:codepoints-to-string($colcount+64)
    return fn:concat($letter,$rowcount)
   
};
:)
(:rename col-width :)
declare function excel:worksheet-cols(
$widths as xs:string*
) as element(ms:cols)
{
    <ms:cols>
    { 
        for $i at $d in $widths
        return <ms:col min="{$d}" max="{$d}" width="{$i}" customWidth="1"/>
    }
    </ms:cols>
};
(: single table, do for multiple :)
declare function excel:create-simple-content-types(
  $worksheet-count as xs:integer,
  $tbl as xs:boolean (:table-count :)
) as element(types:Types)
{
    let $content-types := 
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        {
           for $i at $d in 1 to $worksheet-count (: d is redundant :)
           let $sheet-name := fn:concat("/xl/worksheets/sheet", $d )
           return
	     <Override PartName={$sheet-name} ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        }
        {
           if($tbl eq xs:boolean("true")) then 
             <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
            else
              ()
        }
       </Types>
    return $content-types
};

declare function excel:create-simple-workbook($worksheet-count as xs:integer) as element(ms:workbook)
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

declare function excel:create-simple-pkg-rels() as element(pr:Relationships)
{
    let $rels :=
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
       </Relationships>
    return $rels
};

declare function excel:create-simple-workbook-rels($worksheet-count as xs:integer) as element(pr:Relationships)
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
(: NAMING create - or constructor , look in triggers.xqy ?  triggerdataevent excel:worksheet-rels :)
declare function excel:create-simple-worksheet-rels() as element(pr:Relationships)
{
    let $worksheetrels:=
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
       </Relationships> 
    return $worksheetrels
};

declare function excel:create-simple-table(
  $tablerange as xs:string, 
  $column-names as xs:string*, 
  $style as xs:boolean
) as element(ms:table)
{

    let $column-count := fn:count($column-names)
    let $table :=
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="{$tablerange}"  totalsRowShown="0" >
         <autoFilter ref="{$tablerange}"/>
         <tableColumns count={$column-count}> (: put quotes :)
         {
           for $i at $d in $column-names
           return <tableColumn id={$d} name={$i}/>
         }
         </tableColumns>
         {
          let $t-style := if($style eq xs:boolean("true")) then  (: if($style) :)
           <tableStyleInfo name="TableStyleMedium10" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
          else ()
          return $t-style
         }
      </table>
    return $table
};

declare function excel:create-simple-worksheet(
  $rows as element(ms:row)*
) as element(ms:worksheet)
{
    let $sheet := <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                      { 
                        ($rows) 
                      }
                      </sheetData>
                  </worksheet>
    return $sheet
};

declare function excel:create-simple-worksheet(
  $rows as element(ms:row)*,
  $colwidths as element(ms:cols),
  $tablepart as xs:boolean
) as element(ms:worksheet)
{
    let $sheet := <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      {
                        ($colwidths)
                      } 
                      <sheetData>
                      { 
                        ($rows) 
                      }
                      </sheetData>
                      {
                        if($tablepart eq xs:boolean("true")) then (: get rid of eq boolean :)
                          <tableParts count="1">
                           <tablePart r:id="rId1" />
                          </tableParts>
                        else ()
                      }
                  </worksheet>
    return $sheet
};

declare function excel:create-simple-pkg(
  $worksheets as element(ms:worksheet)*
) as binary()
{
    let $ws-count := fn:count($worksheets)
    let $content-types := excel:create-simple-content-types($ws-count, (:fn:false()/fn:true:)xs:boolean("false"))
    let $workbook := excel:create-simple-workbook($ws-count)
    let $rels :=  excel:create-simple-pkg-rels()
    let $workbookrels :=  excel:create-simple-workbook-rels($ws-count)
    let $package := excel:generate-simple-xl-pkg($content-types, $workbook, $rels, $workbookrels, $worksheets)
    return $package

};


(:=============ADDED ===================================================== :)
declare function excel:a1-row($a1)
{
       fn:replace($a1,("[A-Z]+"),"")
};

declare function excel:a1-column($a1)
{
       fn:replace($a1,("\d+"),"")
};


declare function excel:passthru($x as node(), $newcell) as node()*
{
   for $i in $x/node() return excel:set-row-cell($i,$newcell)
};


declare function excel:insert-cell($origcell, $newcell)
{
 if($newcell/@r = $origcell/@r) then 
    $newcell
 else if(fn:empty($origcell/preceding-sibling::*))
 then
    (
     if($newcell/@r lt $origcell/@r)
     then
        ($newcell,$origcell)
     else $origcell
     )
 else if(fn:not(fn:empty($origcell/following-sibling::*)) and 
         $newcell/@r > $origcell/@r and 
         $newcell/@r < $origcell/following-sibling::*/@r) then  
           ($origcell,$newcell) 
 else if(fn:not(fn:empty($origcell/following-sibling::*)) and
         $newcell/@r < $origcell/@r and 
         $newcell/@r > $origcell/preceding-sibling::*/@r and
         $newcell/@r < $origcell/following-sibling::*/@r
        ) then  
           ($newcell,$origcell)
 else if(fn:empty($origcell/following-sibling::*) and
          $newcell/@r > $origcell/@r) then ($origcell,$newcell)
else $origcell

 
};

declare function excel:set-row-cell($x, $newcell)
{
 
      typeswitch($x)
       case text() return $x
       case document-node() return document {$x/@*,excel:passthru($x,$newcell)}
       case element(ms:c) return excel:insert-cell($x,$newcell)
       case element() return  element{fn:name($x)} {$x/@*,excel:passthru($x,$newcell)}
       default return $x

};


(:
declare function excel:row($cell)
{
  <ms:row r={excel:a1-row($cell/@r)}>{$cell}</ms:row> 
 
};

declare function excel:passthru-workbook($x as node(), $newSheetData) as node()*
{
   for $i in $x/node() return excel:wb-set-sheetdata($i,$newSheetData)
};

declare function excel:wb-set-sheetdata($x, $newSheetData)
{
 
      typeswitch($x)
       case text() return $x
       case document-node() return document {$x/@*,excel:passthru-workbook($x,$newSheetData)}
       case element(ms:sheetData) return  $newSheetData
       case element() return  element{fn:name($x)} {$x/@*,excel:passthru-workbook($x,$newSheetData)}
       default return $x

};

:)

(:
declare function excel:ws-set-cells($sheet as node(), $cells as element(ms:c)*) as node()*
{
   let $sheetData := $sheet/ms:worksheet/ms:sheetData

   let $finalsheet := (
   for $c in $cells return xdmp:set($sheetData,(
   let $refrow := excel:a1-row($c/@r)
   let $origrow := $sheetData/ms:row[@r=$refrow]

   (: pass multiple cells per row, have to order, send in groups :)

   let $newrow := if(fn:empty($origrow)) then excel:row($c)
                  else excel:set-row-cell($origrow,$c)

   let $rows := if(fn:exists($sheetData/ms:row[@r=$refrow])) then
                   for $r at $d in $sheetData/ms:row
                    (:not likely, but row could exist with no cells
                      update for that ?:)
                   let $row := if($r/@r = $newrow/@r) then $newrow 
                            else $r
                   return $row
                else for $newrow in ($sheetData/ms:row,$newrow)             
                     order by $newrow/@r cast as xs:integer
                     return $newrow
                     
   let $newSheetData := element ms:sheetData{ $sheetData/@*, $rows}   
   (: return (xdmp:set($sheetData,$newSheetData),$sheetData) :)
   return $newSheetData )),$sheetData)

return excel:wb-set-sheetdata($sheet/ms:worksheet, $finalsheet)                      
}; :)


declare function excel:row($cell)
{
  <ms:row r={excel:a1-row($cell/@r)}>{$cell}</ms:row> 
 
};

declare function excel:passthru-workbook($x as node(), $newSheetData) as node()*
{
   for $i in $x/node() return excel:wb-set-sheetdata($i,$newSheetData)
};

declare function excel:wb-set-sheetdata($x, $newSheetData)
{
 
      typeswitch($x)
       case text() return $x
       case document-node() return document {$x/@*,excel:passthru-workbook($x,$newSheetData)}
       case element(ms:sheetData) return  $newSheetData
       case element() return  element{fn:name($x)} {$x/@*,excel:passthru-workbook($x,$newSheetData)}
       default return $x

};


declare function excel:ws-set-cells($sheet as node(), $cells as element(ms:c)*) as node()*
{
   let $sheetData := $sheet/ms:worksheet/ms:sheetData

   let $finalsheet := (
   for $c in $cells return xdmp:set($sheetData,(
   let $refrow := excel:a1-row($c/@r)
   let $origrow := $sheetData/ms:row[@r=$refrow]

   (: pass multiple cells per row, have to order, send in groups :)

   let $newrow := if(fn:empty($origrow)) then excel:row($c)
                  else excel:set-row-cell($origrow,$c)

   let $rows := if(fn:exists($sheetData/ms:row[@r=$refrow])) then
                   for $r at $d in $sheetData/ms:row
                    (:not likely, but row could exist with no cells
                      update for that ?:)
                   let $row := if($r/@r = $newrow/@r) then $newrow 
                            else $r
                   return $row
                else for $newrow in ($sheetData/ms:row,$newrow)             
                     order by $newrow/@r cast as xs:integer
                     return $newrow
                     
   let $newSheetData := element ms:sheetData{ $sheetData/@*, $rows}   
   (: return (xdmp:set($sheetData,$newSheetData),$sheetData) :)
   return $newSheetData )),$sheetData)

return excel:wb-set-sheetdata($sheet/ms:worksheet, $finalsheet)                    

};


