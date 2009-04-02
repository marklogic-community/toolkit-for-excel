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

declare function excel:directory-uris(
  $directory as xs:string
) as xs:string*
{
     cts:uris("","document",cts:directory-query($directory,"infinity"))
};


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
    cts:uri-match(fn:concat($directory,"*sheet*.xml"))
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
(:   for future, feed a (), determine file by xml type, assume given in right order :)
declare function excel:create-simple-xlsx(
  $worksheets as element(ms:worksheet)*
) as binary()
{
    let $ws-count := fn:count($worksheets)
    let $content-types := excel:content-types($ws-count,0)
    let $workbook := excel:workbook($ws-count)
    let $rels :=  excel:pkg-rels()
    let $workbookrels :=  excel:workbook-rels($ws-count)
    let $package := excel:xl-pkg($content-types, $workbook, $rels, $workbookrels, $worksheets)
    return $package

};

(:optional worksheetrels, table parameters :)
declare function excel:xl-pkg(
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
declare function excel:xl-pkg(
    $content-types as node(),
    $workbook as node(),
    $rels as node(),
    $workbookrels as node(),
    $sheets as node()*,
    $worksheetrels as node()*,
    $table as node()*
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
};

(: a couple of ways to create-row(s), use excel:row constructor :)
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

(: dates are stored as a julian number with an @ for style which indicates display format :)
(: need to update these to include style :)
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

declare function excel:cell($a1-ref as xs:string, $value as xs:anyAtomicType?, $formula as xs:string)
{
    if($value castable as xs:integer or fn:empty($value)) then     
              <ms:c r={$a1-ref}>
                   <ms:f>{$formula}</ms:f>
                   {
                    if(fn:not($value eq 0) and fn:not(fn:empty($value)))
                    then
                       <ms:v>{$value}</ms:v>
                    else ()
                   }
              </ms:c>
    else
              <ms:c r={$a1-ref} t="inlineStr"> 
                    <ms:is>
                        <ms:t>{$value}</ms:t>
                    </ms:is>
              </ms:c>
};

declare function excel:row($cells)
{
   
  <ms:row r={excel:a1-row($cells[1]/@r)}>{$cells}</ms:row> 
 
};

declare function excel:a1-to-r1c1($a1notation as xs:string)
{
    (: not sure if we need, probably, stubbing out for now :)
    <foo/>
};
(: currently limited to 702 columns :)
(: base 26 arithmetic, want sequence numbers, each 0-25 (1-26), take number, repeatedly div mod til end :)(: mary function for hex , recursively mod til < 26 :)

(:  BUGS!! ORIG BEFORE ONE BELOW
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
:)

(: WORKS PROPERLY TO 702 :)
(:
declare function excel2:r1c1-to-a1(
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

                      let $next-letter-check := $colcount - fn:floor($coldiv2 * 26)
                      let $next-letter := $colcount - ($coldiv * 26)

                     (: let $final := if($next-letter-check eq 0) 
                                    then "Z"  
                                    else :) 
                      let $final := if($next-letter-check = $next-letter and $next-letter-check mod 26 = 0)then 
                                               fn:codepoints-to-string($next-letter-check +90)   
                                           else fn:codepoints-to-string($next-letter+64) 
                      return fn:concat($first-letter,$final)  

    return   fn:concat($letter,$rowcount)   
   
};

:)
(: works for all and then some, excel has limit of WID(16384) for columns and 1048576 for rows 
   should we check and return error for # out of range? :)

declare function excel:r1c1-to-a1(
  $rowcount as xs:integer, 
  $colcount as xs:integer
) 
{   let $first-letter := (: "FOO" :)
                         if($colcount >= 703) then
                let $newcol := fn:floor($colcount div 702)

                let $flcheck := $colcount - fn:floor($newcol* 702)
                let $l := $colcount - ($newcol* 702)


                let $first-letter := if($flcheck = $l and $flcheck mod 702 = 0) then 
                     fn:codepoints-to-string($newcol+63)
                 else
                 fn:codepoints-to-string($newcol+64)
                return $first-letter
                else ""
                    
    let $ucol := if($colcount >= 703) then
                    (: $colcount :)
                     let $delta := $colcount mod 702 
                     let $ncol := if($delta <= 26 )then 
                                        $delta + 26 
                                   else $delta   
                     return $ncol
                 else 
                    $colcount      

    let $coldiv := fn:floor($ucol div 26 )
    let $letter := if($ucol <= 26 and $colcount > 703)
                     then "ZZ" else 
                   if($ucol <= 26)
                    
                   then                  
                     fn:codepoints-to-string($ucol+64)
                 
                   else
                      let $coldiv := fn:floor($ucol div 26)
                      let $coldiv2 := $ucol div 26

                      let $first-letter := if($coldiv2 eq $coldiv )then 
                                                fn:codepoints-to-string($coldiv+63)  
                                           else  fn:codepoints-to-string($coldiv+64)

                      let $next-letter-check := $ucol - fn:floor($coldiv2 * 26)
                      let $next-letter := $ucol - ($coldiv * 26)

                     (: let $final := if($next-letter-check eq 0) 
                                    then "Z"  
                                    else :) 

                      let $final := if($next-letter-check = $next-letter and $next-letter-check mod 26 = 0)then 
                                               fn:codepoints-to-string($next-letter-check +90)   
                                           else fn:codepoints-to-string($next-letter+64) 
                      return fn:concat($first-letter,$final)  

    return   fn:concat($first-letter,$letter,$rowcount)   
   
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

declare function excel:column-width(
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

declare function excel:pkg-rels() as element(pr:Relationships)
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
(: auto-filter, style are optional - defaults are "true", and no style respectively :)
declare function excel:table(
  $table-number as xs:integer,
  $tablerange as xs:string, 
  $column-names as xs:string*,
  $auto-filter as xs:boolean 
) as element(ms:table)
{

    let $disp-name := fn:concat("Table",$table-number)
    let $id := $table-number

    let $column-count := fn:count($column-names)
    let $table :=
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id={$id} name={$disp-name} displayName={$disp-name} ref="{$tablerange}"  totalsRowShown="0" >
         {
         if($auto-filter)then
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
      </table>
    return $table
};

declare function excel:table(
  $table-number as xs:integer,
  $tablerange as xs:string, 
  $column-names as xs:string*
) as element(ms:table)
{

    let $disp-name := fn:concat("Table",$table-number)
    let $id := $table-number

    let $column-count := fn:count($column-names)
    let $table :=
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id={$id} name={$disp-name} displayName={$disp-name} ref="{$tablerange}"  totalsRowShown="0" >
             <autoFilter ref="{$tablerange}"/>
             <tableColumns count={$column-count}> 
             {
               for $i at $d in $column-names
               return <tableColumn id={$d} name={$i}/>
             }
             </tableColumns>
      </table>
    return $table
};

declare function excel:table(
  $table-number as xs:integer,
  $tablerange as xs:string, 
  $column-names as xs:string*,
  $auto-filter as xs:boolean, 
  $style as xs:boolean
) as element(ms:table)
{

    let $disp-name := fn:concat("Table",$table-number)
    let $id := $table-number

    let $column-count := fn:count($column-names)
    let $table :=
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id={$id} name={$disp-name} displayName={$disp-name} ref="{$tablerange}"  totalsRowShown="0" >
         {
         if($auto-filter)then
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
    let $sheet := <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                      { 
                        ($rows) 
                      }
                      </sheetData>
                  </worksheet>
    return $sheet
};

declare function excel:worksheet(
  $rows as element(ms:row)*,
  $colwidths as element(ms:cols)?
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
                  </worksheet>
    return $sheet
};

declare function excel:worksheet(
  $rows as element(ms:row)*,
  $colwidths as element(ms:cols)?,
  $tbl-count as xs:integer
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
                        if($tbl-count  gt 0) then
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

(:=============ADDED ===================================================== :)
(:utility functions to grab Column or Row from "A1" notation :)
declare function excel:a1-row($a1)
{
       fn:replace($a1,("[A-Z]+"),"")
};

declare function excel:a1-column($a1)
{
       fn:replace($a1,("\d+"),"")
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

declare function excel:passthru($x as node(), $newcell) as node()*
{
   for $i in $x/node() return excel:set-row-cell($i,$newcell)
};


declare function excel:insert-cell($origcell, $newcell)
{

 if($newcell/@r = $origcell/@r) then 
    $newcell  (: cell exists :)
 else if(fn:empty($origcell/preceding-sibling::*))
 then
    (
     if($newcell/@r lt $origcell/@r)
     then
          ($newcell,$origcell) 
     else if($newcell/@r gt $origcell/@r and ( ($newcell/@r lt $origcell/following-sibling::*/@r) or fn:not(fn:exists($origcell/following-sibling::*/@r)))) then ($origcell, $newcell) else 
       ($origcell) 
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
else  $origcell  

 
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

declare function excel:ws-set-cells($v_sheet as element(ms:worksheet) , $cells as element(ms:c)*) as node()*
{  
   let $sheetDataTst := $v_sheet//ms:sheetData
   let $sheet := if(fn:empty($sheetDataTst)) then
                        let $row := excel:row(excel:cell("A1",""))
                        let $tmpSheet := excel:worksheet($row)
                        return excel:wb-set-sheetdata($v_sheet, $tmpSheet//ms:sheetData)
                 else $v_sheet

   let $sheetData :=  $sheet//ms:sheetData
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

return excel:wb-set-sheetdata($sheet(:/ms:worksheet:), $finalsheet)                      
};

declare function excel:julian-to-gregorian($excel-julian-day)
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
   let $finmonth := if(fn:string-length($month cast as xs:string) eq 1) then fn:concat("0",$month) else $month
   let $findate := fn:concat($year,"-", $finmonth, "-",$day,"T00:00:00")

   (: return  ($day, $month, $year) :)
   return   xs:dateTime($findate)

};

declare function excel:gregorian-to-julian($year, $month, $day)
{
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

