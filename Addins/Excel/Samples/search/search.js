function openXlsx(docuri)
{
	// alert("docuri for testOpen():"+ docuri);
        var tokens = docuri.split("/");
	var filename = tokens[tokens.length-1];
         //alert("filename"+filename);
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
       //alert("URL: "+url);
	//window.external.openXlsx(filename,docuri);
     var msg = MLA.openXlsx(tmpPath, docuri, url, "zeke","zeke");
     // var msg = window.external.OpenXlsx(tmpPath, docuri, url, "zeke","zeke");
}
