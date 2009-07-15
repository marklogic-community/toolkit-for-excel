function insertImage(picuri)
{
       //alert("inserting images");
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}
/*
function copyPasteSlideToActive(docuri, slideidx)
{
       alert("copyPasteSlideToActive(docuri,slideidx)");

       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       //use filename to get slide number for now
       //get .pptx name from docuri for now
       //until we embed in xml somewhere (properties?)

       var tmpfilename = filename.replace(".GIF","");
       tmpfilename = tmpfilename.replace("Slide","");

       var slideidx = tmpfilename;
       var idx = docuri.indexOf("_GIF");
       var tmpuri1 = docuri.substring(0,idx);
       
       var tmpuri2 = docuri.substring(1,idx);
       var newuri = tmpuri1+".pptx";
       var newfilename = tmpuri2+".pptx";
       

       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;

       var filename = docuri.substring(1,docuri.length);
       var msg = MLA.copyPasteSlideToActive(tmpPath, filename,slideidx, url, "oslo","oslo");
}
*/
function copyPasteSlideToActive(docuri, slideidx, retainidx)
{
      // alert("in this function");

       var retain=document.getElementById("retain"+retainidx).checked;
       var tmpPath = MLA.getTempPath();

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;
      
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1]; 
       //var filename = docuri.substring(1,docuri.length);
       //alert("url: "+url+"  filename: "+filename+" slidedix: "+slideidx+" retain: "+retain);
       var msg = MLA.copyPasteSlideToActive(tmpPath, filename,slideidx, url, "oslo","oslo",retain);
}

function openPPTX(docuri)
{
       //alert("docuri for testOpen():"+ docuri);
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];
       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/search/download-support.xqy?uid="+docuri;

       var msg = MLA.openPPTX(tmpPath, filename, url, "oslo","oslo");
}

