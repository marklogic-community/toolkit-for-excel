function insertImage(picuri)
{
       var config = MLA.getConfiguration();
       var fullurl= config.url;
       //alert("config url"+fullurl);
       var picuri = fullurl + "/search/download-support.xqy?uid="+picuri;
       var msg = MLA.insertImage(picuri,"oslo","oslo");
}

function copyPasteSlideToActive(docuri)
{

       // alert("here");
       // alert("docuri for testOpen():"+ docuri);
       var tokens = docuri.split("/");
       var filename = tokens[tokens.length-1];

       //use filename to get slide number for now
       //get .pptx name from docuri for now
       //until we embed in xml somewhere (properties?)

       //alert("filename "+filename);
       var tmpfilename = filename.replace(".GIF","");
       tmpfilename = tmpfilename.replace("Slide","");

       //var slideidx = parseInt(tmpfilename);
       var slideidx = tmpfilename;
       //alert("tmpfilename "+tmpfilename + "slideidx: "+slideidx);
       var idx = docuri.indexOf("_GIF");
       var tmpuri1 = docuri.substring(0,idx);
       var tmpuri2 = docuri.substring(1,idx);
       var newuri = tmpuri1+".pptx";
       var newfilename = tmpuri2+".pptx";
       

       //alert("newuri: "+newuri);

       //alert("filename"+filename);
       var tmpPath = MLA.getTempPath();
       //alert("here2 "+tmpPath) 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       //var url = fullurl + "/search/download-support.xqy?uid="+docuri;
       var url = fullurl + "/search/download-support.xqy?uid="+newuri;

       //alert("fullurl"+fullurl);
       //alert("URL: "+url);
	//window.external.openXlsx(filename,docuri);
       var msg = MLA.copyPasteSlideToActive(tmpPath, newfilename,slideidx, url, "oslo","oslo");
     // var msg = window.external.OpenXlsx(tmpPath, docuri, url, "zeke","zeke");
}

$(document).ready(function(){

$("ul.thumb li").hover(function() {
	$(this).css({'z-index' : '10'}); /*Add a higher z-index value so this image stays on top*/ 
	$(this).find('img').addClass("hover").stop() /* Add class of "hover", then stop animation queue buildup*/
		.animate({
			/*marginTop: '-2px', /* The next 4 lines will vertically align this image */ 
			/*marginLeft: '-130px',*/
			/*top: '5%', */			/*left: '10%',*/
			width: '325px', /* Set new width */
			height: '275px', /* Set new height */
			padding: '10px'
		}, 200); /* this value of "200" is the speed of how fast/slow this hover animates */

	} , function() {
	$(this).css({'z-index' : '0'}); /* Set z-index back to 0 */
	$(this).find('img').removeClass("hover").stop()  /* Remove the "hover" class , then stop animation queue buildup*/
		.animate({
			marginTop: '0', /* Set alignment back to default */
			marginLeft: '0',
			top: '0',
			left: '0',
			width: '250px', /* Set width back to default */
			height: '200px', /* Set height back to default */
			padding: '5px'
		}, 400);
});
});

