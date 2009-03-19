function saveXlsxToML()
{
       var ele = document.getElementById("ML-Save");
       var doctitle = ele.value;
       if(doctitle=="")
       {
	   doctitle="Default.xlsx";
       }
       else
       {
	   doctitle=doctitle+".xlsx";
       }

       var tmpPath = MLA.getTempPath(); 

       var config = MLA.getConfiguration();
       var fullurl= config.url;
       var url = fullurl + "/save/upload2.xqy?uid="+doctitle;

       var saveas = tmpPath+doctitle;

       var msg = MLA.saveActiveWorkbook(tmpPath, doctitle, url, "zeke","zeke");

       if(msg=="")
	       alert("workbook:" + doctitle + " saved.");
}
	
