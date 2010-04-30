/* 
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

MarkLogicContentControlSupport.js - javascript api captures Content Control events in Word and return SimpleContentControl objects.
*/

function contentControlOnEnter(id,tag,title, type, lockcontrol, lockcontents, parenttag, parentid)
{
	try{
	var msg = "";

	var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
            mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid;

	    onEnterHandler(mlacontrolref);
	}
	catch(err)
 	{
		msg="error in contentControlOnEnter: "+err.description;
                //throw("Error: Not able to create SimpleContentControl from input. ");
	}

	//alert("ENTER ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);
	//
	//
        
	return msg;

}

function contentControlOnExit(id, tag, title, type, lockcontrol, lockcontents, parenttag, parentid)
{
	try{
	var msg = "";

        var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
	    mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid;

	    onExitHandler(mlacontrolref);
	}
	catch(err)
 	{
		msg="error in contentControlOnExit: "+err.description;
               // throw("Error: Not able to create SimpleContentControl from input. ");
	}
	
	//alert("EXIT ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);

	return msg;
}

function contentControlAfterAdd(id, tag, title, type, lockcontrol, lockcontents, parenttag, parentid)
{

	//HERE'S WHERE WE'LL ADD METADATA!!
	//How to get metadata type? from tag?
	try{
	var msg = "";

        var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
	    mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid
	
	    afterAddHandler(mlacontrolref);
	}
	catch(err)
 	{
		msg="error in contentControlAfterAdd: "+err.description;
                //throw("Error: Not able to create SimpleContentControl from input. ");
	}

//alert("AFTER ADD ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);

	return msg;
}

function contentControlBeforeDelete(id, tag, title, type, lockcontrol, lockcontents, parenttag, parentid)
{
	//HERE'S WHERE WE'LL DELETE METADATA!!
	try{
	var msg = "";

        var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
	    mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid;

	    beforeDeleteHandler(mlacontrolref);
	}
	catch(err)
 	{
		msg="error in contentControlBeforeDelete: "+err.description;
                //throw("Error: Not able to create SimpleContentControl from input. ");
	}

//	alert("BEFORE DELETE ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);

	return msg;
}

function contentControlBeforeContentUpdate(id, tag, title, type, lockcontrol, lockcontents, parenttag, parentid)
{
	try{
	var msg = "";

        var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
            mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid;
	}
	catch(err)
 	{
		msg="error in contentControlBeforeContentUpdate: "+err.description;
                //throw("Error: Not able to create SimpleContentControl from input. ");
	}

//	alert("BEFORE CONTENT UPDATE ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);

	return msg;
}

function contentControlBeforeStoreUpdate(id, tag, title, type, lockcontrol, lockcontents, parenttag, parentid)
{
	try{
	var msg = "";

        var mlacontrolref = new MLA.SimpleContentControl(id); 
	    mlacontrolref.tag = tag;
            mlacontrolref.title = title; 
            mlacontrolref.type = type;
            mlacontrolref.lockcontrol = lockcontrol;
            mlacontrolref.lockcontents = lockcontents;
            mlacontrolref.parentTag = parenttag;
	    mlacontrolref.parentID = parentid;


	}
	catch(err)
 	{
		msg="error in contentControlBeforeStoreUpdate: "+err.description;
                //throw("Error: Not able to create SimpleContentControl from input. ");
	}

//	alert("BEFORE STORE UPDATE ---> message "+mlacontrolref.id+" tag "+mlacontrolref.tag+" title"+mlacontrolref.title+" type "+mlacontrolref.type + "lockcontrol"+ mlacontrolref.lockcontrol + " lockcontents"+ mlacontrolref.lockcontents +" parentTag "+mlacontrolref.parentTag + " parentID: "+ mlacontrolref.parentID);

	return msg;
}
