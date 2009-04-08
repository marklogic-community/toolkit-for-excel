/* Copyright 2002 Jean-Claude Manoli [jc@manoli.net]
 *
 * This software is provided 'as-is', without any express or implied warranty.
 * In no event will the author(s) be held liable for any damages arising from
 * the use of this software.
 * 
 * Permission is granted to anyone to use this software for any purpose,
 * including commercial applications, and to alter it and redistribute it
 * freely, subject to the following restrictions:
 * 
 *   1. The origin of this software must not be misrepresented; you must not
 *      claim that you wrote the original software. If you use this software
 *      in a product, an acknowledgment in the product documentation would be
 *      appreciated but is not required.
 * 
 *   2. Altered source versions must be plainly marked as such, and must not
 *      be misrepresented as being the original software.
 * 
 *   3. This notice may not be removed or altered from any source distribution.
 */ 

var treeSelected = null; //last treeNode clicked

//pre-load tree nodes images
var imgPlus = new Image();
imgPlus.src="images/treenodeplus.gif";
var imgMinus = new Image();
imgMinus.src="images/treenodeminus.gif";
var imgDot = new Image();
imgPlus.src="images/treenodedot.gif";


function findNode(el)
{
// Takes element and determines if it is a treeNode.
// If not, seeks a treeNode in its parents.
	while (el != null)
	{
		if (el.className == "treeNode")
		{
			break;
		}
		else
		{
			el = el.parentNode;
		}
	}
	return el;
}


function clickAnchor(el)
{
// handles click on a TOC link
//
	expandNode(el.parentNode);
	selectNode(el.parentNode);
	el.blur();
}


function selectNode(el)
{
// Un-selects currently selected node, if any, and selects the specified node
//
	if (treeSelected != null)
	{
		setSubNodeClass(treeSelected, 'A', 'treeUnselected');
	}
	setSubNodeClass(el, 'A', 'treeSelected');
	treeSelected = el;
}


function setSubNodeClass(el, nodeName, className)
{
// Sets the specified class name on el's first child that is a nodeName element
//
	var child;
	for (var i=0; i < el.childNodes.length; i++)
	{
		child = el.childNodes[i];
		if (child.nodeName == nodeName)
		{
			child.className = className;
			break;
		}
	}
}


function expandCollapse(el)
{
//	If source treeNode has child nodes, expand or collapse view of treeNode
//
	if (el == null)
		return;	//Do nothing if it isn't a treeNode
		
	var child;
	var imgEl;
	for(var i=0; i < el.childNodes.length; i++)
	{
		child = el.childNodes[i];
		if (child.src)
		{
			imgEl = child;
		}
		else if (child.className == "treeSubnodesHidden")
		{
			child.className = "treeSubnodes";
			imgEl.src = "images/treenodeminus.gif";
			break;
		}
		else if (child.className == "treeSubnodes")
		{
			child.className = "treeSubnodesHidden";
			imgEl.src = "images/treenodeplus.gif";
			break;
		}
	}
}


function expandNode(el)
{
//	If source treeNode has child nodes, expand it
//
	var child;
	var imgEl;
	for(var i=0; i < el.childNodes.length; i++)
	{
		child = el.childNodes[i];
		if (child.src)
		{
			imgEl = child;
		}
		if (child.className == "treeSubnodesHidden")
		{
			child.className = "treeSubnodes";
			imgEl.src = "images/treenodeminus.gif";
			break;
		}
	}
}

function syncTree(name)
{
// Selects and scrolls into view the node that references the specified URL
//
	var tocEl = findHref(document.getElementById('treeRoot'), name);
	if (tocEl != null)
	{
		selectAndShowNode(tocEl);
		return;
	} 

}

// this findHref function is specific to this content
// I don't understand why this does not work in IE 
function findHref(node, name)
{
//	name = name.replace(/&/g, "&amp\;");
	name = name.replace(/%20/g, " ");
	if ( name.indexOf("?") >= 0 ) {
	name = "#" + name;
	var anchors = node.getElementsByTagName('A');
	for (var i = 0; i < anchors.length; i++)
	{
		el = anchors[i];
		var ahref = new String();
		ahref = el.getAttribute('href');
		// IE seems to make the href a fully qualified path
		if ( navigator.userAgent.indexOf("MSIE") >= 0 ) {
		ahref = ahref.substr(ahref.indexOf("#"), ahref.length);
		}
		if ( ahref == name  ) {
			return el;
		}
	}
	} else {
// look on the div id for fast track
	var divs = node.getElementsByTagName('DIV');
	for (var i = 0; i < divs.length; i++)
	{
		el = divs[i];
		var divid = new String();
		divid = el.getAttribute('id');
		if ( divid == name ) {
			return el;
		}
	}
	}

}

// danny; 6-19-2006; This is a new version of this function.  It
// simply always scrolls to the selected item.  I could never
// get the original version to scroll properly.  
function selectAndShowNode(node)
{
// Selects and scrolls into view the specified node

	var el = findNode(node);
	if (el != null) 
	{
		selectNode(el);
		do 
		{
			expandNode(el);
			el = findNode(el.parentNode);
		} while ((el != null))  
		
	//vertical scroll element into view
        var treeDiv = document.getElementById('tree');
        var nodePosition;

        nodePosition = node.offsetTop + treeDiv.offsetTop;
        // offset by 600 to make up for the picture and the form 
        // comment out because we have divs, not frames
        // window.scroll(0, nodePosition - 600 );
        
        // scroll the div
        if (window.XMLHttpRequest) {
        document.getElementById('contents').scrollTop = nodePosition - 140;
        } else {
        }
        }
}

function resizeTree()
{
	var treeDiv = document.getElementById('tree');
	//treeDiv.setAttribute('style', 'width: ' + document.body.offsetWidth + 'px; height: ' + (document.body.offsetHeight - 27) + 'px;');
	treeDiv.style.width = document.documentElement.offsetWidth;
	treeDiv.style.height = document.documentElement.offsetHeight - 27;
}
