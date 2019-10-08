
var myFile = File.openDialog("Select transdoc INDD file");
var file;
var path = "";

if(myFile != null) {
	app.open(myFile)
	} else{
	exit();
	};

var items = [];
var changed = 0;	
var allTables = app.activeDocument.stories.everyItem().tables.everyItem().getElements();
var cellength = allTables[0].cells.length;

for(var i = 0; i<cellength; i++){
	items.push([allTables[0].cells[i].texts[0].contents, allTables[0].cells[i+1].texts[0].contents])
	i++;
}

app.activeDocument.close();



//---------------------
	showDialog();
	
	
	var mytext;
	function showDialog(title) {
    
    var w = new Window("dialog", title);
    w.orientation = "row";
    w.alignChildren = "top";
    
    var editGroup = w.add("group");
    editGroup.alignChildren = "left";
    editGroup.orientation = "column";

    var langsEditBox = editGroup.add("edittext", undefined, "new files folder path");
    langsEditBox.preferredSize.width = 240
    
    var buttonGroup = w.add("group");
    buttonGroup.size = { width: 80, height: 100 };
    buttonGroup.alignChildren = "right";
    buttonGroup.orientation = "column";
    buttonGroup.add("button", undefined, "OK");
    buttonGroup.add("button", undefined, "Cancel");

    if (w.show() == 1) {

        mytext = langsEditBox.text;
		//alert(mytext);
      
        if(mytext.charAt(mytext.length - 1) == "\\")
		{
		path = mytext;
		//alert("ze slaszem");
		}  
       else {
		path = mytext + "\\" ;
		//alert("bez slasza");
	   }
    }
    else {
        exit();
    }
}

//--------------------

//alert("w tabelce jest "+ items.length + " linkow do zmiany");

var doc = app.activeDocument;

var oSourceDoc = app.activeDocument;
var mLinks = app.activeDocument.links.everyItem().getElements();
var links = doc.allGraphics;

var lines = [];
var bad = [];
var thisLinkIsmulti = false;

for(var i=links.length-1;i>=0;i--){
	
	var myOldLinkName = mLinks[i].name;				

	var ext = links[i].itemLink.name.substr(links[i].itemLink.name.lastIndexOf("."));  
    var old = File(links[i].itemLink.filePath);
   	
    var oldNME = links[i].itemLink.name;
	var oldLNK = links[i].itemLink;
   	var found = false;
	
	var Howmany = LinkUsage(oldLNK);
	
	
	
	
	//alert("link " +oldNME+ " jest razy: " + Howmany + " i jest status multi na " + thisLinkIsmulti) ;
    for(var j=0;j<items.length;j++){
    		//alert("path: " +path);
			//alert("newname: " +newname);
			//alert("FILE: "+file);
			
			
    	if(items[j][0] == oldNME){
				
    		var newname =  items[j][1]; 
			try{
    			
				
				//old.rename(newname);  
				//links[i].itemLink.relink(File(old.toString().replace(links[i].itemLink.name,newname)));
				
				//old.rename(newname);  
				file = new File(path + newname);
				links[i].itemLink.relink(file);
				
				
				changed++;	
				
				
				var temp = j+1;
				//alert("lines pusz: " + temp);
				lines.push(j+1);
				found = true;
				break;
			}//end try
				catch(err){
				//alert( items: " + items[j][0] + "\r\n" + " file nowy: " + file+ "\r\n"+ " multi: " + Howmany );
				//alert("Cant relink " + items[j][0] + " to  "+ items[j][1] +"  because: " + "\r\n" + err);
				bad.push(j+1);
			}//end catch
			
    	}//end if

    } //end for j
	
} //end for i

//alert(lines);
//alert(lines.length);



function LinkUsage(myLink) {
	var myLinksNumber = 0;
		for (var c =  0; c < oSourceDoc.links.length; c++) {
		if (myLink.filePath == oSourceDoc.links[c].filePath) {
			myLinksNumber += 1;
		}
	}
	return myLinksNumber;
}

//alert("linkow bylo " + items.length + " a ostatecznie zmieniono " +  changed);
//alert("linkow bylo " + mLinks.length + " a ostatecznie zmieniono " +  changed);

//alert("lines imp: " + lines);



if(myFile != null) {
	app.open(myFile)
	} else{
	exit();
	};
var allTables = app.activeDocument.stories.everyItem().tables.everyItem().getElements();

var linesMax = Math.max.apply(Math, lines);
//alert("max: " +linesMax);
//alert("lines length: " +lines.length);
for(var z=1;z<linesMax+1;z++){
	
	for(var s=0;s<lines.length;s++){
	//alert("z: " +z +" s: " + s + "      lines[s]: "  +  lines[s]);
		if(z == lines[s]){
			allTables[0].rows[z-1].fillColor = "C=75 M=5 Y=100 K=0";
				
		}
	}
	
}


var badMax = Math.max.apply(Math, bad);

for(var r=1;r<badMax+1;r++){
	
	for(var t=0;t<bad.length;t++){
	
		if(r == bad[t]){
			allTables[0].rows[r-1].fillColor = "C=15 M=100 Y=100 K=0";
				
		}
	}
	
}






