var selectedEntity = null;

function initialize() {
  var xmlDoc;
  var xslDoc;

  xmlDoc = new ActiveXObject('Microsoft.XMLDOM')
  xmlDoc.async = false;

  xslDoc = new ActiveXObject('Microsoft.XMLDOM')
  xslDoc.async = false;

  xmlDoc.load("xmlTree.xml");
  xslDoc.load("xmlTree.xsl");

  folderTree.innerHTML = xmlDoc.documentElement.transformNode(xslDoc);
}

function clickOnEntity(entity) {

  if( entity.url != '#' &&  entity.url != '' ) 
      top.contents.location = entity.url;

  if(entity.open == "false") {
    expand(entity, true);
  }
  else {
    collapse(entity);
  }

  selectedEntity = entity;
  window.event.cancelBubble = "true";
}

function expand(entity) {
  var stateImage;
  var oImage;

  stateImage = document.all["stateImage" + entity.id];
  
  oImage = entity.childNodes[0].all["image"];
  oImage.src = entity.imageOpen;

  if(stateImage) {
		stateImage.src = stateImage._open;
  }

  for(i=0; i < entity.childNodes.length; i++) {
    if(entity.childNodes[i].tagName == "DIV") {
      entity.childNodes[i].style.display = "block";
    }
  }
  
  entity.open = "true";
}

function collapse(entity) {
  var stateImage;
  var oImage;
  
  var i;

  stateImage = document.all["stateImage" + entity.id];

  oImage = entity.childNodes[0].all["image"];
  oImage.src = entity.image;

  if(stateImage) {
		stateImage.src = stateImage._closed;
  }

  for(i=0; i < entity.childNodes.length; i++) {
      if(entity.childNodes[i].tagName == "DIV") {
        if(entity.id != "folderTree") entity.childNodes[i].style.display = "none";
        collapse(entity.childNodes[i]);
      }
    }
    
  entity.open = "false";
}

function expandAll(entity) {
  var oImage;
  var i;

  expand(entity, false);

  for(i=0; i < entity.childNodes.length; i++) {
    if(entity.childNodes[i].tagName == "DIV") {
      expandAll(entity.childNodes[i]);
    }
  }
}
