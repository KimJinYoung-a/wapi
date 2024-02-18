function DrawTnGridTag(iGridName,width,height){
    var gridTag = '';
    gridTag += " <OBJECT";
    gridTag += " ID='" + iGridName + "' ";
    gridTag += " Name='" + iGridName + "' ";
    gridTag += " classid='clsid:83173DA2-A77F-4DA2-8B5B-E272B1E0C79F' ";
    gridTag += " codebase='http://scm.10x10.co.kr/lib/util/cab/AxTenGrid.cab#version=1,0,1,1' ";
    gridTag += " width=" + width;
    gridTag += " height=" + height;
    gridTag += " align=center";
    gridTag += " hspace=0";
    gridTag += " vspace=0";
    gridTag += " ></OBJECT>";
    
    document.write(gridTag);
}

function DrawTnDTPicker(iDTPName,defaultDate){
    DrawTnDTPicker2(iDTPName,defaultDate,110,24,10);
}

function DrawTnDTPicker2(iDTPName,defaultDate,width,height,fontsize){
    var DTPTag = '';
    DTPTag += " <OBJECT";
    DTPTag += " ID='" + iDTPName + "' ";
    DTPTag += " Name='" + iDTPName + "' ";
    DTPTag += " classid='clsid:5EBD07F7-E4A9-4921-B5BD-801F9378392F' ";
    DTPTag += " codebase='http://scm.10x10.co.kr/lib/util/cab/AxTenGrid.cab#version=1,0,0,3' ";
    DTPTag += " width=" + width;
    DTPTag += " height=" + height;
    DTPTag += " align=center";
    DTPTag += " hspace=0";
    DTPTag += " vspace=0";
    DTPTag += " >";
    DTPTag += " <PARAM NAME='defaultDate' VALUE='" + defaultDate + "'>";
    DTPTag += " <PARAM NAME='FontSize' VALUE='" + fontsize + "'>";
    DTPTag += " </OBJECT>";
    document.write(DTPTag);
}
