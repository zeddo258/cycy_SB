let $ = require('jquery') // Module jquery to select
let fs = require('fs') // Module fs to rw file
let modal = document.getElementById('optionModal') 
const reader =  require('xlsx-color') // Module xlsx
const file = reader.readFile('./test.xlsx')

// Get sheet into array
let worksheets = {} 
let contacts = {}


let num = 0 
for (const sheetName of file.SheetNames) {
    worksheets[sheetName] = reader.utils.sheet_to_json(file.Sheets[sheetName])
}

modal.style.display = "none"; 

document.getElementById("inf_refuse").onclick = function() {
    modal.style.display = "none"; 
    var selection = document.getElementById("contacts-table");
    // remove all the option to reset form
    for(var i = 0; i < selection.length; i++ ) {
        selection.remove(i);
    }
} 

document.getElementById("inf_accept").onclick = function() {
    var temp = $("#contacts-table option:selected").val(); 
    var target = contacts[temp]; 
    writeData(target);
    alert("完成更新資料庫！！！");
    window.location.href = 'index.html';
} 


$('#find').on('click', () => {
    if ( document.getElementById("id").value == "" ) {
        alert("請輸入學號/人事代碼");
    }

    else { 
        worksheets = {}
        contacts = {}
        for (const sheetName of file.SheetNames) {
            worksheets[sheetName] = reader.utils.sheet_to_json(file.Sheets[sheetName])
        }
        var temp = worksheets["Sheet1"]
        let found = false; 
        num = 0; 
        for ( var i = 0; i <  temp.length; i++ ) {
            if ( parseInt(temp[i]["學號"]) == parseInt(document.getElementById("id").value) && !temp[i]["狀態"]) {
                var person = {
                    "id": temp[i]["學號"], 
                    "date": temp[i]["日期"], 
                    "w10_32": temp[i]["(32)Windows10"],
                    "w10_64": temp[i]["(64)Windows10"],
                    "off_16": temp[i]["Office2016"],
                    "off_19": temp[i]["Office2019"], 
                    "off_16m": temp[i]["Office2016Mac"], 
                    "off_19m": temp[i]["Office2019Mac"], 
                    "sas_93": temp[i]["SAS9.3"], 
                    "sas_94": temp[i]["SAS9.4"], 
                    "vs_2012": temp[i]["VisualStudio2012"], 
                    "vs_2013": temp[i]["VisualStudio2013"], 
                    "vs_2015": temp[i]["VisualStudio2015"], 
                    "eviews": temp[i]["EVIEWS"], 
                    "nat": temp[i]["自然輸入法"], 
                    "wu": temp[i]["無蝦米"], 
                    "jin": temp[i]["金蝶333"] 
                }; 


                var string = temp[i]["日期"] + ",\t" +
                             temp[i]["姓名"] + ",\t" + 
                             temp[i]["學號"] + ",\t" + 
                             getSoftwareFrom(temp[i]);  
                             
                $('#contacts-table').addClass("text-base leading-relaxed text-black").append(new Option(string, ++num));
                contacts[num] = person; 
                found = true; 
            }  
                          
        } 
        if (found)
            modal.style.display = "flex"; 
        else 
            alert("此學號/人事代碼沒有紀錄")
    }
    
    
})


function match(currentCell, targetCell) {
    if (currentCell["(32)Windows10"] == targetCell["w10_32"] &&  
        currentCell["(64)Windows10"] ==  targetCell["w10_64"] && 
        currentCell["Office2016"] == targetCell["off_16"] &&
        currentCell["Office2019"] == targetCell["off_19"] &&
        currentCell["Office2016Mac"] == targetCell["off_16m"] && 
        currentCell["Office2019Mac"] == targetCell["off_19m"] &&
        currentCell["SAS9.3"] == targetCell["sas_93"] &&
        currentCell["SAS9.4"] == targetCell["sas_94"] &&
        currentCell["VisualStudio2012"] == targetCell["vs_2012"] && 
        currentCell["VisualStudio2013"] == targetCell["vs_2013"] && 
        currentCell["VisualStudio2015"] == targetCell["vs_2015"] && 
        currentCell["EVIEWS"] == targetCell["eviews"] &&
        currentCell["自然輸入法"] == targetCell["nat"]  && 
        currentCell["無蝦米"] == targetCell["wu"] && 
        currentCell["金蝶333"] == targetCell["jin"] &&
        currentCell["學號"] == targetCell["id"] && 
        currentCell["日期"] == targetCell["date"] ) 
        return true
    return false 
}


function getSoftwareFrom(currentCell) {
    var str = "";
    if (parseInt(currentCell["(32)Windows10"]) > 0 ) str += "(32)Windows10: " + parseInt(currentCell["(32)Windows10"]) + " | "; 
    if (parseInt(currentCell["(64)Windows10"]) > 0 ) str += "(64)Windows10: " + parseInt(currentCell["(32)Windows10"]) + " | "; 
    if (parseInt(currentCell["Office2016"]) > 0 ) str += "Office2016: " + parseInt(currentCell["Office2016"]) + " | "; 
    if (parseInt(currentCell["Office2019"]) > 0 ) str += "Office2019: " + parseInt(currentCell["Office2019"]) + " | "; 
    if (parseInt(currentCell["Office2016Mac"]) > 0 )  str += "Office2016Mac: " + parseInt(currentCell["Office2016Mac"]) + " | "; 
    if (parseInt(currentCell["Office2019Mac"]) > 0 ) str += "Office2019Mac: " + parseInt(currentCell["Office2019Mac"]) + " | "; 
    if (parseInt(currentCell["SAS9.3"]) > 0 ) str += "SAS9.3: " + parseInt(currentCell["SAS9.3"]) + " | "; 
    if (parseInt(currentCell["SAS9.4"]) > 0 ) str += "SAS9.4: " + parseInt(currentCell["SAS9.4"]) + " | "; 
    if (parseInt(currentCell["VisualStudio2012"]) > 0 ) str += "VisualStudio2012: " + parseInt(currentCell["VisualStudio2012"]) + " | "; 
    if (parseInt(currentCell["VisualStudio2013"]) > 0 ) str += "VisualStudio2013: " + parseInt(currentCell["VisualStudio2013"]) + " | "; 
    if (parseInt(currentCell["VisualStudio2015"]) > 0 ) str += "VisualStudio2015: " + parseInt(currentCell["VisualStudio2015"]) + " | "; 
    if (parseInt(currentCell["EVIEWS"]) > 0 ) str += "EVIEWS: " + parseInt(currentCell["EVIEWS"]) + " | "; 
    if (parseInt(currentCell["自然輸入法"]) > 0 )  str += "自然輸入法: " + parseInt(currentCell["自然輸入法"]) + " | "; 
    if (parseInt(currentCell["無蝦米"]) > 0 ) str += "無蝦米: " + parseInt(currentCell["無蝦米"]) + " | "; 
    if (parseInt(currentCell["金蝶333"]) > 0 )  str += "金蝶333: " + parseInt(currentCell["金蝶333"]) + " | "; 
    return str; 
}

function writeData(targetCell) {
   
    var temp = worksheets["Sheet1"];
    console.log(targetCell["id"]); 
    for ( var i = 0; i <  temp.length; i++ ) {
        if (match(temp[i],targetCell))
            temp[i]["狀態"] = "已還"; 
    } 
    reader.utils.sheet_add_json(file.Sheets["Sheet1"], worksheets.Sheet1);
    reader.writeFile(file,'./test.xlsx'); 

}