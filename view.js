let $ = require('jquery') // Module jquery to select
let fs = require('fs') // Module fs to rw file
let info = document.getElementById('ifModal')
let modal = document.getElementById('defaultModal')
let alertModal = document.getElementById('alertModal')

const reader =  require('xlsx-color') // Module xlsx
const file = reader.readFile('./test.xlsx')

let worksheets = {} 
let total = 0 

document.getElementById("modal0_accept").onclick = function() {
  modal.style.display = "none"; 
} 

document.getElementById("inf_refuse").onclick = function() {
    modal.style.display = "none"; 
    var selection = document.getElementById("contacts-table");
    // remove all the option to reset form
    for(var i = 0; i < selection.length; i++ ) {
        selection.remove(i);
    }
} 


$('#alertAccept').on('click', () => {
    alertModal.style.display = "none";
})

alertModal.style.display = "none"; 
info.style.display = "none"; 

document.getElementById("inf_accept").onclick = function() {
    info.style.display = "none"; 
    writeData(); 
    alert("完成寫入資料庫");
    reset(); 
    window.location.href = 'index.html';
}

document.getElementById("inf_refuse").onclick = function() {
    info.style.display = "none"; 
}


$('#submit').on('click', () => {
        
        total = parseInt(document.getElementById("o-19w").value) + 
                    parseInt(document.getElementById("o-16w").value) +
                    parseInt(document.getElementById("o-19m").value) + 
                    parseInt(document.getElementById("o-16m").value) + 
                    parseInt(document.getElementById("w10-32").value) +  
                    parseInt(document.getElementById("w10-64").value) +  
                    parseInt(document.getElementById("sas93").value) + 
                    parseInt(document.getElementById("sas-94").value) +  
                    parseInt(document.getElementById("vs-15").value) +
                    parseInt(document.getElementById("vs-13").value) +
                    parseInt(document.getElementById("vs-12").value) +
                    parseInt(document.getElementById("ev").value) + 
                    parseInt(document.getElementById("nat").value) + 
                    parseInt(document.getElementById("wu").value) +
                    parseInt(document.getElementById("usb").value)  

        
        if ( total > 0  && document.getElementById("id").value != "" ) {
            var str = getSoftware(); 
            $('#contacts-table').addClass("text-xl leading-relaxed text-gray-500 dark:text-gray-400").append(new Option("名字："+document.getElementById("name").value), document.getElementById("name").value);
            $('#contacts-table').addClass("text-xl leading-relaxed text-gray-500 dark:text-gray-400").append(new Option("學號："+document.getElementById("id").value,document.getElementById("id").value));
            $('#contacts-table').addClass("text-xl leading-relaxed text-gray-500 dark:text-gray-400").append(new Option("科系："+document.getElementById("dep").value,document.getElementById("dep").value));
            $('#contacts-table').addClass("text-xl leading-relaxed text-gray-500 dark:text-gray-400").append(new Option("電話："+document.getElementById("phone").value, document.getElementById("phone").value));
            for ( var j = 0; j < str.length; j++) {
        
                $('#contacts-table').addClass("text-xl leading-relaxed text-gray-500 dark:text-gray-400").append(new Option(str[j], str[j]));
            }
            info.style.display = "flex";
        
        }

        else if ( total == 0 && document.getElementById("id").value == "" || document.getElementById("date").value == "" ) 
            alertModal.style.display = "flex"; 
        
})


function reset() {
    const inputs = document.querySelectorAll('#date, #dep,#id,#name,#phone,#o-19w,#o-16w,#o-19m,#o-16m,#w10-32,#w10-64,#sas93,#sas-94,#vs-15,#vs-13,#vs-12,#ev,#nat,#wu,#usb')
    inputs.forEach(input => {
    input.value = '';
    });
 }

function getSoftware() {
    var str = [];

    if (parseInt(document.getElementById("o-19w").value) > 0 ) str.push("Office 19(Windows): " + document.getElementById("o-19w").value) ;  
    if (parseInt(document.getElementById("o-16w").value) > 0 ) str.push("Office 16(Windows): " + document.getElementById("o-16w").value) ; 
    if (parseInt(document.getElementById("o-19m").value) > 0 ) str.push("Office 19(Mac): " + document.getElementById("o-19m").value) ; 
    if (parseInt(document.getElementById("o-16m").value) > 0 ) str.push("Office 16(Mac): " + document.getElementById("o-16m").value) ; 
    if (parseInt(document.getElementById("w10-32").value) > 0 ) str.push("Windows-10(32 bits): " + document.getElementById("w10-32").value) ; 
    if (parseInt(document.getElementById("w10-64").value) > 0 ) str.push("Windows-10(64 bits): " + document.getElementById("w10-64").value) ; 
    if (parseInt(document.getElementById("sas93").value) > 0 ) str.push("SAS 9.3: " + document.getElementById("sas93").value) ; 
    if (parseInt(document.getElementById("sas-94").value) > 0 ) str.push("SAS 9.4: " + document.getElementById("sas-94").value) ; 
    if (parseInt(document.getElementById("vs-15").value) > 0 ) str.push("Visual Studio 15: " + document.getElementById("vs-15").value) ; 
    if (parseInt(document.getElementById("vs-13").value) > 0 ) str.push("Visual Studio 13: " + document.getElementById("vs-13").value) ; 
    if (parseInt(document.getElementById("vs-12").value) > 0 ) str.push("Visual Studio 12: " + document.getElementById("vs-12").value) ; 
    if (parseInt(document.getElementById("ev").value) > 0 ) str.push("EVIEWS: " + document.getElementById("ev").value) ; 
    if (parseInt(document.getElementById("nat").value) > 0 ) str.push("自然輸入法： " + document.getElementById("nat").value) ; 
    if (parseInt(document.getElementById("wu").value) > 0 ) str.push("無瑕米： " + document.getElementById("wu").value) ; 
    if (parseInt(document.getElementById("usb").value) > 0 ) str.push("金蝶333: " + document.getElementById("usb").value) ; 

    return str
}

function writeData() {
    worksheets = {} 
            
    for (const sheetName of file.SheetNames) {
        worksheets[sheetName] = reader.utils.sheet_to_json(file.Sheets[sheetName])
    }
    
    var temp = worksheets["Sheet1"]
    var sno = 0
    if(temp.length == 0 ) sno = 1
    else {
        sno = temp[temp.length - 1]["流水"] + 1
    }
    worksheets.Sheet1.push({
        "流水" : sno,
        "姓名" : document.getElementById("name").value, 
        "狀態" : "", 
        "日期" : document.getElementById("date").value, 
        "學號" : document.getElementById("id").value, 
        "單位" : document.getElementById("dep").value,
        "電話" : document.getElementById("phone").value, 
        "總數" : total,
        "(32)Windows10" : document.getElementById("w10-32").value,
        "(64)Windows10" : document.getElementById("w10-64").value,
        "Office2016" : document.getElementById("o-16w").value,
        "Office2019" : document.getElementById("o-19w").value,
        "Office2016Mac" : document.getElementById("o-16m").value, 
        "Office2019Mac" : document.getElementById("o-19m").value,
        "SAS9.3" : document.getElementById("sas93").value,
        "SAS9.4" : document.getElementById("sas-94").value,
        "VisualStudio2012" : document.getElementById("vs-12").value, 
        "VisualStudio2013" : document.getElementById("vs-13").value, 
        "VisualStudio2015" : document.getElementById("vs-15").value, 
        "EVIEWS" : document.getElementById("ev").value, 
        "自然輸入法" : document.getElementById("nat").value, 
        "無蝦米" : document.getElementById("wu").value, 
        "金蝶333" : document.getElementById("usb").value,
    })

    reader.utils.sheet_add_json(file.Sheets["Sheet1"], worksheets.Sheet1)
    reader.writeFile(file,'./test.xlsx')
    

}