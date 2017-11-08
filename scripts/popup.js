var url = "https://api.trello.com";
var key = "bffdd838fd24b5832de6f72e47987533";
var token = "685daea308b5bff7e3d2da00807bcb34a12f1e24685928fc444acb0875d01429";
var reader = new FileReader();
var xls = XLSX;
var wb = {
    "Sheets": {},
    "Props": {},
    "SSF": {},
    "SheetNames": []
};


document.addEventListener('DOMContentLoaded', function() {
    var checkPageButton = document.getElementById('cl');
    checkPageButton.addEventListener('click', function() {
        getColumns("Cm1VvpPQ");
    }, false);
}, false);

function s2ab(s) {
    if(typeof ArrayBuffer !== 'undefined') {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    } else {
        var buf = new Array(s.length);
        for (var i=0; i!=s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
}

function getColumns(idTrello){
    //cns49FIZ
    $.ajax({
        url : url+"/1/boards/"+idTrello+"?fields=name&lists=open&key=" + key +"&token=" + token,
        type : "get",
        async: false,
        success : (data) => {
            if(!data) return;
            data.lists.forEach(function(item) {
                wb.SheetNames.push(item.name);
                getItensColum(item, item.name);
            });
            exportExcel();
        }
    });
}

function getItensColum(colum, wsName){
    $.ajax({
        url : url + "/1/lists/"+ colum.id +"/cards?actions=commentCard&key=" + key +"&token=" + token,
        type : "get",
        async: false,
        success : (data) => {
            gerarSheets(data, wsName)
        }
    });
}

function gerarSheets(itensColumns, wsName){
    var ws = {};
    var rang = { s: {c:0, r:0}, e: {c:0, r:0 } };
    var titulos = ["Descrição","Email enviado","Orçamento enviado","Escopo aprovado"];
    
    for (var i = 0; i < titulos.length; i++) {
        var cell = { v: "", t: "" };
        if(rang.e.c < i) rang.e.c = i;
        cell.v = titulos[i];
        var cellRefTit = xls.utils.encode_cell({ c: rang.e.c, r: rang.e.r });
        ws[cellRefTit] = verificaCell(cell);
    }
   
    itensColumns.forEach((item, index) => {
        //descri
        var cell = { v: "", 
                     t: "",
                    //   s: {
                    //    top:{style: "medium", color: { rgb: "FFFFAA00" }},
                    //    bottom:{style: "medium", color: { rgb: "FFFFAA00" }},
                    //    left:{style: "medium", color: { rgb: "FFFFAA00" }},
                    //    right:{style: "medium", color: { rgb: "FFFFAA00" }}
                    // } 
                };

        if(rang.e.r < index) rang.e.r = index;
        cell.v = item.name;
        //cell.s = { font: {sz: 16, bold: true, color: { rgb: "FFFFAA00" }} };
        var cell_ref = xls.utils.encode_cell({ c: 0, r: rang.e.r + 1 });
        ws[cell_ref] = verificaCell(cell);
        
        //teste
        var tests = [
            { r: /Email enviado/, cell: 1 },
            { r: /Orçamento enviado/, cell: 2 },
            { r: /Escopo aprovado/, cell: 3 }
        ];

        if(item.actions == null || item.actions.length == 0) return;
        item.actions.forEach((action) => {
            cell = { v: "", t: "" };
            tests.forEach((test) => {
                var regex = test.r.exec(action.data.text);
                if(regex != null && regex.length){
                    try {
                        cell.v = /\d{2}[/]\d{2}[/]\d{4}/.exec(action.data.text)[0];
                    }
                    catch (e) {
                        cell.v = "";
                    }
                    //cell.s = { font: {sz: 16, bold: true, color: { rgb: "FFFFAA00" }} };
                    var cellRefTest = xls.utils.encode_cell({ c: test.cell, r: rang.e.r + 1 });
                    ws[cellRefTest] = verificaCell(cell);
                }
            });
        });
    });
    var wscols = [
        {wch: 80},
        {wch: 30},
        {wch: 30},
        {wch: 30}
    ];

    ws['!ref'] = xls.utils.encode_range(rang);
    ws['!cols'] = wscols;
    wb.Sheets[wsName] = ws;
}

function exportExcel(){
    var wbout = xls.write(wb, {bookType:"xlsx", bookSST:true, type: 'binary', cellStyles: true});
    var blob = new Blob([s2ab(wbout)],{type:"application/octet-stream"});
    var URL = window.URL || window.webkitURL;
    var downloadUrl = URL.createObjectURL(blob);
    chrome.downloads.download({url:downloadUrl, filename:"teste.xlsx"},function(id) { });
} 

function verificaCell(cell){
    if(typeof cell.v === 'number') cell.t = 'n';
    else if(typeof cell.v === 'boolean') cell.t = 'b';
    else cell.t = 's';

    return cell;
}
