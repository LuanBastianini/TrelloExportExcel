var url = "https://api.trello.com";
var key = "bffdd838fd24b5832de6f72e47987533";
var token = "685daea308b5bff7e3d2da00807bcb34a12f1e24685928fc444acb0875d01429";


document.addEventListener('DOMContentLoaded', function() {
    var checkPageButton = document.getElementById('cl');
    checkPageButton.addEventListener('click', function() {
        getColumns("cns49FIZ");
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
                console.log(item);
                getItensColum(item);
            });
        }
    });
}

function getItensColum(colum){
    $.ajax({
        url : url + "/1/lists/"+ colum.id +"/cards?actions=commentCard&key=" + key +"&token=" + token,
        type : "get",
        async: false,
        success : (data) => {
            console.log(data);
            // exportExcel(data, colum);
        }
    });
}

function exportExcel(itensColumns, colum){
    var reader = new FileReader();
    var xls = XLSX;
    var ws = {};
    var ws_name = colum.name;
    var wb = {
        "Sheets": {},
        "Props": {},
        "SSF": {},
        "SheetNames": []
    };

    var cellRefTit = xls.utils.encode_cell({ c: 0, r: 0 });
    ws[cellRefTit] = "Descrição";

    cellRefTit = xls.utils.encode_cell({ c: 1, r: 0 });
    ws[cellRefTit] = "Email enviado";

    cellRefTit = xls.utils.encode_cell({ c: 2, r: 0 });
    ws[cellRefTit] = "Orçamento enviado";

    cellRefTit = xls.utils.encode_cell({ c: 3, r: 0 });
    ws[cellRefTit] = "Escopo aprovado";
    
    var rang = { cell: 0, row: 0 };
    itensColumns.forEach((itemColum) => {
        //descri
        var cell_ref = xls.utils.encode_cell({ c: 0, r: rang.row });
        ws[cell_ref] = item.name;
        
        //teste
        var tests = [
            { r: /Email enviado/, cell: 1, qtd: 0 },
            { r: /Orçamento enviado/, cell: 2, qtd: 0 },
            { r: /Escopo aprovado/, cell: 3, qtd: 0 },
            { r: /teste/, cell: 4, qtd: 0 }
        ];

        if(item.actions == null || item.actions.length == 0) continue;
        item.actions.forEach((action) => {
            tests.forEach((test) => {
                if(test.exec(action.data.text).length){
                    test.qtd++;
                    //gambiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiii
                    var cellRefTest = xls.utils.encode_cell({ c: test.cell, r: rang.row });
                    ws[cellRefTest] = test.qtd;
                }
            });
        });
    });
}