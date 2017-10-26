document.addEventListener('DOMContentLoaded', function() {
    var checkPageButton = document.getElementById('cl');
    checkPageButton.addEventListener('click', function() {
        //Para buscar todas listas
        //https://api.trello.com/1/boards/Cm1VvpPQ?fields=id,name&lists=open&list_fields=id,name,closed,pos&key=175bab29d30ac5c2db6f9b9de3de5b4b&token=b7f75def67760a80d454cf834ff6414fc75d5f4342662927fd0cec6c0af5c9c8
      $.get("https://api.trello.com/1/lists/5811dd0416356cda5996a190/cards?actions=commentCard&key=175bab29d30ac5c2db6f9b9de3de5b4b&token=b7f75def67760a80d454cf834ff6414fc75d5f4342662927fd0cec6c0af5c9c8",
                function(datas){
        /* original data */
        var data = [
                ["Nome card","comentarios",3],
                [datas[0].actions[0].data.card.name, datas[0].actions[0].data.text, null],
                [datas[1].actions[0].data.card.name, datas[1].actions[0].data.text,"0.3"], 
                [datas[2].actions[0].data.card.name, datas[2].actions[0].data.text, "qux"]
            ]
        var ws_name = "SheetJS";
        var reader = new FileReader();

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

        /* require XLSX */
        //var XLSX = require('xlsx')
        var X = XLSX;

        /* set up workbook objects -- some of these will not be required in the future */
        var wb = {}
        wb.Sheets = {};
        wb.Props = {};
        wb.SSF = {};
        wb.SheetNames = [];

        /* create worksheet: */
        var ws = {}

        /* the range object is used to keep track of the range of the sheet */
        var range = {s: {c:0, r:0}, e: {c:0, r:0 }};

        /* Iterate through each element in the structure */
        for(var R = 0; R != data.length; ++R) {
            if(range.e.r < R) range.e.r = R;
            for(var C = 0; C != data[R].length; ++C) {
             if(range.e.c < C) range.e.c = C;

            /* create cell object: .v is the actual data */
                var cell = { v: data[R][C] };
                if(cell.v == null) continue;

            /* create the correct cell reference */
                var cell_ref = X.utils.encode_cell({c:C,r:R});

            /* determine the cell type */
                if(typeof cell.v === 'number') cell.t = 'n';
                else if(typeof cell.v === 'boolean') cell.t = 'b';
                else cell.t = 's';

            /* add to structure */
            ws[cell_ref] = cell;
        }
        }
        ws['!ref'] = X.utils.encode_range(range);

        /* add worksheet to workbook */
        wb.SheetNames.push(ws_name);
        wb.Sheets[ws_name] = ws;

        /* write file */
        var wbout = XLSX.write(wb, {bookType:"xlsx", bookSST:true, type: 'binary'});
        var blob = new Blob([s2ab(wbout)],{type:"application/octet-stream"});
        var URL = window.URL || window.webkitURL;
        var downloadUrl = URL.createObjectURL(blob);
        chrome.downloads.download({url:downloadUrl, filename:"teste.xlsx"},function(id) { });
    });
    }, false);
  }, false);

