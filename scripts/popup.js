document.addEventListener('DOMContentLoaded', function() {
    var checkPageButton = document.getElementById('cl');
    checkPageButton.addEventListener('click', function() {
      $.get("https://api.trello.com/1/lists/59ede17bf56387cc316c98f8/cards?actions=commentCard&key=713ac2e39493aed425ac298cba624de4&token=19edbfbddc426811dffe535365752bbc9163cd29f3cf9467eaf1f45d5ca09fc4",
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

