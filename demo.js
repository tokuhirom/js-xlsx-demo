"use strict";

var data = [
    {date:'2014-06-11', num:2300},
    {date:'2014-06-10', num:11300},
    {date:'2014-06-09', num:300},
    {date:'2014-06-08', num:3400},
    {date:'2014-06-07', num:4300},
    {date:'2014-06-06', num:8300},
    {date:'2014-06-05', num:3300},
    {date:'2014-06-04', num:331300},
    {date:'2014-06-03', num:300},
    {date:'2014-06-02', num:3700},
    {date:'2014-06-01', num:8300}
];

function ReportCtrl($scope) {
    $scope.data = data;

    function downloadCSV() {
        function datenum(v, date1904) {
            if(date1904) v+=1462;
            var epoch = Date.parse(v);
            return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        }
    
        function sheet_from_array_of_arrays(data, opts) {
            var ws = {};
            var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
            for(var R = 0; R != data.length; ++R) {
                for(var C = 0; C != data[R].length; ++C) {
                    if(range.s.r > R) range.s.r = R;
                    if(range.s.c > C) range.s.c = C;
                    if(range.e.r < R) range.e.r = R;
                    if(range.e.c < C) range.e.c = C;
                    var cell = {v: data[R][C] };
                    if(cell.v == null) continue;
                    var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
                    
                    if(typeof cell.v === 'number') cell.t = 'n';
                    else if(typeof cell.v === 'boolean') cell.t = 'b';
                    else if(cell.v instanceof Date) {
                        cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                        cell.v = datenum(cell.v);
                    }
                    else cell.t = 's';
                    
                    ws[cell_ref] = cell;
                }
            }
            if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
            return ws;
        }
    
        var data = [];
        $('.report tr').each(function (i, tr) {
        var row = [];
            $(tr).find('th,td').each(function (j, td) {
                var o = ''+td.innerHTML;
                if ($(td).hasClass("report-num")) {
                    o = parseInt(o, 10);
                }
                row.push(o);
            });
            data.push(row);
        });
        console.log(data);

        function Workbook() {
            if(!(this instanceof Workbook)) return new Workbook();
            this.SheetNames = [];
            this.Sheets = {};
        }

        var key = XLSX.utils.encode_cell({c:0,r:0});
        var ws = sheet_from_array_of_arrays(data);

        var workbook = new Workbook();
        workbook.SheetNames.push("売上");
        workbook.Sheets["売上"] = ws;

        var wbout = XLSX.write(workbook, {
            bookType:'xlsx',
            bookSST:true,
            type:'binary'
        });

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }

        saveAs(new Blob([s2ab(wbout)],{type:""}), "report.xlsx");
    }

    $scope.downloadCSV = function () {
        downloadCSV();
    };

    $scope.uriage = function (row) {
        return row.num * 1000;
    };
    $scope.tax = function (uriage) {
        return uriage * 0.08;
    };
}

