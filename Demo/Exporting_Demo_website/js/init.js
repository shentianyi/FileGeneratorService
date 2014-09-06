var chart;
var containers = ['container_1', 'container_2', 'container_3'];
//'HighchartsExport.axd'
// var url = 'http://192.168.0.253:9002/HighChartsFileService/';
//var url = 'HighchartsExport.axd';
 //var url = 'http://42.121.111.38:9002/HighChartsFileService/';
 // var url = 'HighchartsExport.axd';
// var url = 'http://127.0.0.1:9002/HighChartsFileService/';
 var url = 'http://localhost:6023/HighChartsFileService/';

$(document).ready(function() {
     var charts = [];
     // charts array

     // generate 3 different charts
     for(var i = 0; i < containers.length; i++) {
          var data = [];
          for(var m = 0; m < 4; m++) {
               data[m] = [];
               for(var n = 0; n < 12; n++) {
                    data[m].push({
                         name : "D-" + m + "-" + n,
                         y : Math.random()*100 
                    });
               }
          }
          var chart = new Highcharts.Chart({
               chart : {
                    renderTo : containers[i]
               },
               title : {
                    text : 'Monthly Average Temperature' + i,
                    x : -20 //center
               },
               subtitle : {
                    text : 'Source: WorldClimate.com' + i,
                    x : -20
               },
               xAxis : {
                    categories : ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
               },
               yAxis : {
                    title : {
                         text : 'Temperature (°C)'
                    },
                    plotLines : [{
                         value : 0,
                         width : 1,
                         color : '#808080'
                    }]
               },
               tooltip : {
                    formatter : function() {
                         return '<b>' + this.series.name + '</b><br/>' + this.x + ': ' + this.y + '°C';
                    }
               },
               legend : {
                    layout : 'vertical',
                    align : 'right',
                    verticalAlign : 'top',
                    x : -10,
                    y : 100,
                    borderWidth : 0
               },
               // exporting module
               exporting : {
                    url : url,
                    filename : 'MyChart',
                    width : 1200, // chart width
                    exportTypes : ['chart', 'png', 'jpeg', 'pdf', 'svg', 'doc', 'docx', 'pptx', 'xls', 'xlsx'] // set download file type
               },
               series : [{
                    name : 'Tokyo' + i,
                    data : data[0]
               }, {
                    name : 'New York' + i,
                    data : data[1]
               }, {
                    name : 'Berlin' + i,
                    data : data[2]
               }, {
                    name : 'London' + i,
                    data : data[3]
               }]
          });
          charts.push(chart);
     }

     // download all charts as pdf
     $("#exportPDFs_btn").click(function() {
          // console.log(charts[0].series);
          Highcharts.exportCharts({
               url : url,
               type : 'pdf',
               charts : charts
          });
     });
     // download all charts as doc
     $("#exportDOCs_btn").click(function() {
          Highcharts.exportCharts({
               url : url,
               type : 'doc',
               width : 600, // chart with
               charts : charts
          });
     });
     // download all charts as docx
     $("#exportDOCXs_btn").click(function() {
          Highcharts.exportCharts({
               url : url,
               type : 'docx',
               width : 600,
               charts : charts
          });
     });
     // download all charts as xls
     $("#exportXLSs_btn").click(function() {
          Highcharts.exportCharts({
               url : url,
               type : 'xls',
               width : 600,
               charts : charts
          });
     });
     // download all charts as xlsx
     $("#exportXLSXs_btn").click(function() {
          Highcharts.exportCharts({
               url : url,
               type : 'xlsx',
               width : 1000,
               charts : charts
          });
     });

     // download all charts as pptx
     $("#exportPPTXs_btn").click(function() {
          Highcharts.exportCharts({
               url : url,
               type : 'pptx',
               width : 700,
               charts : charts
          });
     });
});
