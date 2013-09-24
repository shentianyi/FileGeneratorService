( function(Highcharts) {
          var Chart = Highcharts.Chart, addEvent = Highcharts.addEvent, extend = Highcharts.extend, defaultOptions = Highcharts.getOptions();

          extend(defaultOptions.lang, {
               downloadDOC : 'Download DOC',
               downloadDOCX : 'Download DOCX',
               downloadPPT : 'Download PPT',
               downloadPPTX : 'Download PPTX',
               downloadXLS : 'Download XLS',
               downloadXLSX : 'Download XLSX'
          });

          var defaultExportButtonsData = {
               png : {
                    textKey : 'downloadPNG',
                    type : 'image/png'
               },
               jpeg : {
                    textKey : 'downloadJPEG',
                    type : 'image/jpeg'
               },
               pdf : {
                    textKey : 'downloadPDF',
                    type : 'application/pdf'
               },
               svg : {
                    textKey : 'downloadSVG',
                    type : 'image/svg+xml'
               },
               doc : {
                    textKey : 'downloadDOC',
                    type : 'application/msword'
               },
               docx : {
                    textKey : 'downloadDOCX',
                    type : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
               },
               xls : {
                    textKey : 'downloadXLS',
                    type : 'application/vnd.ms-excel'
               },
               xlsx : {
                    textKey : 'downloadXLSX',
                    type : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
               },
               ppt : {
                    textKey : 'downloadPPT',
                    type : 'application/vnd.ms-powerpoint'
               },
               pptx : {
                    textKey : 'downloadPPTX',
                    type : 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
               }
          };
          var exportTableFileTypes = ["xls", "xlsx"];
          var defaultExportButtons = {
               'chart' : {
                    textKey : 'printChart',
                    onclick : function() {
                         this.print();
                    }
               }
          };
          $.each(defaultExportButtonsData, function(key, value) {
               defaultExportButtons[key] = {
                    textKey : value.textKey,
                    onclick : function() {
                         this.exportChart({
                              type : key
                         });
                    }
               };
          });
          // console.log(defaultExportButtons);
          extend(defaultOptions.exporting, {
               enableExtendExport : true,
               exportTypes : ['png']
          });

          extend(Chart.prototype, {
               addExportButton : function() {
                    var exportingOptions = this.options.exporting, buttons = exportingOptions.buttons;
                    if(exportingOptions.enableExtendExport) {
                         extend(buttons.contextButton, {
                              menuItems : []
                         });
                         for(var i = 0; i < exportingOptions.exportTypes.length; i++) {
                              buttons.contextButton.menuItems.push(defaultExportButtons[exportingOptions.exportTypes[i]]);
                         }
                    }
               },
               exportChart : function(options, chartOptions) {
                    options = options || {};

                    var chart = this, chartExportingOptions = chart.options.exporting, svg = chart.getSVG(Highcharts.merge({
                         chart : {
                              borderRadius : 0
                         }
                    }, chartExportingOptions.chartOptions, chartOptions, {
                         exporting : {
                              sourceWidth : options.sourceWidth || chartExportingOptions.sourceWidth,
                              sourceHeight : options.sourceHeight || chartExportingOptions.sourceHeight
                         }
                    }));

                    // merge the options
                    options = Highcharts.merge(chart.options.exporting, options);
                    //console.log(chart.series);
                    //console.log(chart.xAxis.labels);
                    //console.log(chart.xAxis.categories);

                    var table = null;
                    table = chart.generateTableData();
                    // if (exportTableFileTypes.indexOf(options.type) != -1) {
                    // table = chart.generateTableData();
                    // }
                    // do the post
                    Highcharts.post(options.url, {
                         filename : options.filename || 'chart',
                         type : defaultExportButtonsData[options.type].type,
                         width : options.width || 0, // IE8 fails to post undefined correctly, so use 0
                         scale : options.scale || 2,
                         svg : svg,
                         muti : 0,
                         tabled : exportTableFileTypes.indexOf(options.type) != -1,
                         table : JSON.stringify(table)
                    });
               },
               generateTableData : function() {
                    // var table={series:[],rows:[][],xAxis:{},xAxisIndex:[]};
                    var table = {
                         series : [],
                         rows : [],
                         xAxis : {}
                    };
                    // generate table xAxis by order
                    var _xAxis = {};
                    for(var i = 0; i < this.series.length; i++) {
                         var data = this.series[i].data;
                         for(var j = 0; j < data.length; j++) {
                              if(!(data[j].x in _xAxis)) {
                                   _xAxis[data[j].x] = data[j].name;
                              }
                         }
                    }
                    var _xAxisIndexs = Object.keys(_xAxis).sort();
                    for(var i = 0; i < _xAxisIndexs.length; i++) {
                         table.xAxis[_xAxisIndexs[i]] = _xAxis[_xAxisIndexs[i]];
                    }
                    // init data
                    for(var i = 0; i < this.series.length; i++) {
                         table.series.push(this.series[i].name);
                         table.rows[i] = {};
                         for(var j = 0; j < _xAxisIndexs.length; j++) {
                              table.rows[i][_xAxisIndexs[j]] = null;
                         }
                         var data = this.series[i].data;
                         for(var j = 0; j < data.length; j++) {
                              if(data[j].x in table.rows[i]) {
                                   table.rows[i][data[j].x]=data[j].y
                              }
                         }
                    }
                    console.log(table);
                    return table;
               }
          });
          Highcharts.exportCharts = function(options, chartOptions) {
               options = options || {};
               var charts = options.charts;
               var svgs = [];
               var chartExportingOptions = charts[0].options.exporting;
               for(var i = 0; i < charts.length; i++) {
                    var chart = charts[i];
                    var svg = chart.getSVG(Highcharts.merge({
                         chart : {
                              borderRadius : 0
                         }
                    }, chartExportingOptions.chartOptions, chartOptions, {
                         exporting : {
                              sourceWidth : options.sourceWidth || chartExportingOptions.sourceWidth,
                              sourceHeight : options.sourceHeight || chartExportingOptions.sourceHeight
                         }
                    }));
                    svgs.push(svg);
               }

               // merge the options
               options = Highcharts.merge(chartExportingOptions, options);
               // do the post
               Highcharts.post(options.url, {
                    filename : options.filename || 'chart',
                    type : defaultExportButtonsData[options.type].type,
                    width : options.width || 0, // IE8 fails to post undefined correctly, so use 0
                    scale : options.scale || 2,
                    svg : svgs,
                    muti : 1,
                    tabled:false
               });
          };
          Chart.prototype.callbacks.unshift(function(chart) {
               chart.addExportButton();
          });

     }(Highcharts));
