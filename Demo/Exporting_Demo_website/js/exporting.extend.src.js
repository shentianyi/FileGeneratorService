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
						type : value.type
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
				if (exportingOptions.enableExtendExport) {
					extend(buttons.contextButton, {
						menuItems : []
					});
					for (var i = 0; i < exportingOptions.exportTypes.length; i++) {
						buttons.contextButton.menuItems.push(defaultExportButtons[exportingOptions.exportTypes[i]]);
					}
				}
			}
		});
		Highcharts.exportCharts = function(options, chartOptions) {
			options = options || {};
			var charts = options.charts;
			var svgs = [];
			var chartExportingOptions = charts[0].options.exporting;
			for (var i = 0; i < charts.length; i++) {
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
				svg : svgs
			});
		};
		Chart.prototype.callbacks.unshift(function(chart) {
			chart.addExportButton();
		});

	}(Highcharts));
