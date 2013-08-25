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

		var defaultExportButtonsData = [{
			key : 'png',
			textKey : 'downloadPNG',
			type : 'image/png'
		}, {
			key : 'jpeg',
			textKey : 'downloadJPEG',
			type : 'image/jpeg'
		}, {
			key : 'pdf',
			textKey : 'downloadPDF',
			type : 'application/pdf'
		}, {
			key : 'svg',
			textKey : 'downloadSVG',
			type : 'image/svg+xml'
		}, {
			key : 'doc',
			textKey : 'downloadDOC',
			type : 'application/msword'
		}, {
			key : 'docx',
			textKey : 'downloadDOCX',
			type : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
		}, {
			key : 'xls',
			textKey : 'downloadXLS',
			type : 'application/vnd.ms-excel'
		}, {
			key : 'xlsx',
			textKey : 'downloadXLSX',
			type : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
		}, {
			key : 'ppt',
			textKey : 'downloadPPT',
			type : 'application/vnd.ms-powerpoint'
		}, {
			key : 'pptx',
			textKey : 'downloadPPTX',
			type : 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
		}];
		var defaultExportButtons = {
			'chart' : {
				textKey : 'printChart',
				onclick : function() {
					this.print();
				}
			}
		};
		$.each(defaultExportButtonsData, function(index, value) {
			defaultExportButtons[value.key] = {
				textKey : value.textKey,
				onclick : function() {
					this.exportChart({type:value.type});
				}
			};
		});
		console.log(defaultExportButtons);
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
		Chart.prototype.callbacks.unshift(function(chart) {
			chart.addExportButton();
		});

	}(Highcharts));
