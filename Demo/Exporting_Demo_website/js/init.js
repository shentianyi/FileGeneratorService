var chart;
var containers = ['container_1', 'container_2', 'container_3'];
$(document).ready(function () {
    var charts = [];
    for (var i = 0; i < containers.length; i++) {
        var chart = new Highcharts.Chart({
            chart: {
                renderTo: containers[i]
            },
            title: {
                text: 'Monthly Average Temperature' + i,
                x: -20 //center
            },
            subtitle: {
                text: 'Source: WorldClimate.com' + i,
                x: -20
            },
            xAxis: {
                categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            },
            yAxis: {
                title: {
                    text: 'Temperature (°C)'
                },
                plotLines: [{
                    value: 0,
                    width: 1,
                    color: '#808080'
                }]
            },
            tooltip: {
                formatter: function () {
                    return '<b>' + this.series.name + '</b><br/>' + this.x + ': ' + this.y + '°C';
                }
            },
            legend: {
                layout: 'vertical',
                align: 'right',
                verticalAlign: 'top',
                x: -10,
                y: 100,
                borderWidth: 0
            },
            exporting: {
                url: 'HighchartsExport.axd',
                filename: 'MyChart',
                width: 1200,
                exportTypes: ['chart', 'png', 'jpeg', 'pdf', 'svg', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
            },
            series: [{
                name: 'Tokyo' + i,
                data: [17.0 * (i + 1), 6.9 * (i + 1), 9.5, 14.5 * (i + 1), 18.2 * (i + 1), 21.5 * (i + 1), 25.2 * (i + 1), 26.5, 23.3, 18.3, 13.9, 9.6]
            }, {
                name: 'New York' + i,
                data: [-0.2, 0.8 * (i + 1), 5.7, 11.3 * (i + 1), 17.0 * (i + 1), 22.0, 24.8, 24.1, 20.1, 14.1, 8.6 * (i + 1), 2.5]
            }, {
                name: 'Berlin' + i,
                data: [-0.9, 0.6 * (i + 1), 3.5 * (i + 1), 8.4, 13.5 * (i + 1), 17.0 * (i + 1), 18.6 * (i + 1), 17.9 * (i + 1), 14.3, 9.0, 3.9, 1.0]
            }, {
                name: 'London' + i,
                data: [3.9 * (i + 1), 4.2, 5.7, 8.5 * (i + 1), 11.9, 15.2, 17.0, 16.6 * (i + 1), 14.2 * (i + 1), 10.3, 6.6, 9.8 * (i + 1)]
            }]
        });
        charts.push(chart);
    }

    $("#exportPDFs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'pdf',
            charts: charts
        });
    });
    $("#exportDOCs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'doc',
            width: 600,
            charts: charts
        });
    });
    $("#exportDOCXs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'docx',
            width: 600,
            charts: charts
        });
    });
    $("#exportXLSs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'xls',
            width: 600,
            charts: charts
        });
    });
    $("#exportXLSXs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'xlsx',
            width: 1000,
            charts: charts
        });
    });

    $("#exportPPTXs_btn").click(function () {
        Highcharts.exportCharts({
            url: 'HighchartsExport.axd',
            type: 'pptx',
            width: 700,
            charts: charts
        });
    });
});
