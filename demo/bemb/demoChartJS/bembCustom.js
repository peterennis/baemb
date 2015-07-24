/* this is used by the bemb project */

var Canvas;

function refreshChart(chartData, options) {

	Canvas = document.getElementById("canvas");
	var ctx = Canvas.getContext("2d");

	window.myLine = new Chart(ctx).Line(chartData, options);
	
}
