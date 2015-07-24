/* this is used by the bemb project */

var Canvas;

function refreshChart(chartData, options) {

	Canvas = document.getElementById("canvas");
	var ctx = Canvas.getContext("2d");

	ctx.canvas.width = 750;
	ctx.canvas.height = 500;
	window.myPie = new Chart(ctx).Pie(chartData, options);
	
}
