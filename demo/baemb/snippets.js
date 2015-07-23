/// StartSnippet: WindowResizedHandlerSnippet
Canvas.width = 400;						//window.innerWidth;
Canvas.height = 400;					//window.innerHeight;


/// StartSnippet: ChartClickHandlerSnippet
var points = window.myLine.getPointsAtEvent(event);
if (points[0] != null) {
	bembData = points[0].value + ', ' + points[1].value;
} else {
	bembData = "";
}


/// StartSnippet: RefreshChartDataBuild
refreshChart({
	labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
	datasets: [{
		label: 'Dataset 1 Name',
		fillColor: 'rgba(220,220,220,0.2)',
		strokeColor: 'rgba(220,220,220,1)',
		pointStrokeColor: '#fff',
		pointHighlightFill: '#fff',
		pointHighlightStroke: 'rgba(220,220,220,1)',
		data: [|DATASET1|]
	}, {
		label: 'Dataset 2 Name',
		fillColor: 'rgba(151,187,205,0.2)',
		strokeColor: 'rgba(151,187,205,1)',
		pointStrokeColor: '#fff',
		pointHighlightFill: '#fff',
		pointHighlightStroke: 'rgba(151,187,205,1)',
		data: [|DATASET2|]
	}]
}, {
	animation: false,
	responsive: true,
	scaleOverride: false,
	//scaleOverride: true,
	// ** Required if scaleOverride is true **
    // Number - The number of steps in a hard coded scale
    //scaleSteps: 10,
    // Number - The value jump in the hard coded scale
    //scaleStepWidth: 25,
    // Number - The scale starting value
    //scaleStartValue: 0,
	scaleBeginAtZero: true,
	maintainAspectRatio: false,
	bezierCurve: true,
	pointHitDetectionRadius: 5,
	pointDotRadius: 5,
	datasetStrokeWidth: 1,
	bezierCurveTension: 0.4
});
