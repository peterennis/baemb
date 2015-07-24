/// StartSnippet: WindowResizedHandlerSnippet
Canvas.width = window.innerWidth;
Canvas.height = window.innerHeight;


/// StartSnippet: ChartClickHandlerSnippet
var points = window.myLine.getPointsAtEvent(event);
if (points[0] != null) {
	bembData = points[0].value + ', ' + points[1].value;
} else {
	bembData = "";
}


/// StartSnippet: RefreshChartDataBuild
refreshChart({
	labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
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
	animation: true, 
 	responsive: true, 
 	bezierCurve: true, 
 	pointHitDetectionRadius: 5, 
 	pointDotRadius: 5, 
 	datasetStrokeWidth: 1, 
 	bezierCurveTension: 0.4 
});
