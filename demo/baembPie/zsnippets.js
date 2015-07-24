/// StartSnippet: WindowResizedHandlerSnippet
Canvas.width = window.innerWidth;
Canvas.height = window.innerHeight;


/// StartSnippet: ChartClickHandlerSnippet
var points = window.Pie.getPointsAtEvent(event);
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
	//Boolean - Whether we should show a stroke on each segment
    segmentShowStroke : true,
    //String - The colour of each segment stroke
    segmentStrokeColor : "#fff",
    //Number - The width of each segment stroke
    segmentStrokeWidth : 2,
    //Number - The percentage of the chart that we cut out of the middle
    percentageInnerCutout : 0, // This is 0 for Pie charts
    //Number - Amount of animation steps
    animationSteps : 100,
    //String - Animation easing effect
    animationEasing : "easeOutBounce",
    //Boolean - Whether we animate the rotation of the Doughnut
    animateRotate : true,
    //Boolean - Whether we animate scaling the Doughnut from the centre
    animateScale : false,
    //String - A legend template
    //legendTemplate : "<ul class=\"<%=name.toLowerCase()%>-legend\"><% for (var i=0; i<segments.length; i++){%><li><span style=\"background-color:<%=segments[i].fillColor%>\"></span><%if(segments[i].label){%><%=segments[i].label%><%}%></li><%}%></ul>"

});
