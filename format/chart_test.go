package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewChart(t *testing.T) {
	target :=  Chart{
		Type: "pie",
		Series:[]ChartSeries{
			{
				Name:       "=Sheet1!$A$30",
				Categories: "=Sheet1!$B$29:$D$29",
				Values:     "=Sheet1!$B$30:$D$30",
			},
		},
		Format: Picture {
			XScale:1.0,
			YScale:1.0,
			OffsetX:15,
			OffsetY:10,
			FPrintsWithSheet: true,
			NoChangeAspect: false,
			FLocksWithSheet: false,
		},
		Legend: ChartLegend {
			Position: "bottom",
			ShowLegendKey: false,
		},
		Title: ChartTitle {
			Name: "Fruit Pie Chart",
		},
		Plot:PlotArea{
			ShowBubbleSize:true,
			ShowCatName:false,
			ShowLeaderLines:false,
			ShowPercent:true,
			ShowSerName:false,
			ShowVal:false,
		},
		ShowBlanksAs:"gap",
	}

	//string
	s, err := NewChart(`{"type":"pie","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Pie Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"gap"}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Chart{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewChart(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Chart{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewChart(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Chart{}, s)
	require.Equal(t, &target, s)
}

