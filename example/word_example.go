package main

import (
	"fmt"

	"github.com/ErmaiSoft/GoOpenXml/word"
)

func main() {
	titleFont := word.Font{Font: "微软雅黑", FontSize: 28, Bold: true, FontColor: "CC0000"}
	normalFont := word.Font{Font: "宋体", FontSize: 10.5, Bold: false, Space: true, FontColor: "000000"}
	normalBoldFont := word.Font{Font: "宋体", FontSize: 10.5, Bold: true, Space: true, FontColor: "000000"}
	normalLine := word.Line{LineHeight: 1.5, LineRule: word.LineRuleExact}
	subTitleFont := word.Font{Font: "微软雅黑", FontSize: 15, Bold: true, FontColor: "000000"}
	contentLine := word.Line{LineRule: word.LineRuleAuto}
	contentFont := word.Font{Font: "宋体", FontSize: 10.5, Bold: false, FontColor: "000000"}

	docx := word.CreateDocx()
	docx.AddParagraph([]word.Paragraph{
		{
			F: titleFont,
			S: word.Line{After: 0.5, LineRule: word.LineRuleExact},
			T: []word.Text{
				{T: "会议纪要", F: &titleFont},
			},
			Rect: []*word.DrawRect{
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 1, PH: 0, PV: 12, T: "line", C: "CC0000"},
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 0.25, PH: 0, PV: 13, T: "line", C: "CC0000"},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "时间"},
				{T: " | 2021-02-02 15:30", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "地点"},
				{T: " | 二楼会议室", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "应到"},
				{T: " | 25 人", F: &normalBoldFont},
				{T: "    实到", F: &normalFont},
				{T: " | 22 人", F: &normalBoldFont},
				{T: "    参会率", F: &normalFont},
				{T: " | 88%", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "出席"},
				{T: " | 张三、李四、王五", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "缺席"},
				{T: " | 麻六", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "列席"},
				{T: " | 周吴", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			S: normalLine,
			T: []word.Text{
				{T: "主持人"},
				{T: " | 李总", F: &normalBoldFont},
			},
		},
		{
			F: word.Font{Font: "宋体", FontSize: 10.5, FontColor: "000000"},
			S: word.Line{LineHeight: 1.5, After: 1, LineRule: word.LineRuleExact},
			T: []word.Text{
				{T: "记录人"},
				{T: " | 张三", F: &normalBoldFont},
			},
			Rect: []*word.DrawRect{
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 1, PH: 0, PV: 12, T: "line", C: "CC0000"},
			}},
		{
			F: subTitleFont,
			S: contentLine,
			T: []word.Text{
				{T: "会议议题", F: &subTitleFont},
			},
		},
		{
			F: contentFont,
			S: contentLine,
			T: []word.Text{
				{T: "关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前",
					F: &contentFont},
			},
		},
		{
			F: contentFont,
			S: contentLine,
			T: []word.Text{
				{T: "关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前关于当前",
					F: &contentFont},
			},
		},
	})

	err := docx.WriteToFile("D:/Dev/Demo/Word/word_test.docx")
	if err != nil {
		fmt.Println(err)
	}
}
