package main

import (
	"fmt"

	"github.com/ErmaiSoft/GoOpenXml/word"
)

func main() {
	titleFont := word.Font{Family: "微软雅黑", Size: 28, Bold: true, Color: "CC0000"}
	normalFont := word.Font{Family: "宋体", Size: 10.5, Bold: false, Space: true, Color: "000000"}
	normalBoldFont := word.Font{Family: "宋体", Size: 10.5, Bold: true, Space: true, Color: "000000"}
	normalLine := word.Line{Height: 1.5, Rule: word.LineRuleExact}
	subTitleFont := word.Font{Family: "微软雅黑", Size: 15, Bold: true, Color: "000000"}
	subTitleLine := word.Line{Rule: word.LineRuleAuto, Height: 1.5}
	contentFont := word.Font{Family: "宋体", Size: 10.5, Bold: false, Color: "000000"}
	contentLine := word.Line{Rule: word.LineRuleAuto, FirstLineChars: 2, Height: 1.5}

	docx := word.CreateDocx()
	docx.AddParagraph([]word.Paragraph{
		{
			F: titleFont,
			L: word.Line{After: 0.8, Rule: word.LineRuleAuto},
			T: []word.Text{
				{T: "会议纪要", F: &titleFont},
			},
			Rect: []*word.DrawRect{
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 1, PH: 0, PV: 16, T: "line", C: "CC0000"},
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 0.25, PH: 0, PV: 17, T: "line", C: "CC0000"},
			},
		},
		{
			F: normalFont,
			L: normalLine,
			T: []word.Text{
				{T: "时间"},
				{T: " | 2021-02-02 15:30", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			L: normalLine,
			T: []word.Text{
				{T: "地点"},
				{T: " | 二楼会议室", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			L: normalLine,
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
			L: normalLine,
			T: []word.Text{
				{T: "出席"},
				{T: " | 张三、李四、王五", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			L: normalLine,
			T: []word.Text{
				{T: "缺席"},
				{T: " | 麻六", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			L: normalLine,
			T: []word.Text{
				{T: "列席"},
				{T: " | 周吴", F: &normalBoldFont},
			},
		},
		{
			F: normalFont,
			L: normalLine,
			T: []word.Text{
				{T: "主持人"},
				{T: " | 李总", F: &normalBoldFont},
			},
		},
		{
			F: word.Font{Family: "宋体", Size: 10.5, Color: "000000"},
			L: word.Line{Height: 1.5, After: 1, Rule: word.LineRuleAuto},
			T: []word.Text{
				{T: "记录人"},
				{T: " | 张三", F: &normalBoldFont},
			},
			Rect: []*word.DrawRect{
				{W: float64(docx.PageSize.W - docx.PageSize.L - docx.PageSize.R), H: 1, PH: 0, PV: 8, T: "line", C: "CC0000"},
			}},
		{
			F: subTitleFont,
			L: subTitleLine,
			T: []word.Text{
				{T: "会议议题", F: &subTitleFont},
			},
		},
		{
			F: contentFont,
			L: contentLine,
			T: []word.Text{
				{T: "2012年11月15日，刚刚当选中共中央总书记的习近平同中外记者见面。总书记的话语温暖人心：“我们的人民热爱生活，期盼有更好的教育、更稳定的工作、更满意的收入、更可靠的社会保障、更高水平的医疗卫生服务、更舒适的居住条件、更优美的环境，期盼着孩子们能成长得更好、工作得更好、生活得更好。人民对美好生活的向往，就是我们的奋斗目标。”",
					F: &contentFont},
			},
		},
		{
			F: contentFont,
			L: contentLine,
			T: []word.Text{
				{T: "八年多来，以习近平同志为核心的党中央坚持以人民为中心的发展思想，顺应人民群众对美好生活的向往，统筹做好各领域民生工作，在学有所教、劳有所得、病有所医、老有所养、住有所居等方面取得新进展，在社会治理方面取得新成就。",
					F: &contentFont},
			},
		},
		{
			F: contentFont,
			L: contentLine,
			T: []word.Text{
				{T: "2020年10月29日，党的十九届五中全会审议通过了《中共中央关于制定国民经济和社会发展第十四个五年规划和二〇三五年远景目标的建议》。规划建议用12个部分展现“十四五”发展的重大任务，其中之一是“改善人民生活品质，提高社会建设水平”。",
					F: &contentFont},
			},
		},
	})

	err := docx.WriteToFile("D:/Dev/Demo/Word/word_test.docx")
	if err != nil {
		fmt.Println(err)
	}
}
