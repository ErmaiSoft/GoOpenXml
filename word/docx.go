package word

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
)

const rIDIndex = int64(7)

const (
	DocumentRelsFileKey = "word/_rels/document.xml.rels"
	DocumentFileKey     = "word/document.xml"
	ImagePath           = "media/"
)

const (
	ppi   float64 = 1440  //每英寸1440像素点
	mmpi  float64 = 25.4  //每英寸等于25.4毫米
	mmEmu float64 = 36000 //绘图像素，每毫米等于36000点
)

type LineRule string

const (
	LineRuleAuto  LineRule = "auto"
	LineRuleExact LineRule = "exact"
)

//DocxFile word文档
type DocxFile struct {
	Document     *W
	Medias       map[string][]byte
	XMLFiles     map[string][]byte
	DocumentRels *DocumentRels
}

//NewDocx 新建文档
func (d *Docx) newDocxFile() *DocxFile {
	xmlFile := make(map[string][]byte)
	xmlFile["[Content_Types].xml"] = []byte(XMLHeader + tplContentTypeXML)
	xmlFile["_rels/.rels"] = []byte(XMLHeader + tplRelsXML)
	xmlFile["docProps/app.xml"] = []byte(XMLHeader + tplAppXML)
	xmlFile["docProps/core.xml"] = []byte(XMLHeader + tplCoreXML)
	xmlFile["docProps/custom.xml"] = []byte(XMLHeader + tplCustomXML)
	xmlFile["word/theme/theme1.xml"] = []byte(XMLHeader + tplThemeXML)
	xmlFile["word/endnotes.xml"] = []byte(XMLHeader + tplEndnotesXML)
	xmlFile["word/fontTable.xml"] = []byte(XMLHeader + tplFontTableXML)
	xmlFile["word/footnotes.xml"] = []byte(XMLHeader + tplFootNotesXML)
	xmlFile["word/settings.xml"] = []byte(XMLHeader + tplSettingsXML)
	xmlFile["word/styles.xml"] = []byte(XMLHeader + tplStylesXML)
	xmlFile["word/webSettings.xml"] = []byte(XMLHeader + tplWebSettingsXML)

	var rels DocumentRels
	err := xml.Unmarshal([]byte(tplDocumentRelsXML), &rels)
	if err != nil {
		fmt.Println(err)
	}

	f := DocxFile{
		Document: &W{
			Wpc:       "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
			MC:        "http://schemas.openxmlformats.org/markup-compatibility/2006",
			O:         "urn:schemas-microsoft-com:office:office",
			R:         "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			M:         "http://schemas.openxmlformats.org/officeDocument/2006/math",
			V:         "urn:schemas-microsoft-com:vml",
			WP14:      "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
			WP:        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
			W10:       "urn:schemas-microsoft-com:office:word",
			W:         "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			W14:       "http://schemas.microsoft.com/office/word/2010/wordml",
			W15:       "http://schemas.microsoft.com/office/word/2012/wordml",
			WPG:       "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
			WPI:       "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
			WNE:       "http://schemas.microsoft.com/office/word/2006/wordml",
			WPS:       "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
			Ignorable: "w14 w15 wp14",
			Body:      &Body{}},
		XMLFiles:     xmlFile,
		DocumentRels: &rels,
	}
	return &f
}

//Docx word
type Docx struct {
	Paragraphs []Paragraph
	PageSize   PageSize
}

//Paragraph 段落
type Paragraph struct {
	F     Font        //字体
	L     Line        //行间距
	T     []Text      //段落文本
	Image *DrawImage  //图片
	Rect  []*DrawRect //形状
}

//DrawRect 形状
type DrawRect struct {
	W  float64 //单位毫米
	H  float64 //单位毫米
	PH float64 //单位毫米
	PV float64 //单位毫米
	C  string
	T  string //"line" "rect"
}

//DrawImage 图片
type DrawImage struct {
	W  float64 //单位毫米
	H  float64 //单位毫米
	PH float64 //单位毫米
	PV float64 //单位毫米
}

//PageSize 页面大小、页边距
type PageSize struct {
	W      float64 //页宽，单位毫米，默认A4，210
	H      float64 //页高，单位毫米，默认A4，297
	T      float64
	R      float64
	B      float64
	L      float64
	Header float64
	Footer float64
}

//Text 文本
type Text struct {
	T string
	F *Font
}

//Font 字体
type Font struct {
	Family string  //字体
	Size   float64 //字号
	Color  string  //字体颜色
	Bold   bool    //是否加粗
	Align  string  //left, right, top, bottom
	Space  bool    //xml:space="preserve" 包含空格
}

//Line 行高
type Line struct {
	Rule           LineRule //auto exact
	Height         float64  //行间距，1倍为1.0, 1.5倍1.5，2倍2.0
	FirstLineChars int64    //首行缩进字数，按汉字字宽算
	Before         float64  //上边间距，1倍为1.0, 1.5倍1.5，2倍2.0
	After          float64  //下边间距，1倍为1.0, 1.5倍1.5，2倍2.0
}

//CreateDocx 创建空白文档
func CreateDocx() *Docx {
	return &Docx{
		PageSize: PageSize{
			W:      210,
			H:      297,
			T:      25.4,
			R:      31.8,
			B:      25.4,
			L:      31.8,
			Header: 16,
			Footer: 16,
		},
	}
}

//AddParagraph 增加段落
func (d *Docx) getPageWidth() int64 {
	return int64(d.PageSize.W / 25.4 * 1440)
}

//AddParagraph 增加段落
func (d *Docx) AddParagraph(p []Paragraph) {
	d.Paragraphs = append(d.Paragraphs, p...)
}

func (f *Font) getFontSize() int64 {
	if f.Size > 0 {
		return int64(f.Size * 2)
	}
	return 21 //默认五号
}

func (p *Paragraph) getLineHeight() int64 {
	if p.L.Height != 0 {
		return int64(math.Floor(240 * p.L.Height)) //行距默认240个点
	}
	return 240 //默认五号，1.5倍行距
}

func (p *Paragraph) getRule() LineRule {
	if len(p.L.Rule) > 0 {
		return p.L.Rule
	}
	return LineRuleAuto
}

func (p *Paragraph) getBefore() int64 {
	return int64(math.Floor(240 * p.L.Before))
}

func (p *Paragraph) getAfter() int64 {
	return int64(math.Floor(240 * p.L.After))
}

func (p *Paragraph) getAlign() string {
	if len(p.F.Align) > 0 {
		return p.F.Align
	}
	return "left"
}

func (p *Paragraph) getFirstLineChars() int64 {
	return 100 * p.L.FirstLineChars
}

// WriteToFile 写到文件中
func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

// Write Write
func (d *Docx) Write(ioWriter io.Writer) (err error) {
	df := d.newDocxFile()

	w := zip.NewWriter(ioWriter)
	for name, content := range df.XMLFiles {
		var writer io.Writer
		writer, err = w.Create(name)
		if err != nil {
			return err
		}

		writer.Write([]byte(content))
	}

	{
		content, _, err := d.writeDocContent()
		var writer io.Writer
		writer, err = w.Create(DocumentFileKey)
		if err != nil {
			return err
		}
		writer.Write([]byte(XMLHeader))
		writer.Write([]byte(content))
	}

	w.Close()
	return
}

func (d *Docx) writeDocContent() (content []byte, rels []byte, err error) {
	w := W{
		Wpc:       "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
		MC:        "http://schemas.openxmlformats.org/markup-compatibility/2006",
		O:         "urn:schemas-microsoft-com:office:office",
		R:         "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		M:         "http://schemas.openxmlformats.org/officeDocument/2006/math",
		V:         "urn:schemas-microsoft-com:vml",
		WP14:      "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
		WP:        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		W10:       "urn:schemas-microsoft-com:office:word",
		W:         "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		W14:       "http://schemas.microsoft.com/office/word/2010/wordml",
		W15:       "http://schemas.microsoft.com/office/word/2012/wordml",
		WPG:       "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
		WPI:       "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
		WNE:       "http://schemas.microsoft.com/office/word/2006/wordml",
		WPS:       "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
		Ignorable: "w14 w15 wp14",
		Body:      &Body{},
	}

	rectIdx := int64(1)

	for _, paragraph := range d.Paragraphs {
		p := P{}
		rPr := RPr{
			RFonts: &RFonts{
				ASCII:    paragraph.F.Family,
				EastAsia: paragraph.F.Family,
				HAnsi:    paragraph.F.Family,
			},
			Color: &Color{Val: paragraph.F.Color},
			Sz:    &Sz{Val: paragraph.F.getFontSize()},
			SzCs:  &SzCs{Val: paragraph.F.getFontSize()},
		}
		if paragraph.F.Bold {
			rPr.B = &Bold{Val: "on"}
			// rPr.BCs = " "
		}

		p.PPr = &PPr{
			SnapToGrid: &SnapToGrid{Val: "0"},
			Spacing: &Spacing{
				Before:   paragraph.getBefore(),
				After:    paragraph.getAfter(),
				Line:     paragraph.getLineHeight(),
				LineRule: paragraph.getRule(),
			},
			Jc:  &Jc{Val: paragraph.getAlign()},
			RPr: &rPr,
		}

		if paragraph.L.FirstLineChars > 0 {
			p.PPr.Ind = &Ind{
				FirstLineChars: paragraph.getFirstLineChars(),
			}
		}

		//文本块
		for _, t := range paragraph.T {
			r := &R{
				T: &T{Text: t.T},
			}
			if t.F != nil {
				r.RPr = &RPr{
					RFonts: &RFonts{
						ASCII:    t.F.Family,
						EastAsia: t.F.Family,
						HAnsi:    t.F.Family,
					},
					Color: &Color{Val: t.F.Color},
					Sz:    &Sz{Val: t.F.getFontSize()},
					SzCs:  &SzCs{Val: t.F.getFontSize()},
				}
				if t.F.Bold {
					r.RPr.B = &Bold{Val: "on"}
					// r.RPr.BCs = " "
				}

				if t.F.Space {
					r.T.Space = "preserve"
				}

			}
			p.Rs = append(p.Rs, r)
		}

		//形状
		for _, r := range paragraph.Rect {
			p.Rs = append(p.Rs, &R{
				Drawing: &Drawing{
					//Anchor 形状
					Anchor: &Anchor{
						DistT:          0,
						DistB:          0,
						DistL:          0,
						DistR:          0,
						SimplePos:      0,
						RelativeHeight: 0,
						BehindDoc:      0,
						Locked:         0,
						LayoutInCell:   1,
						AllowOverlap:   1,
						AnchorID:       "69E31D9A",
						EditID:         "48F3AB62",
						WpSimplePos:    &SimplePos{X: 0, Y: 0},
						PositionH: &PositionH{
							RelativeFrom: "column",
							PosOffset:    &PosOffset{Text: fmt.Sprintf("%d", millMeterToEMU(r.PH))},
						},
						PositionV: &PositionV{
							RelativeFrom: "paragraph",
							PosOffset:    &PosOffset{Text: fmt.Sprintf("%d", millMeterToEMU(r.PV))},
						},
						Extent:            &Extent{CX: millMeterToEMU(r.W), CY: 0},
						EffectExtent:      &EffectExtent{},
						WrapNone:          "",
						DocPr:             &DocPr{ID: rectIdx, Name: fmt.Sprintf("%s%d", PrefixRID, rectIdx)},
						CNvGraphicFramePr: &CNvGraphicFramePr{},
						Graphic: &Graphic{
							A: "http://schemas.openxmlformats.org/drawingml/2006/main",
							GraphicData: &GraphicData{
								URI: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
								Wsp: &Wsp{
									CNvCnPr: "",
									WpsSpPr: &WpsSpPr{
										Xfrm: &Xfrm{
											FlipV: 1,
											AOff:  &AOff{X: 0, Y: 0},
											AExt:  &AExt{CX: millMeterToEMU(r.W), CY: 0},
										},
										PrstGeom: &PrstGeom{
											Prst:  r.T, //形状
											AVLst: "",
										},
										Ln: &Ln{
											W: millMeterToEMU(r.H), //线高
											SolidFill: &SolidFill{
												SrgbClr: &SrgbClr{Val: r.C}, //颜色
											},
										},
									},
									BodyPr: "",
								},
							},
						},
					},
				},
			})

			rectIdx++
		}

		w.Body.AddSect(&p)
	}

	content, err = xml.Marshal(w)
	if err != nil {
		fmt.Println(err)
	}

	return
}

//毫米转像素
func millMeterToPixel(mm float64) int64 {
	return int64(mm / mmpi * ppi)
}

//毫米转绘图像素
func millMeterToEMU(mm float64) int64 {
	return int64(mm * mmEmu)
}
