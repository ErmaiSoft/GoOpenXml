package word

import "encoding/xml"

//RelationshipTypeImage 图片文档映射关系类型
const RelationshipTypeImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

//PrefixRID rId前缀
const PrefixRID = "rId"

//W word文档
type W struct {
	XMLName   xml.Name `xml:"w:document"`
	Wpc       string   `xml:"xmlns:wpc,attr"`    //"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
	MC        string   `xml:"xmlns:mc,attr"`     //"http://schemas.openxmlformats.org/markup-compatibility/2006"
	O         string   `xml:"xmlns:o,attr"`      //"urn:schemas-microsoft-com:office:office"
	R         string   `xml:"xmlns:r,attr"`      //"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	M         string   `xml:"xmlns:m,attr"`      //"http://schemas.openxmlformats.org/officeDocument/2006/math"
	V         string   `xml:"xmlns:v,attr"`      //"urn:schemas-microsoft-com:vml"
	WP14      string   `xml:"xmlns:wp14,attr"`   //"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
	WP        string   `xml:"xmlns:wp,attr"`     //"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	W10       string   `xml:"xmlns:w10,attr"`    //"urn:schemas-microsoft-com:office:word"
	W         string   `xml:"xmlns:w,attr"`      //"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	W14       string   `xml:"xmlns:w14,attr"`    //"http://schemas.microsoft.com/office/word/2010/wordml
	W15       string   `xml:"xmlns:w15,attr"`    //"http://schemas.microsoft.com/office/word/2012/wordml"
	WPG       string   `xml:"xmlns:wpg,attr"`    //"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
	WPI       string   `xml:"xmlns:wpi,attr"`    //"http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
	WNE       string   `xml:"xmlns:wne,attr"`    //"http://schemas.microsoft.com/office/word/2006/wordml"
	WPS       string   `xml:"xmlns:wps,attr"`    //"http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
	Ignorable string   `xml:"mc:Ignorable,attr"` //"w14 w15 wp14"
	Body      *Body    `xml:"w:body"`
}

//SetBody 添加文档主体
func (w *W) SetBody(b *Body) {
	w.Body = b
}

//Body Word文档主体
type Body struct {
	Sects []Sect
}

//AddSect 增加章节
func (b *Body) AddSect(s Sect) {
	b.Sects = append(b.Sects, s)
}

//Sect 章节
type Sect interface {
	Sfunc()
}

//P 段落
type P struct {
	XMLName      xml.Name `xml:"w:p"`
	RsidR        string   `xml:"w:rsidR,attr,omitempty"`
	RsidRDefault string   `xml:"w:rsidRDefault,attr,omitempty"`
	PPr          *PPr     `xml:"w:pPr,omitempty"`
	Rs           []Run    `xml:"w:r,omitempty"`
}

//Sfunc P sect
func (p *P) Sfunc() {
}

//Run 段落内容
type Run interface {
	Rfunc()
}

//PPr 段落属性
type PPr struct {
	XMLName    xml.Name    `xml:"w:pPr"`
	SnapToGrid *SnapToGrid `xml:"w:snapToGrid,omitempty"`
	Spacing    *Spacing    `xml:"w:spacing,omitempty"`
	Ind        *Ind        `xml:"w:ind,omitempty"`
	Jc         *Jc         `xml:"w:jc,omitempty"`
	RPr        *RPr        `xml:"w:rPr,omitempty"`
}

//Ind 首行：缩进、行高
type Ind struct {
	XMLName        xml.Name `xml:"w:ind"`
	FirstLineChars int64    `xml:"w:firstLineChars,attr"` //首行缩进字符数，100是一个字符
	LeftChars      int64    `xml:"w:leftChars,attr"`      //左缩进字符数，100是一个字符
	RightChars     int64    `xml:"w:rightChars,attr"`     //右缩进字符数，100是一个字符
	// FirstLine      int64    `xml:"w:firstLine,attr"`
}

// <w:ind w:firstLineChars="200" w:firstLine="420" />

//R 块
type R struct {
	XMLName xml.Name `xml:"w:r"`
	RPr     *RPr     `xml:"w:rPr,omitempty"`     //块属性
	T       *T       `xml:"w:t,omitempty"`       //文字
	Drawing *Drawing `xml:"w:drawing,omitempty"` //图片
}

//Rfunc run sect
func (r *R) Rfunc() {
}

//SnapToGrid 对齐网格
type SnapToGrid struct {
	XMLName xml.Name `xml:"w:snapToGrid"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

//RPr 文字块属性
type RPr struct {
	XMLName xml.Name `xml:"w:rPr"`
	RFonts  *RFonts  `xml:"w:rFonts,omitempty"`
	B       *Bold    `xml:"w:b,omitempty"`
	// BCs     string   `xml:"w:bCs,omitempty"`
	Color *Color `xml:"w:color"`
	Sz    *Sz    `xml:"w:sz"`
	SzCs  *SzCs  `xml:"w:szCs"`
}

//Bold 加粗
type Bold struct {
	XMLName xml.Name `xml:"w:b"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

//RFonts <w:rFonts w:ascii="微软雅黑" w:eastAsia="微软雅黑" w:hAnsi="微软雅黑"/>
type RFonts struct {
	XMLName  xml.Name `xml:"w:rFonts"`
	ASCII    string   `xml:"w:ascii,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	HAnsi    string   `xml:"w:hAnsi,attr,omitempty"`
}

//Color <w:color w:val="3A3838"/>
type Color struct {
	XMLName xml.Name `xml:"w:color"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

//Sz 字号大小，如：14号字为28，Word上的字号乘以2 <w:sz w:val="56"/>
type Sz struct {
	XMLName xml.Name `xml:"w:sz"`
	Val     int64    `xml:"w:val,attr,omitempty"`
}

//SzCs 未知 <w:szCs w:val="56"/>
type SzCs struct {
	XMLName xml.Name `xml:"w:szCs"`
	Val     int64    `xml:"w:val,attr,omitempty"`
}

// Spacing 行间距
// <w:spacing w:before="360" w:after="120" w:line="480" w:lineRule="auto" w:beforeAutospacing="0" w:afterAutospacing="0"/>
// http://officeopenxml.com/WPspacing.php
// Values are in twentieths of a point. A normal single-spaced paragaph has a w:line value of 240, or 12 points.
// To specify units in hundreths of a line, use attributes 'afterLines'/'beforeLines'.
// The space between adjacent paragraphs will be the greater of the 'line' spacing of each paragraph, the spacing
// after the first paragraph, and the spacing before the second paragraph. So if the first paragraph specifies 240
// after and the second 80 before, and they are both single-spaced ('line' value of 240), then the space between
// the paragraphs will be 240.
// Specifies how the spacing between lines as specified in the line attribute is calculated.
// Note: If the value of the lineRule attribute is atLeast or exactly, then the value of the line attribute is interpreted as 240th of a point. If the value of lineRule is auto, then the value of line is interpreted as 240th of a line.
type Spacing struct {
	XMLName           xml.Name `xml:"w:spacing"`
	Before            int64    `xml:"w:before,attr,omitempty"`
	After             int64    `xml:"w:after,attr,omitempty"`
	Line              int64    `xml:"w:line,attr,omitempty"`
	LineRule          LineRule `xml:"w:lineRule,attr,omitempty"`
	BeforeAutospacing int64    `xml:"w:beforeAutospacing"`
	AfterAutospacing  int64    `xml:"w:afterAutospacing"`
}

//Jc 对齐方式 <w:jc w:val="left"/>
type Jc struct {
	XMLName xml.Name `xml:"w:jc"`
	Val     string   `xml:"w:val,attr,omitempty"`
}

//T 文本
type T struct {
	XMLName xml.Name `xml:"w:t"`
	Space   string   `xml:"xml:space,attr,omitempty"` //"preserve"
	// Space string `xml:"w:space,attr,omitempty"`
	Text string `xml:",chardata"`
}

//Drawing 绘图
type Drawing struct {
	XMLName xml.Name `xml:"w:drawing"`
	Inline  *Inline  `xml:"wp:inline,omitempty"` //插入图片
	Anchor  *Anchor  `xml:"wp:anchor,omitempty"` //插入形状
}

//Inline 绘图边框
type Inline struct {
	XMLName           xml.Name           `xml:"wp:inline"`
	DistT             int64              `xml:"distT,attr"`
	DistB             int64              `xml:"distB,attr"`
	DistL             int64              `xml:"distL,attr"`
	DistR             int64              `xml:"distR,attr"`
	Extent            *Extent            `xml:"wp:extent"`
	EffectExtent      *EffectExtent      `xml:"wp:effectExtent"`
	DocPr             *DocPr             `xml:"wp:docPr"`
	CNvGraphicFramePr *CNvGraphicFramePr `xml:"wp:cNvGraphicFramePr"`
	Graphic           *Graphic           `xml:"a:graphic"`
}

//Extent  绘图范围
type Extent struct {
	XMLName xml.Name `xml:"wp:extent"`
	CX      int64    `xml:"cx,attr"`
	CY      int64    `xml:"cy,attr"`
}

//EffectExtent 绘图有效范围
type EffectExtent struct {
	XMLName xml.Name `xml:"wp:effectExtent"`
	L       int64    `xml:"l,attr"` //左边距
	T       int64    `xml:"t,attr"` //上边距
	R       int64    `xml:"r,attr"` //右边距
	B       int64    `xml:"b,attr"` //下边距
}

//WrapNone 不断行
type WrapNone struct {
	XMLName xml.Name `xml:"wp:wrapNone"`
}

//DocPr 文档属性，唯一就行，好像没鸟用
type DocPr struct {
	XMLName xml.Name `xml:"wp:docPr"`
	ID      int64    `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

//CNvGraphicFramePr 图形框架属性
type CNvGraphicFramePr struct {
	XMLName           xml.Name           `xml:"wp:cNvGraphicFramePr"`
	GraphicFrameLocks *GraphicFrameLocks `xml:"a:graphicFrameLocks"`
}

//GraphicFrameLocks 图形框架锁
type GraphicFrameLocks struct {
	XMLName        xml.Name `xml:"a:graphicFrameLocks"`
	A              string   `xml:"xmlns:a,attr"` //"http://schemas.openxmlformats.org/drawingml/2006/main"
	NoChangeAspect int64    `xml:"noChangeAspect,attr"`
}

//Graphic 图形
type Graphic struct {
	XMLName     xml.Name     `xml:"a:graphic"`
	A           string       `xml:"xmlns:a,attr"` //"http://schemas.openxmlformats.org/drawingml/2006/main"
	GraphicData *GraphicData `xml:"a:graphicData"`
}

//GraphicData 图形数据
type GraphicData struct {
	XMLName xml.Name `xml:"a:graphicData"`
	//uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" 插入图片
	//uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 插入形状
	URI string `xml:"uri,attr"`
	Pic *Pic   `xml:"pic:pic,omitempty"` //图片
	Wsp *Wsp   `xml:"wps:wsp,omitempty"` //形状
}

//Pic 图形
type Pic struct {
	XMLName  xml.Name  `xml:"pic:pic"`
	NSPic    string    `xml:"xmlns:pic,attr"` //"http://schemas.openxmlformats.org/drawingml/2006/picture"
	NvPicPr  *NvPicPr  `xml:"pic:nvPicPr"`
	BlipFill *BlipFill `xml:"pic:blipFill"`
	PicSpPr  *PicSpPr  `xml:"pic:spPr"`
}

//NvPicPr pic:nvPicPr
type NvPicPr struct {
	XMLName  xml.Name `xml:"pic:nvPicPr"`
	CNvPr    *CNvPr   `xml:"pic:cNvPr"`
	CNvPicPr string   `xml:"pic:cNvPicPr"`
}

//CNvPr pic:cNvPr
type CNvPr struct {
	XMLName xml.Name `xml:"pic:cNvPr"`
	ID      int64    `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

//BlipFill 填充
type BlipFill struct {
	XMLName xml.Name `xml:"pic:blipFill"`
	Blip    *Blip    `xml:"a:blip"`
	Stretch *Stretch `xml:"a:stretch"`
}

//Blip a:blip
type Blip struct {
	XMLName xml.Name `xml:"a:blip"`
	Embed   string   `xml:"r:embed,attr"` //填充图片对应rel ID
}

//Stretch 拉伸
type Stretch struct {
	XMLName  xml.Name `xml:"a:stretch"`
	FillRect string   `xml:"a:fillRect"`
}

//PicSpPr pic:spPr
type PicSpPr struct {
	XMLName  xml.Name  `xml:"pic:spPr"`
	Xfrm     *Xfrm     `xml:"a:xfrm"`
	PrstGeom *PrstGeom `xml:"a:prstGeom"`
}

//Xfrm a:xfrm
type Xfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	FlipV   int64    `xml:"flipV,attr"`
	AOff    *AOff    `xml:"a:off"`
	AExt    *AExt    `xml:"a:ext"`
}

//AOff a:off
type AOff struct {
	XMLName xml.Name `xml:"a:off"`
	X       int64    `xml:"x,attr"`
	Y       int64    `xml:"y,attr"`
}

//AExt "a:ext
type AExt struct {
	XMLName xml.Name `xml:"a:ext"`
	CX      int64    `xml:"cx,attr"` //图片宽度，36000为1毫米
	CY      int64    `xml:"cy,attr"` //图片高度，36000为1毫米
}

//PrstGeom 几何形状，rect：矩形
type PrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom"`
	Prst    string   `xml:"prst,attr"`
	AVLst   string   `xml:"a:avLst"`
}

//Anchor 形状
type Anchor struct {
	XMLName           xml.Name           `xml:"wp:anchor"`
	DistT             int64              `xml:"distT,attr"`
	DistB             int64              `xml:"distB,attr"`
	DistL             int64              `xml:"distL,attr"`
	DistR             int64              `xml:"distR,attr"`
	SimplePos         int64              `xml:"simplePos,attr"`      //默认0
	RelativeHeight    int64              `xml:"relativeHeight,attr"` //默认0
	BehindDoc         int64              `xml:"behindDoc,attr"`      //默认0
	Locked            int64              `xml:"locked,attr"`         //默认0
	LayoutInCell      int64              `xml:"layoutInCell,attr"`   //默认1
	AllowOverlap      int64              `xml:"allowOverlap,attr"`   //默认1
	AnchorID          string             `xml:"wp14:anchorId,attr"`  //"69E31D9A"
	EditID            string             `xml:"wp14:editId,attr"`    //"48F3AB62"
	WpSimplePos       *SimplePos         `xml:"wp:simplePos"`
	PositionH         *PositionH         `xml:"wp:positionH"`
	PositionV         *PositionV         `xml:"wp:positionV"`
	Extent            *Extent            `xml:"wp:extent"`
	EffectExtent      *EffectExtent      `xml:"wp:effectExtent"`
	WrapNone          string             `xml:"wp:wrapNone"`
	DocPr             *DocPr             `xml:"wp:docPr"`
	CNvGraphicFramePr *CNvGraphicFramePr `xml:"wp:cNvGraphicFramePr"`
	Graphic           *Graphic           `xml:"a:graphic"`
}

//SimplePos wp:simplePos
type SimplePos struct {
	XMLName xml.Name `xml:"wp:simplePos"`
	X       int64    `xml:"x,attr"`
	Y       int64    `xml:"y,attr"`
}

//PositionH wp:positionH
type PositionH struct {
	XMLName      xml.Name   `xml:"wp:positionH"`
	RelativeFrom string     `xml:"relativeFrom,attr"` //column
	PosOffset    *PosOffset `xml:"wp:posOffset"`
}

//PositionV wp:positionV
type PositionV struct {
	XMLName      xml.Name   `xml:"wp:positionV"`
	RelativeFrom string     `xml:"relativeFrom,attr"` //paragraph
	PosOffset    *PosOffset `xml:"wp:posOffset"`
}

//PosOffset wp:posOffset
type PosOffset struct {
	XMLName xml.Name `xml:"wp:posOffset"`
	Text    string   `xml:",chardata"`
}

//Wsp word形状数据，wps:wsp Word Processing Shape
type Wsp struct {
	XMLName xml.Name `xml:"wps:wsp"`
	CNvCnPr string   `xml:"wps:cNvCnPr"`
	WpsSpPr *WpsSpPr `xml:"wps:spPr"`
	BodyPr  string   `xml:"wps:bodyPr"`
}

//WpsSpPr wps:spPr
type WpsSpPr struct {
	XMLName  xml.Name  `xml:"wps:spPr"`
	Xfrm     *Xfrm     `xml:"a:xfrm"`
	PrstGeom *PrstGeom `xml:"a:prstGeom"`
	Ln       *Ln       `xml:"a:ln"`
}

//Ln 线
type Ln struct {
	XMLName   xml.Name   `xml:"a:ln"`
	W         int64      `xml:"w,attr"`      //线宽
	SolidFill *SolidFill `xml:"a:solidFill"` //填充
}

//SolidFill 实心填充
type SolidFill struct {
	XMLName xml.Name `xml:"a:solidFill"`
	SrgbClr *SrgbClr `xml:"a:srgbClr"`
}

//SrgbClr 填充颜色
type SrgbClr struct {
	XMLName xml.Name `xml:"a:srgbClr"`
	Val     string   `xml:"val,attr"` //颜色 RGB
}

//WpsStyle 样式
type WpsStyle struct {
	XMLName   xml.Name   `xml:"wps:style"`
	LnRef     *LnRef     `xml:"a:lnRef"`
	FillRef   *FillRef   `xml:"a:fillRef"`
	EffectRef *EffectRef `xml:"a:effectRef"`
	FontRef   *FontRef   `xml:"a:fontRef"`
}

//LnRef a:lnRef
type LnRef struct {
	XMLName   xml.Name   `xml:"a:lnRef"`
	IDX       int64      `xml:"idx,attr"`
	SchemeClr *SchemeClr `xml:"a:schemeClr"`
}

//FillRef a:fillRef
type FillRef struct {
	XMLName   xml.Name   `xml:"a:fillRef"`
	IDX       int64      `xml:"idx,attr"`
	SchemeClr *SchemeClr `xml:"a:schemeClr"`
}

//EffectRef a:effectRef
type EffectRef struct {
	XMLName   xml.Name   `xml:"a:effectRef"`
	IDX       int64      `xml:"idx,attr"`
	SchemeClr *SchemeClr `xml:"a:schemeClr"`
}

//FontRef a:fontRef
type FontRef struct {
	XMLName   xml.Name   `xml:"a:fontRef"`
	IDX       int64      `xml:"idx,attr"`
	SchemeClr *SchemeClr `xml:"a:schemeClr"`
}

//SchemeClr a:schemeClr
type SchemeClr struct {
	XMLName xml.Name `xml:"a:schemeClr"`
	Val     string   `xml:"val,attr"`
}

//Relationship 文档映射关系
type Relationship struct {
	XMLName xml.Name `xml:"Relationship"`
	ID      string   `xml:"Id,attr"`     //rId9
	Target  string   `xml:"Target,attr"` //media/image2.png
	Type    string   `xml:"Type,attr"`   //"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
}

//DocumentRels 文档映射关系
type DocumentRels struct {
	XMLName       xml.Name        `xml:"Relationships"`
	Relationships []*Relationship `xml:"Relationship"`
}
