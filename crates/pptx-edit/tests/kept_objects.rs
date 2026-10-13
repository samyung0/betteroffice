//! Objects the editor does not model (charts, SmartArt, equations, video and
//! audio, transitions and animations) survive an edit to the slide that holds
//! them, spelled as PowerPoint writes them.

use std::collections::BTreeMap;

use pptx_edit::{DeckSession, EditCtx, TextStyle};

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNG: &[u8] = &[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4];

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="mp4" ContentType="video/mp4"/>
<Default Extension="m4a" ContentType="audio/mp4"/>
<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
<Override PartName="/ppt/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>
<Override PartName="/ppt/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>
<Override PartName="/ppt/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>
<Override PartName="/ppt/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

const PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#;

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>"#;

const TITLE: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="228600"/><a:ext cx="6096000" cy="685800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Cells divide</a:t></a:r></a:p></p:txBody></p:sp>"#;

const CHART_FRAME: &str = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="1143000"/><a:ext cx="4572000" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>"#;

const SMARTART_FRAME: &str = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram 3"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="5486400" y="1143000"/><a:ext cx="4572000" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId3" r:lo="rId4" r:qs="rId5" r:cs="rId6"/></a:graphicData></a:graphic></p:graphicFrame>"#;

/// PowerPoint's spelling of an inserted equation: the math in an `a14`
/// choice and a locked text box filled with its rendered picture as fallback.
const EQUATION: &str = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath><m:sSup><m:sSupPr><m:ctrlPr><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>x</m:t></m:r></m:e><m:sup><m:r><a:rPr lang="en-US" i="1"><a:latin typeface="Cambria Math"/></a:rPr><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"><a:spLocks noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noTextEdit="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="4114800"/><a:ext cx="1625600" cy="369332"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#;

const VIDEO: &str = r#"<p:pic><p:nvPicPr><p:cNvPr id="6" name="Lecture clip"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rId8"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId9"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="5486400" y="4114800"/><a:ext cx="3048000" cy="1714500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#;

const AUDIO: &str = r#"<p:pic><p:nvPicPr><p:cNvPr id="7" name="Narration"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="rId11"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId12"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="9144000" y="4114800"/><a:ext cx="487680" cy="487680"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#;

/// A Fade transition as PowerPoint 2010 and later writes it.
const TRANSITION: &str = r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="slow" p14:dur="1500"><p:fade/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>"#;

/// One click of the main sequence: the title (2) appears.
const CLICK_TITLE: &str = r#"<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"#;

/// The next click: the chart (3) fades in.
const CLICK_CHART: &str = r#"<p:par><p:cTn id="7" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="8" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="9" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="10" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="3"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="11" dur="500"/><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"#;

/// A trigger: clicking the title (2) plays or pauses the narration (7).
const TITLE_TRIGGER: &str = r#"<p:seq concurrent="1" nextAc="seek"><p:cTn id="12" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="13" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="14" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="15" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="16" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>"#;

const VIDEO_NODE: &str = r#"<p:video><p:cMediaNode vol="80000"><p:cTn id="17" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="6"/></p:tgtEl></p:cMediaNode></p:video>"#;

const AUDIO_NODE: &str = r#"<p:audio><p:cMediaNode vol="80000"><p:cTn id="18" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cMediaNode></p:audio>"#;

const BUILD_TITLE: &str = r#"<p:bldP spid="2" grpId="0"/>"#;
const BUILD_CHART: &str = r#"<p:bldGraphic spid="3" grpId="0"><p:bldAsOne/></p:bldGraphic>"#;

/// The slide's timing as PowerPoint writes it: the main sequence's clicks,
/// then trigger sequences, then the media nodes, and the build list.
fn timing(clicks: &str, triggers: &str, media: &str, builds: &str) -> String {
    format!(
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{clicks}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>{triggers}{media}</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>{builds}</p:bldLst></p:timing>"#
    )
}

fn full_timing() -> String {
    timing(
        &[CLICK_TITLE, CLICK_CHART].concat(),
        TITLE_TRIGGER,
        &[VIDEO_NODE, AUDIO_NODE].concat(),
        &[BUILD_TITLE, BUILD_CHART].concat(),
    )
}

fn slide() -> String {
    let timing = full_timing();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{TITLE}{CHART_FRAME}{SMARTART_FRAME}{EQUATION}{VIDEO}{AUDIO}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{TRANSITION}{timing}</p:sld>"#
    )
}

fn relationship(id: &str, kind: &str, target: &str) -> String {
    let kind = if kind.starts_with("http") {
        kind.to_owned()
    } else {
        format!("{R}/{kind}")
    };
    format!(r#"<Relationship Id="{id}" Type="{kind}" Target="{target}"/>"#)
}

fn slide_rels() -> String {
    let media = "http://schemas.microsoft.com/office/2007/relationships/media";
    let drawing = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
    let relationships = [
        relationship("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
        relationship("rId2", "chart", "../charts/chart1.xml"),
        relationship("rId3", "diagramData", "../diagrams/data1.xml"),
        relationship("rId4", "diagramLayout", "../diagrams/layout1.xml"),
        relationship("rId5", "diagramQuickStyle", "../diagrams/quickStyle1.xml"),
        relationship("rId6", "diagramColors", "../diagrams/colors1.xml"),
        relationship("rId13", drawing, "../diagrams/drawing1.xml"),
        relationship("rId7", "image", "../media/image1.png"),
        relationship("rId8", "video", "../media/media1.mp4"),
        relationship("rId9", media, "../media/media1.mp4"),
        relationship("rId10", "image", "../media/image2.png"),
        relationship("rId11", "audio", "../media/media2.m4a"),
        relationship("rId12", media, "../media/media2.m4a"),
    ];
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#,
        relationships.concat()
    )
}

const CHART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cells</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>G1</c:v></c:pt><c:pt idx="1"><c:v>M</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>"#;

const CHART_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>"#;

const DIAGRAM_DATA: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{0}" type="doc"/><dgm:pt modelId="{1}"><dgm:t><a:bodyPr/><a:p><a:r><a:t>Prophase</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{2}" srcId="{0}" destId="{1}"/></dgm:cxnLst><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId13" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>"#;

const DIAGRAM_DRAWING: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/><dsp:sp modelId="{1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1371600" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Prophase</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#;

const MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#;

const MASTER_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>"#;

const THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture"><a:themeElements><a:clrScheme name="Fixture"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>"#;

fn deck() -> Vec<u8> {
    let diagram = |root: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:{root} xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:microsoft.com/office/officeart/2005/8/{root}/default"/>"#
        )
        .into_bytes()
    };
    let bytes = |body: &str| body.as_bytes().to_vec();
    let parts: Vec<(String, Vec<u8>)> = vec![
        ("[Content_Types].xml".into(), bytes(CONTENT_TYPES)),
        ("_rels/.rels".into(), bytes(ROOT_RELS)),
        ("ppt/presentation.xml".into(), bytes(PRESENTATION)),
        (
            "ppt/_rels/presentation.xml.rels".into(),
            bytes(PRESENTATION_RELS),
        ),
        ("ppt/slides/slide1.xml".into(), slide().into_bytes()),
        (
            "ppt/slides/_rels/slide1.xml.rels".into(),
            slide_rels().into_bytes(),
        ),
        ("ppt/charts/chart1.xml".into(), bytes(CHART)),
        ("ppt/charts/_rels/chart1.xml.rels".into(), bytes(CHART_RELS)),
        (
            "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx".into(),
            b"PK\x03\x04 embedded chart workbook".to_vec(),
        ),
        ("ppt/diagrams/data1.xml".into(), bytes(DIAGRAM_DATA)),
        ("ppt/diagrams/layout1.xml".into(), diagram("layoutDef")),
        ("ppt/diagrams/quickStyle1.xml".into(), diagram("styleDef")),
        ("ppt/diagrams/colors1.xml".into(), diagram("colorsDef")),
        ("ppt/diagrams/drawing1.xml".into(), bytes(DIAGRAM_DRAWING)),
        ("ppt/media/image1.png".into(), PNG.to_vec()),
        ("ppt/media/image2.png".into(), [PNG, b"poster"].concat()),
        (
            "ppt/media/media1.mp4".into(),
            b"\0\0\0\x18ftypmp42 video".to_vec(),
        ),
        (
            "ppt/media/media2.m4a".into(),
            b"\0\0\0\x18ftypM4A  audio".to_vec(),
        ),
        ("ppt/slideLayouts/slideLayout1.xml".into(), bytes(LAYOUT)),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels".into(),
            bytes(LAYOUT_RELS),
        ),
        ("ppt/slideMasters/slideMaster1.xml".into(), bytes(MASTER)),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(),
            bytes(MASTER_RELS),
        ),
        ("ppt/theme/theme1.xml".into(), bytes(THEME)),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn text(parts: &BTreeMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(parts[path].clone()).unwrap()
}

/// The XML with every start tag's attributes in name order: an edited slide
/// is written back in the writer's attribute order, not the source's.
fn canonical(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(start) = rest.find('<') {
        out.push_str(&rest[..start]);
        let end = start + rest[start..].find('>').unwrap() + 1;
        let tag = &rest[start..end];
        rest = &rest[end..];
        if tag.starts_with("</") || tag.starts_with("<?") {
            out.push_str(tag);
            continue;
        }
        let close = if tag.ends_with("/>") { "/>" } else { ">" };
        let body = &tag[1..tag.len() - close.len()];
        let (name, attributes) = body.split_once(' ').unwrap_or((body, ""));
        let pieces: Vec<&str> = attributes.split('"').collect();
        let mut pairs: Vec<String> = pieces
            .chunks(2)
            .filter(|pair| pair.len() == 2)
            .map(|pair| format!(r#"{}"{}""#, pair[0].trim(), pair[1]))
            .collect();
        pairs.sort();
        out.push('<');
        out.push_str(name);
        for pair in pairs {
            out.push(' ');
            out.push_str(&pair);
        }
        out.push_str(close);
    }
    out.push_str(rest);
    out
}

/// Types into the title, the one shape on the slide the editor models (and
/// the one the animation reveals), and returns the source and saved packages.
fn edit_the_slide_and_save() -> (BTreeMap<String, Vec<u8>>, BTreeMap<String, Vec<u8>>) {
    let source = deck();
    let session = DeckSession::open(&source, 4101).unwrap();
    let title = |session: &DeckSession| {
        session.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name == "Title 1")
            .unwrap()
            .text_stories[0]
            .clone()
    };
    session
        .insert_text(
            &EditCtx::local("kept"),
            &title(&session).id,
            12,
            " fast",
            &TextStyle::default(),
        )
        .unwrap();
    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 4102).unwrap();
    assert_eq!(title(&reopened).plain_text(), "Cells divide fast");
    (parts(&source), parts(&saved))
}

fn saved_slide(saved: &BTreeMap<String, Vec<u8>>) -> String {
    canonical(&text(saved, "ppt/slides/slide1.xml"))
}

fn assert_parts_kept(
    source: &BTreeMap<String, Vec<u8>>,
    saved: &BTreeMap<String, Vec<u8>>,
    paths: &[&str],
) {
    for path in paths {
        assert_eq!(saved.get(*path), Some(&source[*path]), "{path} changed");
    }
}

fn assert_relationships_kept(saved: &BTreeMap<String, Vec<u8>>, ids: &[&str]) {
    let rels = text(saved, "ppt/slides/_rels/slide1.xml.rels");
    for relationship in slide_rels().split("<Relationship ").skip(1) {
        let relationship = format!(
            "<Relationship {}",
            relationship.trim_end_matches("</Relationships>")
        );
        if ids
            .iter()
            .any(|id| relationship.contains(&format!(r#"Id="{id}""#)))
        {
            assert!(
                rels.contains(&relationship),
                "{relationship} is missing: {rels}"
            );
        }
    }
}

#[test]
fn an_edited_slide_keeps_its_chart_and_the_chart_parts() {
    let (source, saved) = edit_the_slide_and_save();
    assert!(saved_slide(&saved).contains(&canonical(CHART_FRAME)));
    assert_parts_kept(
        &source,
        &saved,
        &[
            "ppt/charts/chart1.xml",
            "ppt/charts/_rels/chart1.xml.rels",
            "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx",
        ],
    );
    assert_relationships_kept(&saved, &["rId2"]);
    assert!(text(&saved, "[Content_Types].xml").contains(r#"PartName="/ppt/charts/chart1.xml""#));
}

#[test]
fn an_edited_slide_keeps_its_smartart_and_the_diagram_parts() {
    let (source, saved) = edit_the_slide_and_save();
    assert!(saved_slide(&saved).contains(&canonical(SMARTART_FRAME)));
    assert_parts_kept(
        &source,
        &saved,
        &[
            "ppt/diagrams/data1.xml",
            "ppt/diagrams/layout1.xml",
            "ppt/diagrams/quickStyle1.xml",
            "ppt/diagrams/colors1.xml",
            "ppt/diagrams/drawing1.xml",
        ],
    );
    assert_relationships_kept(&saved, &["rId3", "rId4", "rId5", "rId6", "rId13"]);
}

#[test]
fn an_edited_slide_keeps_its_equation_and_the_fallback_picture() {
    let (source, saved) = edit_the_slide_and_save();
    assert!(saved_slide(&saved).contains(&canonical(EQUATION)));
    assert_parts_kept(&source, &saved, &["ppt/media/image1.png"]);
    assert_relationships_kept(&saved, &["rId7"]);
}

#[test]
fn an_edited_slide_keeps_its_video_audio_and_their_media_parts() {
    let (source, saved) = edit_the_slide_and_save();
    let slide = saved_slide(&saved);
    assert!(slide.contains(&canonical(VIDEO)));
    assert!(slide.contains(&canonical(AUDIO)));
    assert_parts_kept(
        &source,
        &saved,
        &[
            "ppt/media/media1.mp4",
            "ppt/media/media2.m4a",
            "ppt/media/image2.png",
        ],
    );
    assert_relationships_kept(&saved, &["rId8", "rId9", "rId10", "rId11", "rId12"]);
    let content_types = text(&saved, "[Content_Types].xml");
    assert!(content_types.contains(r#"Extension="mp4""#));
    assert!(content_types.contains(r#"Extension="m4a""#));
}

#[test]
fn an_edited_slide_keeps_its_transition_and_animations() {
    let (_, saved) = edit_the_slide_and_save();
    let tail = format!(
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{TRANSITION}{}</p:sld>",
        full_timing()
    );
    assert!(saved_slide(&saved).ends_with(&canonical(&tail)));
}

/// Deletes the named shapes from the slide and returns the saved slide.
fn delete_and_save(names: &[&str]) -> String {
    let session = DeckSession::open(&deck(), 4103).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    for name in names {
        let shape = slide
            .shapes
            .iter()
            .find(|shape| shape.name == *name)
            .unwrap();
        session
            .remove_shape(&EditCtx::local("kept"), &slide.id, &shape.id)
            .unwrap();
    }
    let saved = parts(&session.save().unwrap());
    canonical(&text(&saved, "ppt/slides/slide1.xml"))
}

/// Every shape id the slide's timing names, and every id on the slide.
fn timing_targets(slide: &str) -> (Vec<String>, Vec<String>) {
    let ids = |marker: &str, within: &str| -> Vec<String> {
        within
            .split(marker)
            .skip(1)
            .map(|rest| rest.split('"').next().unwrap().to_owned())
            .collect()
    };
    let (tree, timing) = slide.split_once("</p:spTree>").unwrap();
    (ids(r#"spid=""#, timing), ids(r#"<p:cNvPr id=""#, tree))
}

#[test]
fn deleting_animated_shapes_and_a_video_prunes_their_timing() {
    let slide = delete_and_save(&["Title 1", "Lecture clip"]);
    // The title's click, the trigger it was and the video's node go, with the
    // title's build entry; the chart's click and the narration's node stay.
    let expected = timing(CLICK_CHART, "", AUDIO_NODE, BUILD_CHART);
    let tail = format!("{TRANSITION}{expected}</p:sld>");
    assert!(slide.ends_with(&canonical(&tail)), "{slide}");
    let (targets, shapes) = timing_targets(&slide);
    assert_eq!(targets, ["3", "3", "7", "3"]);
    assert!(targets.iter().all(|id| shapes.contains(id)), "{shapes:?}");
}

#[test]
fn deleting_every_animated_shape_drops_the_timing() {
    let slide = delete_and_save(&["Title 1", "Chart 2", "Lecture clip", "Narration"]);
    let tail = format!("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{TRANSITION}</p:sld>");
    assert!(slide.ends_with(&canonical(&tail)), "{slide}");
}
