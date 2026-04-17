$ErrorActionPreference = "Stop"

$root = "C:\Users\Jason Yang\Desktop\Chemical Capex"
$outFile = Join-Path $root "china_chemicals_rd_capex_case_study.pptx"
$work = Join-Path $root "_pptx_build"

if (Test-Path $work) {
    Remove-Item -LiteralPath $work -Recurse -Force
}

$null = New-Item -ItemType Directory -Path $work
@("_rels", "docProps", "ppt", "ppt\_rels", "ppt\slides", "ppt\slides\_rels", "ppt\slideLayouts", "ppt\slideLayouts\_rels", "ppt\slideMasters", "ppt\slideMasters\_rels", "ppt\theme") |
    ForEach-Object { $null = New-Item -ItemType Directory -Path (Join-Path $work $_) }

function Write-Utf8File {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}

function Escape-Xml {
    param([string]$Text)
    return $Text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;").Replace("'", "&apos;")
}

function ParagraphXml {
    param([string]$Text, [int]$Size, [string]$Color, [switch]$Bold)
    $boldAttr = if ($Bold) { ' b="1"' } else { "" }
    $safe = Escape-Xml $Text
    return "<a:p><a:pPr algn=""l""/><a:r><a:rPr lang=""en-US"" sz=""$Size""$boldAttr><a:solidFill><a:srgbClr val=""$Color""/></a:solidFill></a:rPr><a:t>$safe</a:t></a:r><a:endParaRPr lang=""en-US"" sz=""$Size""/></a:p>"
}

function TextShapeXml {
    param([int]$Id, [string]$Name, [int]$X, [int]$Y, [int]$Cx, [int]$Cy, [string[]]$Paragraphs)
    $body = $Paragraphs -join ""
    return @"
<p:sp>
  <p:nvSpPr><p:cNvPr id="$Id" name="$Name"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="$X" y="$Y"/><a:ext cx="$Cx" cy="$Cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>$body</p:txBody>
</p:sp>
"@
}

function SlideXml {
    param([string]$Title, [string[]]$Bullets, [string]$Footer)
    $title = TextShapeXml -Id 2 -Name "Title" -X 457200 -Y 228600 -Cx 11277600 -Cy 685800 -Paragraphs @((ParagraphXml -Text $Title -Size 2500 -Color "133B5C" -Bold))
    $body = TextShapeXml -Id 3 -Name "Body" -X 685800 -Y 1219200 -Cx 10744200 -Cy 4420000 -Paragraphs (($Bullets | ForEach-Object { ParagraphXml -Text ("- " + $_) -Size 1650 -Color "1F1F1F" }))
    $footerText = if ($Footer) { @((ParagraphXml -Text $Footer -Size 1000 -Color "666666")) } else { @() }
    $footer = TextShapeXml -Id 4 -Name "Footer" -X 685800 -Y 6248400 -Cx 10744200 -Cy 304800 -Paragraphs $footerText
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F7F4EE"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      $title
      $body
      $footer
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

$slides = @(
    @{ Title = "China Chemicals R&D CAPEX: Executive View"; Bullets = @("Strongest public China benchmarks are BASF, Solvay, Clariant, Arkema, and Nouryon.", "Disclosed investment sizes range from CHF 45 million at Clariant to more than RMB 4 billion cumulative at Solvay.", "Most investments cluster in Shanghai and the Yangtze River Delta and combine application labs with pilot or process-development capability.", "Public disclosures are specific on capability mix but not on blast chambers, explosion cells, or similar high-hazard lab hardware."); Footer = "Sources dated Mar 2021 to Nov 2025; prepared Apr 17, 2026" },
    @{ Title = "China Benchmark Cases"; Bullets = @("BASF: around EUR 280 million in Innovation Campus Shanghai since 2012; two new R&D buildings announced Jun 28, 2023.", "Solvay: more than RMB 4 billion in its China R&I hub since 2005; new Shanghai research building announced Sep 6, 2023.", "Clariant: CHF 45 million for One Clariant Campus; over 13,000 m2 of labs plus dedicated catalyst R&D center announced in Mar 2021.", "Arkema: around EUR 600 million in Changshu platform integrating production and R&D; largest Arkema R&D center in Asia.", "Nouryon: amount undisclosed, but 2025-2026 buildout includes a larger Shanghai innovation center, Tianjin organic peroxides innovation center, and Jiaxing metal-alkyl lab."); Footer = "Core benchmark set for China chemicals R&D CAPEX" },
    @{ Title = "Why These Cases Matter"; Bullets = @("BASF is the best general benchmark for process engineering, solids handling, digital automation, and cleaning/process labs.", "Solvay is the clearest example of a China application-and-pilot hall model with robotics and automation capability.", "Clariant is the strongest public example for a specialized lab footprint with catalyst prep, characterization, and multi-scale performance testing.", "Arkema is especially relevant for organic peroxides because Luperox products are explicitly part of the Changshu R&D story.", "Nouryon is now the most segment-relevant comparator because the company has announced both organic-peroxides and metal-alkyl innovation capability in China."); Footer = "Use BASF, Solvay, and Clariant as scale benchmarks; use Arkema and Nouryon for segment read-across" },
    @{ Title = "Safety Read-Across"; Bullets = @("No reviewed public source explicitly discloses blast chambers, detonation cells, or bunkerized explosive test rooms in China R&D centers.", "Nouryon's 2021 Tianjin organic peroxides site describes latest technology for safety, energy efficiency, and environmental protection.", "Nouryon's 2025 Jiaxing metal-alkyl expansion includes advanced thermal oxidizer systems plus blending and cylinder-transfilling capability.", "United Initiators says its Hefei site uses improved local process and higher HSE operating standard.", "Gulbrandsen's organometallic materials emphasize product stewardship, customer training, and employee safety, but not in China."); Footer = "Best public safety proxies are operating-model descriptions rather than lab-hardware disclosures" },
    @{ Title = "Organic Peroxides Competitor Screen"; Bullets = @("Arkema is the highest-priority external benchmark: strong China footprint, largest Asia R&D center, and organic peroxide relevance.", "United Initiators has China commercial and production footprint, but this review found no public China R&D center disclosure.", "AKPA Kimya shows meaningful global R&D investment in Turkey, but no clear public China R&D footprint.", "Pergan, NOF, Dongsung, Daoming, and Chinasun have weak or unverified public China R&D evidence in the source set.", "Competitor one-pagers should start with Arkema and United Initiators, then only expand if more local evidence emerges."); Footer = "Screening based on accessible public disclosures as of Apr 17, 2026" },
    @{ Title = "Metal Alkyls Competitor Screen"; Bullets = @("LANXESS is the strongest China-adjacent R&D footprint among the named competitors: 10 subsidiaries with eight R&D centers and production sites in China.", "Albemarle has broad China commercial and operating footprint, but the clearly disclosed R&D site in the reviewed materials is Baton Rouge, not China.", "Gulbrandsen is a credible global organometallic benchmark with a 2023 aluminum-alkyl plant investment in India, not China.", "For China-focused metal alkyl intelligence, LANXESS deserves the first one-pager and Albemarle the second."); Footer = "China evidence is materially stronger for LANXESS than for Albemarle or Gulbrandsen" },
    @{ Title = "Recommended Follow-Up"; Bullets = @("Build Tier 1 one-pagers for Arkema, Nouryon benchmark, LANXESS, and United Initiators.", "Use Albemarle, Gulbrandsen, and AKPA Kimya as Tier 2 pages only if the project needs a fuller global competitor set.", "Treat Pergan, NOF, Daoming, Chinasun, and Dongsung as research gaps unless the team can provide exact local legal-entity names or Chinese-language leads.", "If deeper safety detail is required, the next step should be Chinese-language EIA, permit, or tender-document research rather than company press releases."); Footer = "Priority order based on public evidence strength and relevance to Nouryon's segment" }
)

$slideCount = $slides.Count
$slideContentTypes = (1..$slideCount | ForEach-Object { "  <Override PartName=""/ppt/slides/slide$_.xml"" ContentType=""application/vnd.openxmlformats-officedocument.presentationml.slide+xml""/>" }) -join "`r`n"
$slideIds = (0..($slideCount - 1) | ForEach-Object { "    <p:sldId id=""$(256 + $_)"" r:id=""rId$(2 + $_)""/>" }) -join "`r`n"
$slideRels = (1..$slideCount | ForEach-Object { "  <Relationship Id=""rId$($_ + 1)"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"" Target=""slides/slide$_.xml""/>" }) -join "`r`n"

$contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
$slideContentTypes
</Types>
"@

$presentationXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
$slideIds
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>
"@

$presentationRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
$slideRels
  <Relationship Id="rId$($slideCount + 2)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId$($slideCount + 3)" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
</Relationships>
"@

$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$appProps = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <PresentationFormat>Widescreen</PresentationFormat>
  <Slides>$slideCount</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr></vt:vector></TitlesOfParts>
  <Company>OpenAI</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"@

$coreProps = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>China Chemicals R&amp;D CAPEX Dashboard</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-04-17T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-17T00:00:00Z</dcterms:modified>
</cp:coreProperties>
'@

$presProps = '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
$viewProps = '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr/><p:notesTextViewPr/><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>'
$slideMaster = '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Office Theme"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F7F4EE"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l"/></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="0" indent="0"/></p:bodyStyle><p:otherStyle><a:defPPr/></p:otherStyle></p:txStyles></p:sldMaster>'
$slideMasterRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'
$slideLayout = '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
$slideLayoutRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'
$theme = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Custom"><a:dk1><a:srgbClr val="1F1F1F"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="133B5C"/></a:dk2><a:lt2><a:srgbClr val="F7F4EE"/></a:lt2><a:accent1><a:srgbClr val="133B5C"/></a:accent1><a:accent2><a:srgbClr val="2E7D6E"/></a:accent2><a:accent3><a:srgbClr val="D08632"/></a:accent3><a:accent4><a:srgbClr val="A33E3E"/></a:accent4><a:accent5><a:srgbClr val="4F6D7A"/></a:accent5><a:accent6><a:srgbClr val="7A5C99"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Custom"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Simple"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
$singleSlideRel = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'

Write-Utf8File -Path (Join-Path $work "[Content_Types].xml") -Content $contentTypes
Write-Utf8File -Path (Join-Path $work "_rels\.rels") -Content $rootRels
Write-Utf8File -Path (Join-Path $work "docProps\app.xml") -Content $appProps
Write-Utf8File -Path (Join-Path $work "docProps\core.xml") -Content $coreProps
Write-Utf8File -Path (Join-Path $work "ppt\presentation.xml") -Content $presentationXml
Write-Utf8File -Path (Join-Path $work "ppt\_rels\presentation.xml.rels") -Content $presentationRels
Write-Utf8File -Path (Join-Path $work "ppt\presProps.xml") -Content $presProps
Write-Utf8File -Path (Join-Path $work "ppt\viewProps.xml") -Content $viewProps
Write-Utf8File -Path (Join-Path $work "ppt\slideMasters\slideMaster1.xml") -Content $slideMaster
Write-Utf8File -Path (Join-Path $work "ppt\slideMasters\_rels\slideMaster1.xml.rels") -Content $slideMasterRels
Write-Utf8File -Path (Join-Path $work "ppt\slideLayouts\slideLayout1.xml") -Content $slideLayout
Write-Utf8File -Path (Join-Path $work "ppt\slideLayouts\_rels\slideLayout1.xml.rels") -Content $slideLayoutRels
Write-Utf8File -Path (Join-Path $work "ppt\theme\theme1.xml") -Content $theme

for ($i = 0; $i -lt $slideCount; $i++) {
    $name = "slide$($i + 1).xml"
    $slide = $slides[$i]
    Write-Utf8File -Path (Join-Path $work "ppt\slides\$name") -Content (SlideXml -Title $slide.Title -Bullets $slide.Bullets -Footer $slide.Footer)
    Write-Utf8File -Path (Join-Path $work "ppt\slides\_rels\$name.rels") -Content $singleSlideRel
}

if (Test-Path $outFile) {
    Remove-Item -LiteralPath $outFile -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($work, $outFile)
Remove-Item -LiteralPath $work -Recurse -Force

Write-Output "Created $outFile"
