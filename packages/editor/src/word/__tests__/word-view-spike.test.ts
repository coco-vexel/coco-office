import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import { fromZip } from '@coco-vexel/coco-office/pack';
import { buildIndices, W14_NS, W_NS } from '@coco-vexel/coco-office/word-view';
import { buildWordEditorViewFromPackage } from '../extract.js';
import { projectWordViewToProseMirror, wordEditorSchema } from '../prosemirror.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="MIXED001">
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="240" w:after="120" w:line="288" w:lineRule="auto"/>
        <w:shd w:val="clear" w:fill="FCE4D6"/>
        <w:keepNext/>
        <w:keepLines/>
        <w:widowControl w:val="0"/>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr><w:t>A</w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:rPr><w:highlight w:val="yellow"/><w:shd w:val="clear" w:fill="D9EAF7"/><w:spacing w:val="-10"/><w:vertAlign w:val="superscript"/><w:strike/><w:smallCaps/></w:rPr><w:t>B</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:del w:id="20" w:author="bob" w:date="2026-04-22T11:00:00Z">
        <w:r><w:rPr><w:rFonts w:eastAsia="方正仿宋_GB2312"/><w:sz w:val="32"/></w:rPr><w:delText>old</w:delText></w:r>
      </w:del>
      <w:ins w:id="10" w:author="alice" w:date="2026-04-22T10:00:00Z">
        <w:r><w:rPr><w:rFonts w:eastAsia="方正仿宋_GB2312"/><w:sz w:val="32"/></w:rPr><w:t>new</w:t></w:r>
      </w:ins>
      <w:r><w:rPr><w:dstrike/><w:caps/></w:rPr><w:t>C</w:t></w:r>
    </w:p>
    <w:p w14:paraId="NUM00001"><w:pPr><w:ind w:left="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>N1</w:t></w:r></w:p>
    <w:p w14:paraId="NUM00002"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>N2</w:t></w:r></w:p>
    <w:p w14:paraId="BREAK001"><w:r><w:br w:type="page"/></w:r></w:p>
    <w:sectPr>
      <w:footerReference w:type="default" r:id="rIdF"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
      <w:pgNumType w:start="7"/>
      <w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}">
  <w:comment w:id="1" w:author="reviewer" w:date="2026-04-22T09:00:00Z">
    <w:p><w:r><w:t>check this</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W_NS}"><w:sdt><w:sdtContent><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:sdtContent></w:sdt></w:ftr>`;

const FIELD_FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W_NS}">
  <w:p>
    <w:r><w:t>Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>4</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t> of </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES   \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>122</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t> section </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SECTIONPAGES   \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>8</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>`;

const HEADER_SDT_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="BODYSDT1">
      <w:sdt>
        <w:sdtPr><w:showingPlcHdr/></w:sdtPr>
        <w:sdtContent><w:r><w:t>body placeholder remains visible</w:t></w:r></w:sdtContent>
      </w:sdt>
    </w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>
  </w:body>
</w:document>`;

const HEADER_SDT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:sdt>
    <w:sdtPr><w:showingPlcHdr/></w:sdtPr>
    <w:sdtContent><w:p w14:paraId="HDRBLOCK"><w:r><w:t>block placeholder</w:t></w:r></w:p></w:sdtContent>
  </w:sdt>
  <w:sdt>
    <w:sdtPr><w:alias w:val="filled header"/></w:sdtPr>
    <w:sdtContent><w:p w14:paraId="HDRREAL1"><w:r><w:t>filled header</w:t></w:r></w:p></w:sdtContent>
  </w:sdt>
  <w:p w14:paraId="HDRINLINE">
    <w:r><w:t>before</w:t></w:r>
    <w:sdt>
      <w:sdtPr><w:showingPlcHdr/></w:sdtPr>
      <w:sdtContent><w:r><w:t>inline placeholder</w:t></w:r></w:sdtContent>
    </w:sdt>
    <w:r><w:t>after</w:t></w:r>
  </w:p>
  <w:p w14:paraId="HDRFALSE">
    <w:sdt>
      <w:sdtPr><w:showingPlcHdr w:val="0"/></w:sdtPr>
      <w:sdtContent><w:r><w:t>explicitly visible</w:t></w:r></w:sdtContent>
    </w:sdt>
  </w:p>
</w:hdr>`;

const TEMPLATE_FIELD_TEXT_BOX_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p w14:paraId="TEMPLATE">
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">
            <w:drawing>
              <wp:anchor>
                <wp:positionH relativeFrom="column"><wp:posOffset>-1143000</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="paragraph"><wp:posOffset>-914400</wp:posOffset></wp:positionV>
                <wp:extent cx="63500" cy="63500"/>
                <wp:wrapNone/>
                <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wps:wsp>
                    <wps:spPr>
                      <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                      <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                    </wps:spPr>
                    <wps:txbx><w:txbxContent><w:p w14:paraId="TXBOX001">
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText xml:space="preserve"> SET 标题 &quot;标题&quot; </w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>标题</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText xml:space="preserve"> SET 正文 &quot;正文&quot; </w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>正文</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    </w:p></w:txbxContent></wps:txbx>
                  </wps:wsp>
                </a:graphicData></a:graphic>
              </wp:anchor>
            </w:drawing>
          </mc:Choice>
          <mc:Fallback>
            <w:pict><v:shape style="position:absolute;margin-left:-90pt;margin-top:-72pt;width:5pt;height:5pt">
              <v:textbox><w:txbxContent><w:p><w:r><w:t>标题正文</w:t></w:r></w:p></w:txbxContent></v:textbox>
            </v:shape></w:pict>
          </mc:Fallback>
        </mc:AlternateContent>
      </w:r>
    </w:p>
    <w:p w14:paraId="CONTENT1"><w:r><w:t>正文开始</w:t></w:r></w:p>
  </w:body>
</w:document>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman"/><w:sz w:val="21"/></w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="table" w:styleId="GridStyle">
    <w:name w:val="Grid Style"/>
    <w:tblPr>
      <w:tblBorders>
        <w:insideH w:val="single" w:color="DEE0E3" w:sz="4"/>
        <w:insideV w:val="single" w:color="DEE0E3" w:sz="4"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="20" w:type="dxa"/>
        <w:left w:w="40" w:type="dxa"/>
        <w:bottom w:w="20" w:type="dxa"/>
        <w:right w:w="40" w:type="dxa"/>
      </w:tblCellMar>
      <w:tblStyleRowBandSize w:val="2"/>
      <w:tblStyleColBandSize w:val="3"/>
    </w:tblPr>
    <w:tcPr>
      <w:vAlign w:val="center"/>
    </w:tcPr>
    <w:tblStylePr w:type="firstRow">
      <w:tcPr><w:shd w:fill="B4C6E7"/></w:tcPr>
      <w:rPr><w:b/><w:color w:val="1F4E79"/></w:rPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastRow">
      <w:tcPr><w:shd w:fill="F4B183"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Horz">
      <w:tcPr><w:shd w:fill="E2F0D9"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band2Horz">
      <w:tcPr><w:shd w:fill="FFF2CC"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Vert">
      <w:tcPr><w:shd w:fill="DDEBF7"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band2Vert">
      <w:tcPr><w:shd w:fill="EADCF8"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="lastCol">
      <w:tcPr><w:shd w:fill="C9DAF8"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="nwCell">
      <w:tcPr><w:shd w:fill="C6E0B4"/></w:tcPr>
      <w:rPr><w:color w:val="C00000"/></w:rPr>
    </w:tblStylePr>
  </w:style>
  <w:style w:type="table" w:styleId="ChildGridStyle">
    <w:name w:val="Child Grid Style"/>
    <w:basedOn w:val="GridStyle"/>
    <w:tblPr>
      <w:tblLook w:firstRow="0" w:lastRow="1" w:noVBand="1"/>
    </w:tblPr>
  </w:style>
</w:styles>`;

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="◆"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;

const TABLE_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="GridStyle"/>
        <w:tblW w:w="7200" w:type="dxa"/>
        <w:tblInd w:w="240" w:type="dxa"/>
        <w:tblCellSpacing w:w="120" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="1800"/>
        <w:gridCol w:w="2200"/>
        <w:gridCol w:w="2200"/>
        <w:gridCol w:w="2200"/>
      </w:tblGrid>
      <w:tr>
        <w:trPr><w:tblHeader/><w:cantSplit/><w:trHeight w:val="720" w:hRule="exact"/></w:trPr>
        <w:tc><w:tcPr><w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar><w:tcBorders><w:left w:val="nil"/><w:bottom w:val="single" w:color="FF0000" w:sz="16"/></w:tcBorders><w:shd w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:t>阶段</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p><w:r><w:t>指标汇总</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge w:val="restart"/><w:noWrap/></w:tcPr><w:p><w:r><w:t>第一季度</w:t></w:r></w:p></w:tc>
        <w:tc>
          <w:p><w:r><w:t>收入</w:t></w:r></w:p>
          <w:tbl>
            <w:tblGrid><w:gridCol w:w="1200"/></w:tblGrid>
            <w:tr><w:tc><w:p><w:r><w:t>嵌套值</w:t></w:r></w:p></w:tc></w:tr>
          </w:tbl>
          <w:p><w:r><w:t>后续段落</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>成本</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>利润</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>40</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>60</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:p><w:r><w:t>110</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>44</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>66</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

const GRID_AWARE_TABLE_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="GridStyle"/>
        <w:tblStyleColBandSize w:val="2"/>
        <w:tblLook w:firstRow="0" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="0"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="1200"/>
        <w:gridCol w:w="1200"/>
        <w:gridCol w:w="1200"/>
        <w:gridCol w:w="1200"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>grid-col-2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>last</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>d</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

const DEFAULT_PARAGRAPH_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="DEFAULT1"><w:r><w:t>default paragraph</w:t></w:r></w:p>
    <w:p w14:paraId="DIRECT1">
      <w:pPr><w:spacing w:after="0"/></w:pPr>
      <w:r><w:t>direct paragraph</w:t></w:r>
    </w:p>
    <w:p w14:paraId="TABSTOP1">
      <w:pPr><w:tabs><w:tab w:val="right" w:pos="8630" w:leader="dot"/></w:tabs></w:pPr>
      <w:r><w:t>left</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>right</w:t></w:r>
    </w:p>
    <w:p w14:paraId="PTAB1">
      <w:pPr><w:tabs><w:tab w:val="center" w:pos="4315"/><w:tab w:val="right" w:pos="8630"/></w:tabs></w:pPr>
      <w:r><w:t>left</w:t></w:r><w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r><w:r><w:t>mid</w:t></w:r><w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r><w:r><w:t>right</w:t></w:r>
    </w:p>
    <w:p w14:paraId="BORDER1">
      <w:pPr><w:pBdr><w:top w:val="double" w:color="00AA00" w:sz="8" w:space="2"/><w:right w:val="single" w:color="0000FF" w:sz="4" w:space="1"/><w:bottom w:val="single" w:color="FF0000" w:sz="12" w:space="0"/><w:left w:val="single" w:color="999999" w:sz="6" w:space="3"/></w:pBdr></w:pPr>
      <w:r><w:t>border paragraph</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

const DEFAULT_PARAGRAPH_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="0"/></w:pPr>
    <w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>`;

const THEME_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="THEME001">
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsiaTheme="majorEastAsia"/>
          <w:color w:themeColor="accent2" w:themeTint="80"/>
        </w:rPr>
        <w:t>theme run</w:t>
      </w:r>
    </w:p>
    <w:p w14:paraId="THEME002">
      <w:r>
        <w:rPr>
          <w:rFonts w:eastAsiaTheme="+mj-ea"/>
        </w:rPr>
        <w:t>theme token east asia</w:t>
      </w:r>
    </w:p>
    <w:p w14:paraId="THEME003">
      <w:r>
        <w:rPr>
          <w:rFonts w:hAnsiTheme="+mn-lt"/>
        </w:rPr>
        <w:t>theme token latin</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr><w:shd w:val="clear" w:themeFill="accent1" w:themeFillShade="80"/></w:tcPr>
          <w:p><w:r><w:t>theme cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Coco Test">
  <a:themeElements>
    <a:clrScheme name="Coco Colors">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2>
      <a:lt2><a:srgbClr val="F3F4F6"/></a:lt2>
      <a:accent1><a:srgbClr val="336699"/></a:accent1>
      <a:accent2><a:srgbClr val="CC3300"/></a:accent2>
      <a:accent3><a:srgbClr val="00AA55"/></a:accent3>
      <a:accent4><a:srgbClr val="8844CC"/></a:accent4>
      <a:accent5><a:srgbClr val="00AACC"/></a:accent5>
      <a:accent6><a:srgbClr val="CCAA00"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Coco Fonts">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="黑体"/><a:cs typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="宋体"/><a:cs typeface="Arial"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  zip.file('word/footer1.xml', FOOTER_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeFieldFooterFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  zip.file('word/footer1.xml', FIELD_FOOTER_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeHeaderSdtFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>`,
  );
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
  );
  zip.file('word/document.xml', HEADER_SDT_DOCUMENT_XML);
  zip.file('word/header1.xml', HEADER_SDT_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeTemplateFieldTextBoxFixtureZip(
  documentXml = TEMPLATE_FIELD_TEXT_BOX_DOCUMENT_XML,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeTableFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', TABLE_DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeTableDefaultLookFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', TABLE_DOC_XML.replace(/\s*<w:tblLook[^>]+\/>\r?\n/, '\n'));
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeTableBasedOnLookFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file(
    'word/document.xml',
    TABLE_DOC_XML.replace('<w:tblStyle w:val="GridStyle"/>', '<w:tblStyle w:val="ChildGridStyle"/>').replace(
      /\s*<w:tblLook[^>]+\/>\r?\n/,
      '\n',
    ),
  );
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeGridAwareTableFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', GRID_AWARE_TABLE_DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeNoParaIdSectionFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        <w:p><w:r><w:t>no para id</w:t></w:r></w:p>
        <w:sectPr>
          <w:footerReference w:type="default" r:id="rIdF"/>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:pgMar w:top="1200" w:right="900" w:bottom="1000" w:left="900"/>
        </w:sectPr>
      </w:body>
    </w:document>`,
  );
  zip.file('word/footer1.xml', FOOTER_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeDefaultParagraphFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', DEFAULT_PARAGRAPH_DOC_XML);
  zip.file('word/styles.xml', DEFAULT_PARAGRAPH_STYLES_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeThemeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', THEME_DOC_XML);
  zip.file('word/theme/theme1.xml', THEME_XML);
  return zip.generateAsync({ type: 'uint8array' });
}

async function makeHyperlinkFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://coco.example/docs" TargetMode="External"/>
    </Relationships>`,
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        <w:p w14:paraId="LINK0001">
          <w:r><w:t>See </w:t></w:r>
          <w:hyperlink r:id="rIdLink" w:tooltip="Coco docs">
            <w:r><w:rPr><w:b/></w:rPr><w:t>Coco</w:t></w:r>
            <w:r><w:t> Docs</w:t></w:r>
          </w:hyperlink>
          <w:r><w:t> and </w:t></w:r>
          <w:hyperlink w:anchor="Section_1"><w:r><w:t>section one</w:t></w:r></w:hyperlink>
        </w:p>
      </w:body>
    </w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

describe('@coco-vexel/coco-office-editor word spike', () => {
  test('preserves external hyperlink relationships, document anchors, and linked run formatting', async () => {
    const pkg = await fromZip(await makeHyperlinkFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const paragraph = view.paragraphs[0]!;

    expect(paragraph.text).toBe('See Coco Docs and section one');
    expect(paragraph.segments).toMatchObject([
      { kind: 'text', text: 'See ' },
      {
        kind: 'text',
        text: 'Coco',
        format: { bold: true },
        hyperlink: { target: 'https://coco.example/docs', tooltip: 'Coco docs' },
      },
      {
        kind: 'text',
        text: ' Docs',
        hyperlink: { target: 'https://coco.example/docs', tooltip: 'Coco docs' },
      },
      { kind: 'text', text: ' and ' },
      { kind: 'text', text: 'section one', hyperlink: { anchor: 'Section_1' } },
    ]);
  });

  test('extracts mixed normal text, comment anchors and revisions in document order', async () => {
    const pkg = await fromZip(await makeFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const paragraph = view.paragraphs[0]!;

    expect(paragraph.segments.filter((segment) => segment.kind === 'revision')).toMatchObject([
      { kind: 'revision', text: 'old', format: { fontFamily: '方正仿宋_GB2312', fontSize: 16 } },
      { kind: 'revision', text: 'new', format: { fontFamily: '方正仿宋_GB2312', fontSize: 16 } },
    ]);

    expect(view.pageLayout).toEqual({
      width: 12240,
      height: 15840,
      margin: {
        top: 1440,
        right: 1800,
        bottom: 1440,
        left: 1800,
      },
      docGrid: {
        type: 'lines',
        linePitch: 312,
        charSpace: 0,
      },
    });
    expect(view.defaults?.run).toMatchObject({ fontFamily: '宋体', fontSize: 10.5 });
    expect(view.sections).toMatchObject([
      {
        sectionIndex: 0,
        footerRefs: [{ type: 'default', partName: 'word/footer1.xml' }],
        headerRefs: [],
        pageLayout: {
          width: 12240,
          height: 15840,
          margin: {
            top: 1440,
            right: 1800,
            bottom: 1440,
            left: 1800,
          },
          docGrid: {
            type: 'lines',
            linePitch: 312,
            charSpace: 0,
          },
        },
        pageNumberStart: 7,
      },
    ]);
    expect(view.footers).toMatchObject([
      {
        kind: 'footer',
        partName: 'word/footer1.xml',
        paragraphs: [{ text: 'footer' }],
      },
    ]);
    expect(paragraph.format?.shading).toEqual({ fill: 'FCE4D6', val: 'clear' });
    expect(paragraph.format).toMatchObject({
      keepNext: true,
      keepLines: true,
      widowControl: false,
    });
    expect(paragraph.text).toBe('ABoldnewC');
    expect(paragraph.segments).toMatchObject([
      { kind: 'text', from: 0, to: 1, text: 'A', format: { fontFamily: '黑体' } },
      { kind: 'text', from: 1, to: 2, text: 'B', format: { characterSpacing: -10, verticalAlign: 'superscript', strike: true, smallCaps: true, highlight: 'yellow', shading: { fill: 'D9EAF7', val: 'clear' } } },
      { kind: 'revision', from: 2, to: 5, text: 'old', revisionId: '20' },
      { kind: 'revision', from: 5, to: 8, text: 'new', revisionId: '10' },
      { kind: 'text', from: 8, to: 9, text: 'C', format: { doubleStrike: true, caps: true } },
    ]);
    expect(view.paragraphs[1]!.format?.numbering).toMatchObject({
      numId: '1',
      ilvl: 0,
      numFmt: 'decimal',
      lvlText: '%1)',
      label: '1)',
    });
    expect(view.paragraphs[1]!.format?.indent).toEqual({ left: 720, firstLine: -360 });
    expect(view.paragraphs[2]!.format?.numbering).toMatchObject({
      numId: '2',
      numFmt: 'bullet',
      label: '◆',
    });
    expect(view.paragraphs[3]!.segments).toMatchObject([
      { kind: 'placeholder', placeholderKind: 'pageBreak', sourceTag: 'br' },
    ]);
    expect(paragraph.spans).toMatchObject([
      { kind: 'comment', id: 'comment:1', from: 1, to: 2, text: 'B' },
      { kind: 'revision', id: 'revision:20', from: 2, to: 5, text: 'old' },
      { kind: 'revision', id: 'revision:10', from: 5, to: 8, text: 'new' },
    ]);

  });

  test('extracts page count fields as field segments instead of cached text', async () => {
    const pkg = await fromZip(await makeFieldFooterFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const footer = view.footers[0]!.paragraphs[0]!;

    expect(footer.text).toBe('Page \uFFFC of \uFFFC section \uFFFC');
    expect(footer.segments).toMatchObject([
      { kind: 'text', text: 'Page ' },
      { kind: 'field', fieldKind: 'pageNumber', cachedText: '4' },
      { kind: 'text', text: ' of ' },
      { kind: 'field', fieldKind: 'totalPages', cachedText: '122' },
      { kind: 'text', text: ' section ' },
      { kind: 'field', fieldKind: 'sectionPages', cachedText: '8' },
    ]);
  });

  test('suppresses placeholder SDT content in headers without hiding filled controls or body content', async () => {
    const pkg = await fromZip(await makeHeaderSdtFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));

    expect(view.paragraphs[0]!.text).toBe('body placeholder remains visible');
    expect(view.headers).toHaveLength(1);
    expect(view.headers[0]!.paragraphs.map((paragraph) => paragraph.text)).toEqual([
      'filled header',
      'beforeafter',
      'explicitly visible',
    ]);
    expect(view.headers[0]!.paragraphs.map((paragraph) => paragraph.paraId)).not.toContain('HDRBLOCK');
  });

  test('suppresses tiny off-page SET-field text boxes without leaking cached template text', async () => {
    const pkg = await fromZip(await makeTemplateFieldTextBoxFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));

    expect(view.paragraphs[0]).toMatchObject({ paraId: 'TEMPLATE', text: '', segments: [] });
    expect(view.paragraphs[1]).toMatchObject({ paraId: 'CONTENT1', text: '正文开始' });
    expect(view.paragraphs.flatMap((paragraph) => paragraph.segments)).not.toContainEqual(
      expect.objectContaining({ kind: 'drawing', drawingKind: 'textBox' }),
    );

    const nonSetPkg = await fromZip(
      await makeTemplateFieldTextBoxFixtureZip(
        TEMPLATE_FIELD_TEXT_BOX_DOCUMENT_XML.replace(' SET 标题 ', ' REF 标题 '),
      ),
    );
    const nonSetView = buildWordEditorViewFromPackage(nonSetPkg, buildIndices(nonSetPkg));
    expect(nonSetView.paragraphs[0]!.segments).toContainEqual(
      expect.objectContaining({
        kind: 'drawing',
        drawingKind: 'textBox',
        text: '标题正文',
        hasBorder: true,
        hasFill: true,
      }),
    );
  });

  test('applies docDefaults, default paragraph style and direct paragraph formatting in order', async () => {
    const pkg = await fromZip(await makeDefaultParagraphFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));

    expect(view.defaults?.paragraph?.spacing).toEqual({
      before: 120,
      after: 240,
      line: 360,
      lineRule: 'auto',
    });
    expect(view.paragraphs[0]!.format?.spacing).toEqual({
      before: 120,
      after: 0,
      line: 360,
      lineRule: 'auto',
    });
    expect(view.paragraphs[0]!.format).toMatchObject({
      fontFamily: '宋体',
      fontSize: 12,
    });
    expect(view.paragraphs[1]!.format?.spacing).toEqual({
      before: 120,
      after: 0,
      line: 360,
      lineRule: 'auto',
    });
    expect(view.paragraphs[2]!.text).toBe('left\tright');
    expect(view.paragraphs[2]!.format?.tabStops).toEqual([
      { alignment: 'right', position: 8630, leader: 'dot' },
    ]);
    expect(view.paragraphs[3]!.text).toBe('left\tmid\tright');
    expect(view.paragraphs[3]!.format?.tabStops).toEqual([
      { alignment: 'center', position: 4315 },
      { alignment: 'right', position: 8630 },
    ]);
    expect(view.paragraphs[4]!.format?.borderBottom).toEqual({
      val: 'single',
      color: 'FF0000',
      size: 12,
      space: 0,
    });
    expect(view.paragraphs[4]!.format?.borders).toMatchObject({
      top: { val: 'double', color: '00AA00', size: 8, space: 2 },
      right: { val: 'single', color: '0000FF', size: 4, space: 1 },
      bottom: { val: 'single', color: 'FF0000', size: 12, space: 0 },
      left: { val: 'single', color: '999999', size: 6, space: 3 },
    });
  });

  test('resolves Word theme fonts and colors while extracting editor view', async () => {
    const pkg = await fromZip(await makeThemeFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));

    expect(view.paragraphs[0]!.segments[0]).toMatchObject({
      kind: 'text',
      text: 'theme run',
      format: {
        fontFamily: '黑体',
        color: 'E69980',
      },
    });
    expect(view.paragraphs[1]!.segments[0]).toMatchObject({
      kind: 'text',
      text: 'theme token east asia',
      format: {
        fontFamily: '黑体',
      },
    });
    expect(view.paragraphs[2]!.segments[0]).toMatchObject({
      kind: 'text',
      text: 'theme token latin',
      format: {
        fontFamily: 'Aptos',
      },
    });
    expect(view.tables[0]!.rows[0]!.cells[0]!.shading).toMatchObject({
      fill: '1A334D',
      val: 'clear',
    });
  });

  test('falls back to stable body paragraph ids when extracting sections without w14:paraId', async () => {
    const pkg = await fromZip(await makeNoParaIdSectionFixtureZip());
    const view = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));

    expect(view.paragraphs[0]?.paraId).toBe('body:p:0');
    expect(view.sections).toMatchObject([
      {
        sectionIndex: 0,
        firstParaId: 'body:p:0',
        lastParaId: 'body:p:0',
        footerRefs: [{ type: 'default', partName: 'word/footer1.xml' }],
        pageLayout: {
          width: 11906,
          height: 16838,
          margin: {
            top: 1200,
            right: 900,
            bottom: 1000,
            left: 900,
          },
        },
      },
    ]);
  });

  test('projects Word view into ProseMirror without making ProseMirror the document source of truth', async () => {
    const pkg = await fromZip(await makeFixtureZip());
    const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const projection = projectWordViewToProseMirror(wordView);
    const state = EditorState.create({ schema: wordEditorSchema, doc: projection.doc });
    const base = projection.paragraphContentPos.get('MIXED001')!;
    const next = state.apply(state.tr.addMark(base, base + 1, wordEditorSchema.marks.strong.create()));

    expect(next.doc.child(0).attrs).toMatchObject({
      paraId: 'MIXED001',
      story: 'body',
      styleId: null,
      format: {
        alignment: 'center',
        spacing: {
          before: 240,
          after: 120,
          line: 288,
          lineRule: 'auto',
        },
      },
    });
    expect(next.doc.child(0).textContent).toBe('ABoldnewC');
    expect(next.doc.child(0).child(0).marks.map((mark) => mark.type.name)).toEqual(['strong']);

  });

  test('preserves cell text without paraIds and projects vertical and horizontal merges', async () => {
    const pkg = await fromZip(await makeTableFixtureZip());
    const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const table = wordView.tables[0]!;

    expect(table.tableStyleId).toBe('GridStyle');
    expect(table.layout).toEqual({
      width: 7200,
      widthType: 'dxa',
      indent: 240,
      cellSpacing: 120,
      cellSpacingType: 'dxa',
      alignment: 'center',
    });
    expect(table.tableLook).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHorizontalBand: false,
      noVerticalBand: true,
      rowBandSize: 2,
      columnBandSize: 3,
    });
    expect(table.rows[0]!.height).toEqual({ val: 720, hRule: 'exact' });
    expect(table.rows[0]!.isHeader).toBe(true);
    expect(table.rows[0]!.cantSplit).toBe(true);
    expect(table.borders).toMatchObject({
      insideH: { val: 'single', color: 'DEE0E3', size: 4 },
      insideV: { val: 'single', color: 'DEE0E3', size: 4 },
    });
    expect(table.rows[0]!.cells[0]!).toMatchObject({
      text: '阶段',
      margin: { top: 60, right: 120, bottom: 30, left: 120 },
      borders: {
        left: { val: 'nil' },
        bottom: { val: 'single', color: 'FF0000', size: 16 },
      },
      shading: { fill: 'D9EAF7' },
    });
    expect(table.rows[0]!.cells[0]!.paragraphs[0]!.format).toMatchObject({
      bold: true,
      color: 'C00000',
    });
    expect(table.rows[0]!.cells[1]!).toMatchObject({
      text: '指标汇总',
      gridSpan: 3,
      margin: { top: 20, right: 40, bottom: 20, left: 40 },
      shading: { fill: 'B4C6E7' },
      verticalAlignment: 'center',
    });
    expect(table.rows[0]!.cells[1]!.paragraphs[0]!.format).toMatchObject({
      bold: true,
      color: '1F4E79',
    });
    expect(table.rows[1]!.cells[0]!).toMatchObject({
      text: '第一季度',
      rowSpan: 3,
      vMerge: 'restart',
      noWrap: true,
      shading: { fill: 'E2F0D9' },
    });
    expect(table.rows[1]!.cells[0]!.paragraphs[0]!.text).toBe('第一季度');
    expect(table.rows[2]!.cells[1]!).toMatchObject({
      text: '100',
      shading: { fill: 'FFF2CC' },
    });
    expect(table.rows[1]!.cells[1]!.blocks.map((block) => block.type)).toEqual([
      'paragraph',
      'table',
      'paragraph',
    ]);
    const nestedTable = table.rows[1]!.cells[1]!.blocks[1];
    expect(nestedTable?.type).toBe('table');
    if (nestedTable?.type !== 'table') throw new Error('nested table fixture missing');
    expect(nestedTable.rows[0]!.cells[0]!.paragraphs[0]!.text).toBe('嵌套值');
    expect(table.rows[1]!.cells[1]!.paragraphs.map((paragraph) => paragraph.text)).toEqual([
      '收入',
      '嵌套值',
      '后续段落',
    ]);
    const projection = projectWordViewToProseMirror(wordView);
    const projectedTable = projection.doc.child(0);
    expect(projectedTable.child(0).child(1).attrs.colspan).toBe(3);
    expect(projectedTable.child(1).child(0).attrs.rowspan).toBe(3);
    expect(projectedTable.textContent).toContain('指标汇总');
    expect(projectedTable.textContent).toContain('第一季度');
  });

  test('uses Word table look defaults when tblLook is omitted', async () => {
    const pkg = await fromZip(await makeTableDefaultLookFixtureZip());
    const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const table = wordView.tables[0]!;

    expect(table.tableLook).toEqual({
      rowBandSize: 2,
      columnBandSize: 3,
    });
    expect(table.rows[0]!.cells[1]!).toMatchObject({
      shading: { fill: 'B4C6E7' },
    });
    expect(table.rows[1]!.cells[1]!).toMatchObject({
      shading: { fill: 'E2F0D9' },
    });
    expect(table.rows[3]!.cells[1]!).toMatchObject({
      shading: { fill: 'FFF2CC' },
    });
  });

  test('preserves parent table look fields when a child table style overrides only some fields', async () => {
    const pkg = await fromZip(await makeTableBasedOnLookFixtureZip());
    const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const table = wordView.tables[0]!;

    expect(table.tableStyleId).toBe('ChildGridStyle');
    expect(table.tableLook).toEqual({
      firstRow: false,
      lastRow: true,
      noVerticalBand: true,
      rowBandSize: 2,
      columnBandSize: 3,
    });
    expect(table.rows[0]!.cells[1]!).toMatchObject({
      shading: { fill: 'E2F0D9' },
    });
    expect(table.rows[1]!.cells[1]!).toMatchObject({
      shading: { fill: 'E2F0D9' },
    });
    expect(table.rows[3]!.cells[1]!).toMatchObject({
      shading: { fill: 'F4B183' },
    });
  });

  test('uses grid columns instead of physical cell indexes for conditional table styles', async () => {
    const pkg = await fromZip(await makeGridAwareTableFixtureZip());
    const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
    const table = wordView.tables[0]!;

    expect(table.rows[0]!.cells[0]!).toMatchObject({
      text: 'wide',
      gridSpan: 2,
      shading: { fill: 'DDEBF7' },
    });
    expect(table.rows[0]!.cells[1]!).toMatchObject({
      text: 'grid-col-2',
      shading: { fill: 'EADCF8' },
    });
    expect(table.rows[0]!.cells[2]!).toMatchObject({
      text: 'last',
      shading: { fill: 'C9DAF8' },
    });
  });
});
