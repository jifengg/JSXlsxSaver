(() => {
    if (typeof global == 'undefined') {
        window.require = function () { };
    }
    let JSZip = require('jszip');
    let XlsxCore = require('js-xlsx-core');
    if (typeof global == 'undefined') {
        JSZip = window.JSZip;
        XlsxCore = window.XlsxCore;
    }
    const {
        Book,
        Sheet,
        Cell,
        ShareString,
        CellStyle,
        CellFont,
        CellFill,
        CellAlignment,
        NumberFormat,
        Image,
        ImageOption,
        HorizontalAlignment,
        VerticalAlignment
    } = XlsxCore;

    const _firstTime = new Date(1902, 0, 1).getTime() + (365 * 2 + 2) * 86400000;

    Book.prototype.MakeXlsx = MakeXlsx;

    var ContentTypeMap = {
        'jpg': 'image/jpeg',
        'png': 'image/png',
    }

    var HyperlinkDefaultFont = {
        Underline: true,
        Color: 10,
        __id: -1,
    }
    var HyperlinkDefaultFontNoID = {
        Underline: true,
        Color: 10
    }

    /**
     * 生成zip文件
     * @param {Book} book 
     * @returns {JSZip}
     */
    function MakeXlsx(book) {
        var cellStyleMap = {};
        var numberFormatMap = {};
        var shareStringMap = {};
        var imageFormatMap = {};
        var imageFileMap = {};
        var cellFontMap = {};
        var cellFillMap = {};

        var zip = new JSZip();
        zip.file('_rels/.rels',
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
</Relationships>`);
        zip.file('docProps/app.xml',
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
 <Application>Microsoft Excel</Application>
 <DocSecurity>0</DocSecurity>
 <ScaleCrop>false</ScaleCrop>
 <HeadingPairs>
  <vt:vector size="2" baseType="variant">
   <vt:variant>
    <vt:lpstr>工作表</vt:lpstr>
   </vt:variant>
   <vt:variant>
    <vt:i4>1</vt:i4>
   </vt:variant>
  </vt:vector>
 </HeadingPairs>
 <TitlesOfParts>
  <vt:vector size="1" baseType="lpstr">
   <vt:lpstr>Sheet_name</vt:lpstr>
  </vt:vector>
 </TitlesOfParts>
 <Company></Company>
 <LinksUpToDate>false</LinksUpToDate>
 <SharedDoc>false</SharedDoc>
 <HyperlinksChanged>false</HyperlinksChanged>
 <AppVersion>15.0300</AppVersion>
</Properties>
`);
        zip.file('docProps/core.xml',
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <dc:creator></dc:creator>
 <cp:lastModifiedBy></cp:lastModifiedBy>
 <dcterms:created xsi:type="dcterms:W3CDTF">2006-09-16T00:00:00Z</dcterms:created>
 <dcterms:modified xsi:type="dcterms:W3CDTF">2019-03-26T07:30:17Z</dcterms:modified>
</cp:coreProperties>
`);

        var shareStringXml = ``;
        var shareStringIndex = 0;
        var workbookXml =
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
 <fileVersion appName="xl" lastEdited="6" lowestEdited="4" rupBuild="14420" />
 <workbookPr filterPrivacy="1" defaultThemeVersion="124226" />
 <bookViews>
  <workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" />
 </bookViews>
 <sheets>`;
        var workbookXmlRels =
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
 <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />  
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
`;
        var ContentTypesXml =
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
 <Default Extension="xml" ContentType="application/xml" />
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
 <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
 <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
 <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
 <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" /> 
`;
        var drwaingContentTypeXml = ``;
        var stylesXml = ``;
        var styleFontsXml = makeFontXml(book.DefaultCellStyle.Font);// `<font><sz val="11" /><color theme="1" /><name val="宋体" /><family val="2" /><scheme val="minor" /></font>\n`;
        var styleFontsCount = 1;
        var cellStyleXfs = `<xf numFmtId="0" fontId="${(styleFontsCount - 1)}" fillId="0" borderId="0" />\n`;
        var cellStylexfsCount = 1;
        var cellXfs = `<xf numFmtId="0" fontId="${(styleFontsCount - 1)}" fillId="0" borderId="0" xfId="0" />\n`;
        var cellXfsCount = 1;
        var cellStyles = `<cellStyle name="常规" xfId="0" builtinId="0" />\n`;
        var cellStylesCount = 1;
        var numberFormatXml = ``;
        var numberFormatCount = 0;
        var fillXml = `<fill><patternFill patternType="none" /></fill>
<fill><patternFill patternType="gray125" /></fill>`;
        var fillCount = 2;

        var fidOffset = 1e6;
        var sheetCount = 0;
        var drawingCount = 0;
        var sheetRelsId = 0;
        for (var sheet of book.Sheets) {
            sheetCount++;
            var sheetName = HtmlEncodeByRegExp(sheet.Name);
            var sheetXml =
                `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
 <dimension ref="A1:A1" />
 <sheetViews>
  <sheetView workbookViewId="0">
   <selection activeCell="A1" sqref="A1" />
  </sheetView>
 </sheetViews>
 <sheetFormatPr defaultRowHeight="${sheet.DefaultHeight}" customHeight="1" defaultColWidth="${sheet.DefaultWidth}" customWidth="1" x14ac:dyDescent="0.15" />`;
            sheetXmlRels = ``;
            var sheetDataXml = ``;
            workbookXml += `<sheet name="${sheetName}" sheetId="${sheetCount}" r:id="${sheet.id}" />\n`;
            workbookXmlRels += `<Relationship Id="${sheet.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetCount}.xml" />\n`
            ContentTypesXml += `<Override PartName="/xl/worksheets/sheet${sheetCount}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />\n`;
            hyperlinksXml = ``;
            var drawingXmlRels = `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
            for (var row in sheet.Datas) {
                var rowData = sheet.Datas[row];
                sheetDataXml += `<row r="${(+row + 1)}"${(sheet.rowHeight[row] ? ` ht="${sheet.rowHeight[row]}"` : ``)} customHeight="1">`;
                for (var col in rowData) {
                    var cell = rowData[col];
                    if (cell.Text != null) {
                        var s = null;
                        var link = cell.Hyperlink;
                        var style = cell.Style;
                        if (link != null) {
                            if (style == null) {
                                style = {};
                            }
                            if (style.Font == null) {
                                style.Font = HyperlinkDefaultFont;
                            } else {
                                style.Font = Object.assign(HyperlinkDefaultFontNoID, style.Font);
                            }
                        }
                        if (style != null) {
                            var styleID = style.__id;
                            //如果样式是公用的且已经添加了
                            if (styleID != null && cellStyleMap[styleID] != null) {
                                s = cellStyleMap[styleID];
                            } else {
                                //数字格式化
                                var numFmtId = 0;
                                var format = style.Format;
                                if (format != null) {
                                    var formatID = format.__id;
                                    if (formatID != null && numberFormatMap[formatID] != null) {
                                        numFmtId = numberFormatMap[formatID];
                                    } else {
                                        numberFormatCount++;
                                        numFmtId = fidOffset + numberFormatCount;
                                        numberFormatXml += `<numFmt numFmtId="${numFmtId}" formatCode="${format.Code}" />\n`;
                                        if (formatID != null) {
                                            numberFormatMap[formatID] = numFmtId;
                                        }
                                    }
                                }
                                //字体
                                var fontid = 0;
                                var font = style.Font;
                                if (font != null) {
                                    var tempFontID = font.__id;
                                    if (tempFontID != null && cellFontMap[tempFontID] != null) {
                                        fontid = cellFontMap[tempFontID];
                                    } else {
                                        var oneFont = makeFontXml(font);
                                        if (oneFont != '<font></font>\n') {
                                            styleFontsXml += oneFont;
                                            styleFontsCount++;
                                            fontid = styleFontsCount - 1;
                                        }
                                        if (tempFontID != null) {
                                            cellFontMap[tempFontID] = fontid;
                                        }
                                    }
                                }
                                //填充
                                var fillid = 0;
                                var fill = style.Fill;
                                if (fill != null) {
                                    var tempFillID = fill.__id;
                                    if (tempFillID != null && cellFillMap[tempFillID] != null) {
                                        fillid = cellFillMap[tempFillID];
                                    } else {
                                        fillXml += ` <fill><patternFill patternType="solid"><fgColor rgb="${fill.BGColor}" /></patternFill></fill>`;
                                        fillCount++;
                                        fillid = fillCount - 1;
                                        if (tempFillID != null) {
                                            cellFillMap[tempFillID] = fillid;
                                        }
                                    }
                                }

                                //对齐
                                var alignment = null;
                                if (style.Alignment != null) {
                                    var align = style.Alignment;
                                    alignment = `<alignment${(align.WrapText ? ' wrapText="1"' : '')}${(align.Horizontal ? ` horizontal="${align.Horizontal}"` : '')}${(align.Vertical ? ` vertical="${align.Vertical}"` : '')} />`;
                                }
                                cellXfs += `<xf numFmtId="${numFmtId}" fontId="${fontid}" fillId="${fillid}" borderId="0" xfId="0" ${(fontid > 0 ? ' applyFont="1"' : '')}${(numFmtId > 0 ? ' applyNumberFormat="1"' : '')}${(fillid > 0 ? ' applyFill="1"' : '')}${(alignment ? ' applyAlignment="1"' : '')}>${alignment || ''}</xf>\n`;

                                cellXfsCount++;
                                s = cellXfsCount - 1;
                                if (styleID != null) {
                                    cellStyleMap[styleID] = s;
                                }
                            }
                        }
                        var isString = false;
                        var v = null;
                        var currentShareStringIndex = null;
                        var currentTxt = cell.Text;
                        if (currentTxt instanceof ShareString) {
                            //判断如果是共用文本，是否已经添加到列表中，如果是，则直接用id
                            if (currentTxt.__id != null) {
                                var strID = currentTxt.__id;
                                if (shareStringMap[strID] != null) {
                                    currentShareStringIndex = shareStringMap[strID];
                                }
                                currentTxt = currentTxt.txt;
                            }
                        }
                        if (typeof (currentTxt) == 'string') {
                            //如果不是共用文本，或者共用文本还没有添加到列表中，则添加到列表中
                            if (currentShareStringIndex == null) {
                                shareStringXml += `<si><t>${HtmlEncodeByRegExp(currentTxt)}</t></si>\n`;
                                currentShareStringIndex = shareStringIndex;
                                shareStringIndex++;
                                shareStringMap[strID] = currentShareStringIndex;
                            }
                            v = currentShareStringIndex;
                            isString = true;
                        } else if (typeof (currentTxt) == 'number') {
                            v = currentTxt
                        } else if (currentTxt instanceof Date) {

                        }
                        sheetDataXml += `<c r="${toColName(+col) + (+row + 1)}"${(isString ? ' t="s"' : '')} ${(s ? 's="' + s + '"' : '')}><v>${v}</v></c>\n`;
                        if (link != null) {
                            sheetRelsId++;
                            sheetXmlRels += `<Relationship Target="${link.Link}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Id="rId${sheetRelsId}" TargetMode="External"/>\n`
                            hyperlinksXml += `<hyperlink ref="${(toColName(col) + (+row + 1))}" r:id="rId${sheetRelsId}" />\n`;
                        }
                    }
                }
                sheetDataXml += `</row>\n`;
            }
            //列宽
            var colXml = '';
            for (var col in sheet.colWidth) {
                var width = sheet.colWidth[col];
                colXml += `<col min="${+col + 1}" max="${+col + 1}" width="${width}" customWidth="1" />\n`;
            }
            if (colXml != '') {
                sheetXml += `<cols>${colXml}</cols>`;
            }
            sheetXml += `<sheetData>${sheetDataXml}</sheetData>`;
            if (sheet.mergeCellDatas != null && sheet.mergeCellDatas.length > 0) {
                sheetXml += `<mergeCells count="${sheet.mergeCellDatas.length}">`;
                for (const mcd of sheet.mergeCellDatas) {
                    sheetXml += `<mergeCell ref="${toColName(mcd.fromCol) + (mcd.fromRow + 1)}:${toColName(mcd.toCol) + (mcd.toRow + 1)}" />\n`;
                }
                sheetXml += `</mergeCells>\n`;
            }
            if (hyperlinksXml != '') {
                sheetXml += `<hyperlinks>${hyperlinksXml}</hyperlinks>\n`;
            }
            sheetXml += `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />\n`;
            //加入图片
            if (sheet.ImageList != null && sheet.ImageList.length > 0) {
                var sheetImageCount = 0;
                var sheetImageIDMap = {};
                var drawingXml = `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`;
                drawingCount++;
                sheetRelsId++;
                for (const { img, col, row, width, height } of sheet.ImageList) {
                    var imgFormat = img.Option.Format;
                    var imgType = img.Option.Type;
                    var filename = `image${img.__id}.${imgFormat}`;
                    if (imageFileMap[filename] == null) {
                        var zipOpt = {};
                        if (typeof img.Data == 'string') {
                            zipOpt.base64 = true;
                        }
                        zip.file(`xl/media/${filename}`, img.Data, zipOpt);
                        imageFileMap[filename] = filename;
                    }
                    if (imageFormatMap[imgFormat] == null) {
                        drwaingContentTypeXml += `<Default Extension="${imgFormat}" ContentType="${ContentTypeMap[imgFormat]}" />\n`;
                        imageFormatMap[imgFormat] = imgFormat;
                    }
                    if (sheetImageIDMap[img.__id] == null) {
                        drawingXmlRels += `<Relationship Id="image${img.__id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${filename}" />\n`;
                        sheetImageIDMap[img.__id] = img.__id;
                    }
                    sheetImageCount++;
                    drawingXml +=
                        `<xdr:twoCellAnchor editAs="oneCell">
<xdr:from>
 <xdr:col>${col}</xdr:col>
 <xdr:colOff>0</xdr:colOff>
 <xdr:row>${row}</xdr:row>
 <xdr:rowOff>0</xdr:rowOff>
</xdr:from>
<xdr:to>
 <xdr:col>${col}</xdr:col>
 <xdr:colOff>${width * 10000}</xdr:colOff>
 <xdr:row>${row}</xdr:row>
 <xdr:rowOff>${height * 10000}</xdr:rowOff>
</xdr:to>
<xdr:pic>
 <xdr:nvPicPr>
  <xdr:cNvPr id="${sheetImageCount}" name="图片 ${sheetImageCount}" />
  <xdr:cNvPicPr>
   <a:picLocks noChangeAspect="1" />
  </xdr:cNvPicPr>
 </xdr:nvPicPr>
 <xdr:blipFill>
  <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="image${img.__id}">
   <a:extLst>
    <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
     <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0" />
    </a:ext>
   </a:extLst>
  </a:blip>
  <a:stretch>
   <a:fillRect />
  </a:stretch>
 </xdr:blipFill>
 <xdr:spPr>
  <a:xfrm>
   <a:off x="0" y="0" />
   <a:ext cx="0" cy="0" />
  </a:xfrm>
  <a:prstGeom prst="rect">
   <a:avLst />
  </a:prstGeom>
 </xdr:spPr>
</xdr:pic>
<xdr:clientData />
</xdr:twoCellAnchor>`
                }
                //
                drawingXmlRels += `</Relationships>`;
                zip.file(`xl/drawings/_rels/drawing${drawingCount}.xml.rels`, drawingXmlRels);
                drawingXml += `</xdr:wsDr>`;
                zip.file(`xl/drawings/drawing${drawingCount}.xml`, drawingXml);
                drwaingContentTypeXml += `<Override PartName="/xl/drawings/drawing${drawingCount}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" />\n`;
                sheetXmlRels += `<Relationship Id="rId${sheetRelsId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingCount}.xml" />`;

                sheetXml += `<drawing r:id="rId${sheetRelsId}" />\n`;
            }
            sheetXml += `</worksheet>`;
            zip.file(`xl/worksheets/sheet${sheetCount}.xml`, sheetXml);
            if (sheetXmlRels != '') {
                sheetXmlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${sheetXmlRels}</Relationships>`;
                zip.file(`xl/worksheets/_rels/sheet${sheetCount}.xml.rels`, sheetXmlRels);
            }
        }
        shareStringXml = `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shareStringIndex}" uniqueCount="${shareStringIndex}">
${shareStringXml}
</sst>
`;
        workbookXml += `</sheets></workbook>`;
        workbookXmlRels += `</Relationships>`;
        if (drwaingContentTypeXml != '') {
            ContentTypesXml += drwaingContentTypeXml;
        }
        ContentTypesXml += `</Types>`;
        stylesXml =
            `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
${numberFormatCount > 0 ? `<numFmts count="${numberFormatCount}">\n${numberFormatXml}</numFmts>` : ''}
 <fonts count="${styleFontsCount}" x14ac:knownFonts="1">
  ${styleFontsXml}
 </fonts>
 <fills count="${fillCount}">
 ${fillXml}
 </fills>
 <borders count="1">
  <border><left /><right /><top /><bottom /><diagonal /></border>
 </borders>
 <cellStyleXfs count="${cellStylexfsCount}">
  ${cellStyleXfs}
 </cellStyleXfs>
 <cellXfs count="${cellXfsCount}">
  ${cellXfs}
 </cellXfs>
 <cellStyles count="${cellStylesCount}">
  ${cellStyles}
 </cellStyles>
 <dxfs count="0" />
 <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9" />
 <extLst>
  <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
   <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1" />
  </ext>
  <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
   <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" />
  </ext>
 </extLst>
</styleSheet>
`

        zip.file('xl/sharedStrings.xml', shareStringXml);
        zip.file('xl/workbook.xml', workbookXml);
        zip.file('xl/_rels/workbook.xml.rels', workbookXmlRels);
        zip.file('[Content_Types].xml', ContentTypesXml);
        zip.file('xl/styles.xml', stylesXml);
        return zip;
    }

    const ColNameChar = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    function toColName(col) {
        col += 1;
        var name = '';
        while (true) {
            var i = (col - 1) % 26;
            name = ColNameChar[i] + name;
            col = Math.floor((col - 1) / 26);
            if (col === 0) {
                break;
            }
        }
        return name;
    }

    function makeFontXml(style) {
        return '<font>'
            + (style.Bold ? `<b />` : ``)
            + (style.Italic ? `<i />` : ``)
            + (style.Underline ? `<u />` : ``)
            + (style.FontSize ? `<sz val="${style.FontSize}" />` : ``)
            + (style.FontName ? `<name val="${style.FontName}" />` : ``)
            + (style.Color ? `<color ${typeof style.Color == 'number' ? `theme` : `rgb`}="${style.Color}" />` : '')
            + '</font>\n';
    }
    /*用正则表达式实现html转码*/
    function HtmlEncodeByRegExp(str) {
        var s = "";
        if (str.length == 0) return "";
        s = str.replace(/&/g, "&amp;");
        s = s.replace(/</g, "&lt;");
        s = s.replace(/>/g, "&gt;");
        // s = s.replace(/ /g, "&nbsp;");
        s = s.replace(/\'/g, "&#39;");
        s = s.replace(/\"/g, "&quot;");
        return s;
    }
    /*4.用正则表达式实现html解码
    htmlDecodeByRegExp:function (str){  
          var s = "";
          if(str.length == 0) return "";
          s = str.replace(/&amp;/g,"&");
          s = s.replace(/&lt;/g,"<");
          s = s.replace(/&gt;/g,">");
          s = s.replace(/&nbsp;/g," ");
          s = s.replace(/&#39;/g,"\'");
          s = s.replace(/&quot;/g,"\"");
          return s;  
    }*/
})();