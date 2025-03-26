/**
 * Copyright (C) 2011-2015 The XDocReport Team <xdocreport@googlegroups.com>
 *
 * All rights reserved.
 *
 * Permission is hereby granted, free  of charge, to any person obtaining
 * a  copy  of this  software  and  associated  documentation files  (the
 * "Software"), to  deal in  the Software without  restriction, including
 * without limitation  the rights to  use, copy, modify,  merge, publish,
 * distribute,  sublicense, and/or sell  copies of  the Software,  and to
 * permit persons to whom the Software  is furnished to do so, subject to
 * the following conditions:
 *
 * The  above  copyright  notice  and  this permission  notice  shall  be
 * included in all copies or substantial portions of the Software.
 *
 * THE  SOFTWARE IS  PROVIDED  "AS  IS", WITHOUT  WARRANTY  OF ANY  KIND,
 * EXPRESS OR  IMPLIED, INCLUDING  BUT NOT LIMITED  TO THE  WARRANTIES OF
 * MERCHANTABILITY,    FITNESS    FOR    A   PARTICULAR    PURPOSE    AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE,  ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package fr.opensagres.poi.xwpf.converter.xhtml.internal;

import static fr.opensagres.poi.xwpf.converter.core.utils.DxaUtil.emu2points;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.XHTMLConstants.*;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.HEIGHT;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.MARGIN_BOTTOM;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.MARGIN_LEFT;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.MARGIN_RIGHT;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.MARGIN_TOP;
import static fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants.WIDTH;

import java.io.IOException;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.microsoft.schemas.vml.CTImageData;
import com.microsoft.schemas.vml.CTShape;
import fr.opensagres.poi.xwpf.converter.core.*;
import fr.opensagres.poi.xwpf.converter.core.styles.run.RunFontStyleDStrikeValueProvider;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSProperty;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFSDT;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromH.Enum;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STTwipsMeasure;
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;

import fr.opensagres.poi.xwpf.converter.core.styles.XWPFStylesDocument;
import fr.opensagres.poi.xwpf.converter.core.styles.run.RunFontStyleStrikeValueProvider;
import fr.opensagres.poi.xwpf.converter.core.styles.run.RunTextHighlightingValueProvider;
import fr.opensagres.poi.xwpf.converter.core.utils.DxaUtil;
import fr.opensagres.poi.xwpf.converter.core.utils.StringUtils;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStyle;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylePropertyConstants;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.styles.CSSStylesDocument;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.utils.SAXHelper;
import fr.opensagres.poi.xwpf.converter.xhtml.internal.utils.StringEscapeUtils;

public class XHTMLMapper
    extends XWPFDocumentVisitor<Object, XHTMLOptions, XHTMLMasterPage>
{
	
	/**
	 * There is no HTML representation for tab. So apply 4 spaces by default
	 */
	static final String TAB_CHAR_SEQUENCE = "&nbsp;&nbsp;&nbsp;&nbsp;";

    private static final String WORD_MEDIA = "word/media/";

    private final ContentHandler contentHandler;
    
    /**
     * To hold paragraph reference and to be used while processing individual runs which has tabs
     */
    private XWPFParagraph currentParagraph;

    private boolean generateStyles;

    private final IURIResolver resolver;

    private AttributesImpl currentRunAttributes;

    private boolean pageDiv;

    public XHTMLMapper( XWPFDocument document, ContentHandler contentHandler, XHTMLOptions options )
        throws Exception
    {
        super( document, options != null ? options : XHTMLOptions.getDefault() );
        this.contentHandler = contentHandler;
        this.resolver = getOptions().getURIResolver();
        this.generateStyles = getOptions().isGenerateStyles();
        this.pageDiv = false;
    }

    @Override
    protected XWPFStylesDocument createStylesDocument( XWPFDocument document )
        throws XmlException, IOException
    {
        return new CSSStylesDocument( document, options.isIgnoreStylesIfUnused(), options.getIndent() );
    }

    @Override
    protected Object startVisitDocument()
        throws Exception
    {
        if ( !options.isFragment() )
        {
            contentHandler.startDocument();
            // html start
            startElement( HTML_ELEMENT );
            // head start
            startElement( HEAD_ELEMENT );
            if ( generateStyles )
            {
                // styles
                ( (CSSStylesDocument) stylesDocument ).save( contentHandler );
            }
            // html end
            endElement( HEAD_ELEMENT );
            // body start
            startElement( BODY_ELEMENT );
        }
        return null;
    }

    @Override
    protected void endVisitDocument()
        throws Exception
    {
        if ( pageDiv )
        {
            endElement( DIV_ELEMENT );
        }
        if ( !options.isFragment() )
        {
            // body end
            endElement( BODY_ELEMENT );
            // html end
            endElement( HTML_ELEMENT );
            contentHandler.endDocument();
        }
    }

    @Override
    protected Object startVisitSDT(XWPFSDT contents, Object container) throws SAXException {

        startElement(DIV_ELEMENT, null);
//        startElement(PRE_ELEMENT, null);
        return null;
    }

//    @Override
//    protected void visitSDTBody(XWPFSDT contents, Object sdtContainer) throws SAXException {
//        characters(contents.getContent().getText());
//    }

    @Override
    protected void endVisitSDT(XWPFSDT contents, Object container, Object sdtContainer) throws SAXException {
//        endElement(PRE_ELEMENT);
        endElement(DIV_ELEMENT);
    }

    @Override
    protected Object startVisitParagraph( XWPFParagraph paragraph, ListItemContext itemContext, Object parentContainer )
        throws Exception
    {
        // 1) create attributes

        // 1.1) Create "class" attributes.
        AttributesImpl attributes = createClassAttribute( paragraph.getStyleID() );

        // 1.2) Create "style" attributes.
        CTPPr pPr = paragraph.getCTP().getPPr();
        CTStyle style = this.getStylesDocument().getStyle(paragraph.getStyleID());
        CTDecimalNumber outlineLvl = null;
        if (style!=null){
            CTPPrGeneral stylePPr = style.getPPr();
            if (stylePPr!=null){
                outlineLvl = stylePPr.getOutlineLvl();
            }
        }
        // 1.3)创建层级样式
        if (pPr!=null && outlineLvl==null){
            outlineLvl = pPr.getOutlineLvl();
            if (outlineLvl!=null && attributes!=null){
                int classIndex = attributes.getIndex("class");
                String value = attributes.getValue(classIndex);
                value+=" "+"outlineLvl-"+outlineLvl.getVal();
                attributes.setValue(classIndex,value);
            }
        }
        CSSStyle cssStyle = getStylesDocument().createCSSStyle( pPr );
        if (cssStyle == null && attributes != null) {
            cssStyle = new CSSStyle( P_ELEMENT, null );
        }
        if(cssStyle != null) {
            cssStyle.addProperty(CSSStylePropertyConstants.WHITE_SPACE, "pre-wrap");            
        }
        attributes = createStyleAttribute( cssStyle, attributes );

        // 2) create element
        //针对1级层级设置标题1样式
        if (outlineLvl!=null){
            String headTag = fetchTitle(outlineLvl.getVal().intValue());

            startElement( headTag, attributes );
        }else {
            startElement( P_ELEMENT, attributes );
        }

        //To handle list items in paragraph
        if(itemContext != null)
        {			
	        startElement( SPAN_ELEMENT, attributes );
	        String text = itemContext.getText();	        
	        if ( StringUtils.isNotEmpty( text ) )
	        {	
	        	text = StringUtils.replaceNonUnicodeChars(text);
	        	text = text + "\u0020";
	        	SAXHelper.characters( contentHandler, StringEscapeUtils.escapeHtml( text ) );
	        }
	        endElement( SPAN_ELEMENT );
		}
        return null;
    }

    @Override
    protected void endVisitParagraph( XWPFParagraph paragraph, Object parentContainer, Object paragraphContainer )
        throws Exception
    {
        // 1.2) Create "style" attributes.
        CTPPr pPr = paragraph.getCTP().getPPr();
        CTDecimalNumber outlineLvl = null;
        CTStyle style = this.getStylesDocument().getStyle(paragraph.getStyleID());
        if (style!=null){
            CTPPrGeneral stylePPr = style.getPPr();
            if (stylePPr!=null){
                outlineLvl = stylePPr.getOutlineLvl();
            }
        }
        // 1.3)创建层级样式
        if (pPr!=null && outlineLvl==null){
            outlineLvl = pPr.getOutlineLvl();
        }
        //针对1级层级设置标题1样式
        if (outlineLvl!=null){
            String headTag = fetchTitle(outlineLvl.getVal().intValue());
            endElement(headTag);
        }else {
            endElement( P_ELEMENT);
        }
    }

    @Override
    protected void visitRun( XWPFRun run, boolean pageNumber, String url, Object paragraphContainer )
        throws Exception
    {	    	
		if(run.getParent() instanceof XWPFParagraph) {
			this.currentParagraph = (XWPFParagraph) run.getParent();
		}

        XWPFParagraph paragraph = run.getParagraph();
        // 1) create attributes

        // 1.1) Create "class" attributes.
        this.currentRunAttributes = createClassAttribute( paragraph.getStyleID() );

        // 1.2) Create "style" attributes.
        CTRPr rPr = run.getCTR().getRPr();
        CSSStyle cssStyle = getStylesDocument().createCSSStyle( rPr );
        if (cssStyle == null && this.currentRunAttributes != null) {
            cssStyle = new CSSStyle( SPAN_ELEMENT, null );
        }
        if (cssStyle != null) {
            cssStyle.addProperty(CSSStylePropertyConstants.WHITE_SPACE, "pre-wrap");
        }
        this.currentRunAttributes = createStyleAttribute( cssStyle, currentRunAttributes );

        if ( url != null )
        {
            // url is not null, generate a HTML a.
            AttributesImpl hyperlinkAttributes = new AttributesImpl();
            SAXHelper.addAttrValue( hyperlinkAttributes, HREF_ATTR, url );
            startElement( A_ELEMENT, hyperlinkAttributes );
        }

        super.visitRun( run, pageNumber, url, paragraphContainer );

        if ( url != null )
        {
            // url is not null, close the HTML a.
            // TODO : for the moment generate space to be ensure that a has some content.
            characters( " " );
            endElement( A_ELEMENT );
        }
        this.currentRunAttributes = null;
        this.currentParagraph = null;
    }

    @Override
    protected void visitEmptyRun( Object paragraphContainer )
        throws Exception
    {
        startElement( BR_ELEMENT );
        endElement( BR_ELEMENT );
    }

    @Override
    protected void visitText( CTText ctText, boolean pageNumber, Object paragraphContainer )
        throws Exception
    {
        if ( currentRunAttributes != null )
        {
            startElement( SPAN_ELEMENT, currentRunAttributes );
        }
        String text = ctText.getStringValue();
        if ( StringUtils.isNotEmpty( text ) )
        {
            // Escape with HTML characters
            characters( StringEscapeUtils.escapeHtml( text ) );
        }
        // else
        // {
        // characters( SPACE_ENTITY );
        // }
        if ( currentRunAttributes != null )
        {
            endElement( SPAN_ELEMENT );
        }
    }
    
    @Override
    protected void visitStyleText(XWPFRun run, String text, Object parent, boolean pageNumber) throws Exception
    {
        //添加字体样式的判断 兼容汉字
    	if(run.getFontFamily() == null) {
            String fontFamilyAscii = getStylesDocument().getFontFamilyAscii(run);
            String fontFamil = fontFamilyAscii;
            String fontFamilyEastAsia = getStylesDocument().getFontFamilyEastAsia(run);
            String fontFamilyHAnsi = getStylesDocument().getFontFamilyHAnsi(run);
            if (fontFamilyHAnsi!=null || "eastAsia".equals(fontFamilyHAnsi)){
                fontFamil = fontFamilyEastAsia;
            }
            run.setFontFamily(fontFamil);
        }
    	
		if(run.getFontSize() <= 0) {
			run.setFontSize(getStylesDocument().getFontSize(run).intValue());
		}
		
		CTRPr rPr = run.getCTR().getRPr();
		
    	// 1) create attributes

        // 1.1) Create "class" attributes.
        AttributesImpl runAttributes = createClassAttribute( currentParagraph.getStyleID() );

        // 1.2) Create "style" attributes.
        CSSStyle cssStyle = getStylesDocument().createCSSStyle( rPr );
        if (cssStyle == null && runAttributes != null) {
            cssStyle = new CSSStyle( SPAN_ELEMENT, null );
        }
        if(cssStyle != null) {
        	Color color = RunTextHighlightingValueProvider.INSTANCE.getValue(rPr, getStylesDocument());
        	if(color != null) cssStyle.addProperty(CSSStylePropertyConstants.BACKGROUND_COLOR, StringUtils.toHexString(color));
        	if(Boolean.TRUE.equals(RunFontStyleStrikeValueProvider.INSTANCE.getValue(rPr, getStylesDocument())) ||
                    Boolean.TRUE.equals(RunFontStyleDStrikeValueProvider.INSTANCE.getValue(rPr, getStylesDocument())))
        		cssStyle.addProperty("text-decoration", "line-through");
        	if(rPr.sizeOfVertAlignArray() > 0) {
        		int align = rPr.getVertAlignArray(0).getVal().intValue();
        		if(STVerticalAlignRun.INT_SUPERSCRIPT == align) {
        			cssStyle.addProperty("vertical-align", "super");
        		}
        		else if(STVerticalAlignRun.INT_SUBSCRIPT == align) {
        			cssStyle.addProperty("vertical-align", "sub");
        		}
        	}	        		
        }
        runAttributes = createStyleAttribute( cssStyle, runAttributes );
        if ( runAttributes != null )
        {
            startElement( SPAN_ELEMENT, runAttributes );
        }        
        if ( StringUtils.isNotEmpty( text ) )
        {
            // Escape with HTML characters
            characters( StringEscapeUtils.escapeHtml( text ) );
        }
        if ( runAttributes != null )
        {
            endElement( SPAN_ELEMENT );
        } 
    }

    @Override
    protected void visitTab( CTPTab o, Object paragraphContainer )
        throws Exception
    {
    }

    @Override
    protected void visitTabs( CTTabs tabs, Object paragraphContainer )
        throws Exception
    {
    	//For some reason tabs become null ???
    	//Add equivalent spaces in html render as no tab in html world
    	if(currentParagraph != null && tabs == null)
    	{
			startElement( SPAN_ELEMENT, null );
			characters(TAB_CHAR_SEQUENCE);
			endElement(SPAN_ELEMENT);
			return;
		}
    	else if(currentParagraph != null && tabs != null)
    	{
            characters(TAB_CHAR_SEQUENCE);
        }
    }

    @Override
    protected void addNewLine( CTBr br, Object paragraphContainer )
        throws Exception
    {
        startElement( BR_ELEMENT );
        endElement( BR_ELEMENT );
    }

    @Override
    protected void pageBreak()
        throws Exception
    {
    }

    @Override
    protected void visitBookmark( CTBookmark bookmark, XWPFParagraph paragraph, Object paragraphContainer )
        throws Exception
    {
        AttributesImpl attributes = new AttributesImpl();
        SAXHelper.addAttrValue( attributes, ID_ATTR, bookmark.getName() );
        startElement( SPAN_ELEMENT, attributes );
        endElement( SPAN_ELEMENT );
    }

    @Override
    protected Object startVisitTable( XWPFTable table, float[] colWidths, Object tableContainer )
        throws Exception
    {
        // 1) create attributes
        // 1.1) Create class attributes.
        AttributesImpl attributes = createClassAttribute( table.getStyleID() );

        // 1.2) Create "style" attributes.
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CSSStyle cssStyle = getStylesDocument().createCSSStyle( tblPr );
        if(cssStyle != null) {
        	cssStyle.addProperty(CSSStylePropertyConstants.BORDER_COLLAPSE, CSSStylePropertyConstants.BORDER_COLLAPSE_COLLAPSE);
        }
        attributes = createStyleAttribute( cssStyle, attributes );

        // 2) create element
        startElement( TABLE_ELEMENT, attributes );
        return null;
    }

    @Override
    protected void endVisitTable( XWPFTable table, Object parentContainer, Object tableContainer )
        throws Exception
    {
        endElement( TABLE_ELEMENT );
    }

    @Override
    protected void startVisitTableRow( XWPFTableRow row, Object tableContainer, int rowIndex, boolean headerRow )
        throws Exception
    {

        // 1) create attributes
        // Create class attributes.
        XWPFTable table = row.getTable();
        AttributesImpl attributes = createClassAttribute( table.getStyleID() );

        // 2) create element
        if ( headerRow )
        {
            startElement( TH_ELEMENT, attributes );
        }
        else
        {
            startElement( TR_ELEMENT, attributes );
        }
    }

    @Override
    protected void endVisitTableRow( XWPFTableRow row, Object tableContainer, boolean firstRow, boolean lastRow,
                                     boolean headerRow )
        throws Exception
    {
        if ( headerRow )
        {
            endElement( TH_ELEMENT );
        }
        else
        {
            endElement( TR_ELEMENT );
        }
    }

    @Override
    protected Object startVisitTableCell( XWPFTableCell cell, Object tableContainer, boolean firstRow, boolean lastRow,
                                          boolean firstCell, boolean lastCell, List<XWPFTableCell> vMergeCells )
        throws Exception
    {
        // 1) create attributes
        // 1.1) Create class attributes.
        XWPFTableRow row = cell.getTableRow();
        XWPFTable table = row.getTable();
        AttributesImpl attributes = createClassAttribute( table.getStyleID() );

        // 1.2) Create "style" attributes.
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        CSSStyle cssStyle = getStylesDocument().createCSSStyle( tcPr );
        //At lease support solid borders for now
        if(cssStyle != null) {
        	TableCellBorder border = getStylesDocument().getTableBorder(table, BorderSide.TOP);
        	if(border != null)
        	{
        		String style = border.getBorderSize() + "px solid " +StringUtils.toHexString(border.getBorderColor()); 
            	cssStyle.addProperty(CSSStylePropertyConstants.BORDER_TOP, style);
        	}        	
        	
        	border = getStylesDocument().getTableBorder(table, BorderSide.BOTTOM);
        	if(border != null)
        	{
        		String style = border.getBorderSize() + "px solid " + StringUtils.toHexString(border.getBorderColor());         	
            	cssStyle.addProperty(CSSStylePropertyConstants.BORDER_BOTTOM, style);
        	}        	
        	
        	border = getStylesDocument().getTableBorder(table, BorderSide.LEFT);
        	if(border != null)
        	{
        		String style = border.getBorderSize() + "px solid " + StringUtils.toHexString(border.getBorderColor());
            	cssStyle.addProperty(CSSStylePropertyConstants.BORDER_LEFT, style);
        	}        	
        	
        	border = getStylesDocument().getTableBorder(table, BorderSide.RIGHT);
        	if(border != null)
        	{
        		String style = border.getBorderSize() + "px solid " + StringUtils.toHexString(border.getBorderColor());
            	cssStyle.addProperty(CSSStylePropertyConstants.BORDER_RIGHT, style);
        	}        	
        }
        attributes = createStyleAttribute( cssStyle, attributes );

        // colspan attribute
        BigInteger gridSpan = stylesDocument.getTableCellGridSpan( cell );
        if ( gridSpan != null )
        {
            attributes = SAXHelper.addAttrValue( attributes, COLSPAN_ATTR, gridSpan.intValue() );
        }

        if ( vMergeCells != null )
        {
            attributes = SAXHelper.addAttrValue( attributes, ROWSPAN_ATTR, vMergeCells.size() );
        }

        // 2) create element
        startElement( TD_ELEMENT, attributes );

        return null;
    }

    @Override
    protected void endVisitTableCell( XWPFTableCell cell, Object tableContainer, Object tableCellContainer )
        throws Exception
    {
        endElement( TD_ELEMENT );
    }

    @Override
    protected void visitHeader( XWPFHeader header, CTHdrFtrRef headerRef, CTSectPr sectPr, XHTMLMasterPage masterPage )
        throws Exception
    {
        // TODO Auto-generated method stub

    }

    @Override
    protected void visitFooter( XWPFFooter footer, CTHdrFtrRef footerRef, CTSectPr sectPr, XHTMLMasterPage masterPage )
        throws Exception
    {
        // TODO Auto-generated method stub

    }

    @Override
    protected void visitPicture( CTPicture picture,
                                 Float offsetX,
                                 Enum relativeFromH,
                                 Float offsetY,
                                 org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STRelFromV.Enum relativeFromV,
                                 org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.STWrapText.Enum wrapText,
                                 Object parentContainer )
        throws Exception
    {

        AttributesImpl attributes = null;
        // Src attribute
        XWPFPictureData pictureData = super.getPictureData( picture );
        if ( pictureData != null )
        {
            // img/@src
            String src = pictureData.getFileName();
            if ( StringUtils.isNotEmpty( src ) )
            {
                src = resolver.resolve( WORD_MEDIA + src );
                attributes = SAXHelper.addAttrValue( attributes, SRC_ATTR, src );
            }

            CTPositiveSize2D ext = picture.getSpPr().getXfrm().getExt();
            CSSStyle style = new CSSStyle( IMG_ELEMENT, null );
            // img/@width
            float width = emu2points( ext.getCx() );
            // img/@height
            float height = emu2points( ext.getCy() );
            style.addProperty(WIDTH, getStylesDocument().getValueAsPoint( width ) );
            style.addProperty(HEIGHT, getStylesDocument().getValueAsPoint( height ) );
            attributes = SAXHelper.addAttrValue( attributes, STYLE_ATTR, style.getInlineStyles() );
        }
        else 
        {    
        	// external link images inserted
        	String link = picture.getBlipFill().getBlip().getLink();
            String src = document.getPackagePart().getRelationships().getRelationshipByID(link).getTargetURI().toString();
        	attributes = SAXHelper.addAttrValue( null, SRC_ATTR, src );
        	
        	CTPositiveSize2D ext = picture.getSpPr().getXfrm().getExt();
        	CSSStyle style = new CSSStyle( IMG_ELEMENT, null );
            // img/@width
            float width = emu2points( ext.getCx() );
            // img/@height
            float height = emu2points( ext.getCy() );
            style.addProperty(WIDTH, getStylesDocument().getValueAsPoint( width ) );
            style.addProperty(HEIGHT, getStylesDocument().getValueAsPoint( height ) );
            attributes = SAXHelper.addAttrValue( attributes, STYLE_ATTR, style.getInlineStyles() );
        }
        if ( attributes != null )
        {
            startElement( IMG_ELEMENT, attributes );
            endElement( IMG_ELEMENT );
        }
    }

    @Override
    protected void visitVmlPicture(org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPicture picture,
                                   Object pdfParentContainer) throws Exception {
        XmlCursor pictureCur = picture.newCursor();
        pictureCur.selectPath("./*");
        while (pictureCur.toNextSelection()) {
            XmlObject obj = pictureCur.getObject();
            if (obj instanceof CTShape) {
                CTShape shape = (CTShape) obj;

                List<CTImageData> imagedataList = shape.getImagedataList();
                for (CTImageData imageData : imagedataList) {
                    XWPFPictureData pictureData = getPictureDataByID(imageData.getId2());
                    visitVmlPicture(pictureData, shape.getStyle(), pdfParentContainer);
                }
            }
        }
        pictureCur.dispose();
    }

    protected void visitVmlPicture(
            XWPFPictureData pictureData, String style, Object pdfParentContainer) throws Exception{
        if (pictureData == null) {
            return;
        }
        IImageExtractor extractor = this.getImageExtractor();
        extractor.extract(WORD_MEDIA + pictureData.getFileName(), pictureData.getData());
        String src;
        AttributesImpl attributes = null;
        src = pictureData.getFileName();
        if (StringUtils.isNotEmpty(src)) {
            src = this.resolver.resolve(WORD_MEDIA + src);
            attributes = SAXHelper.addAttrValue(attributes, SRC_ATTR, src);
        }
        attributes = SAXHelper.addAttrValue(attributes, STYLE_ATTR, style);
        if ( attributes != null )
        {
            startElement( IMG_ELEMENT, attributes );
            endElement( IMG_ELEMENT );
        }
    }

    public void setActiveMasterPage( XHTMLMasterPage masterPage )
    {
        if ( pageDiv )
        {
            try
            {
                endElement( DIV_ELEMENT );
            }
            catch ( SAXException e )
            {
                e.printStackTrace();
            }
        }
        AttributesImpl attributes = new AttributesImpl();
        CSSStyle style = new CSSStyle( DIV_ELEMENT, null );
        CTSectPr sectPr = masterPage.getSectPr();
        CTPageSz pageSize = sectPr.getPgSz();
        if ( pageSize != null )
        {
            // Width
        	STTwipsMeasure width = pageSize.xgetW();
            if ( width != null )
            {
                style.addProperty( WIDTH, getStylesDocument().getValueAsPoint( DxaUtil.dxa2points( width ) ) );
            }
        }

        CTPageMar pageMargin = sectPr.getPgMar();
        if ( pageMargin != null )
        {
            // margin bottom
        	STSignedTwipsMeasure marginBottom = pageMargin.xgetBottom();
            if ( marginBottom != null )
            {
                float marginBottomPt = DxaUtil.dxa2points( marginBottom );
                style.addProperty( MARGIN_BOTTOM, getStylesDocument().getValueAsPoint( marginBottomPt ) );
            }
            // margin top
            STSignedTwipsMeasure marginTop = pageMargin.xgetTop();
            if ( marginTop != null )
            {
                float marginTopPt = DxaUtil.dxa2points( marginTop );
                style.addProperty( MARGIN_TOP, getStylesDocument().getValueAsPoint( marginTopPt ) );
            }
            // margin left
            STTwipsMeasure marginLeft = pageMargin.xgetLeft();
            if ( marginLeft != null )
            {
                float marginLeftPt = DxaUtil.dxa2points( marginLeft );
                style.addProperty( MARGIN_LEFT, getStylesDocument().getValueAsPoint( marginLeftPt ) );
            }
            // margin right
            STTwipsMeasure marginRight = pageMargin.xgetRight();
            if ( marginRight != null )
            {
                float marginRightPt = DxaUtil.dxa2points( marginRight );
                style.addProperty( MARGIN_RIGHT, getStylesDocument().getValueAsPoint( marginRightPt ) );
            }
        }
        String s = style.getInlineStyles();
        if ( StringUtils.isNotEmpty( s ) )
        {
            SAXHelper.addAttrValue( attributes, STYLE_ATTR, s );
        }
        try
        {
            startElement( DIV_ELEMENT, attributes );
        }
        catch ( SAXException e )
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        pageDiv = true;

    }

    public XHTMLMasterPage createMasterPage( CTSectPr sectPr )
    {
        return new XHTMLMasterPage( sectPr );
    }

    private void startElement( String name )
        throws SAXException
    {
        startElement( name, null );
    }

    private void startElement( String name, Attributes attributes )
        throws SAXException
    {
        SAXHelper.startElement( contentHandler, name, attributes );
    }

    private void endElement( String name )
        throws SAXException
    {
        SAXHelper.endElement( contentHandler, name );
    }

    private void characters( String content )
        throws SAXException
    {
        SAXHelper.characters( contentHandler, content );
    }

    @Override
    public CSSStylesDocument getStylesDocument()
    {
        return (CSSStylesDocument) super.getStylesDocument();
    }

    private AttributesImpl createClassAttribute( String styleID )
    {
        String classNames = getStylesDocument().getClassNames( styleID );
        if ( StringUtils.isNotEmpty( classNames ) )
        {
            return SAXHelper.addAttrValue( null, CLASS_ATTR, classNames );
        }
        return null;
    }

    private AttributesImpl createStyleAttribute( CSSStyle cssStyle, AttributesImpl attributes )
    {
        if ( cssStyle != null )
        {
            String inlineStyles = cssStyle.getInlineStyles();
            if ( StringUtils.isNotEmpty( inlineStyles ))
            {
                //类样式生成在行内
                if (attributes!=null && !generateStyles){
                    int classIndex = attributes.getIndex("class");
                    String value = attributes.getValue(classIndex);
                    String[] styleClass = value.split(" ");
                    List<CSSStyle> cssStyles = this.getStylesDocument().getCSSStyles();
                    Map<String,String> map = new HashMap<>();
                    //类样式处理到style内
                    for (String className : styleClass) {
                        for (CSSStyle style : cssStyles) {
                            if (style.getTagName().equals(cssStyle.getTagName())&&
                                (style.getClassName()==null||
                                (style.getClassName()!=null && style.getClassName().equals(className)))){
                                List<CSSProperty> properties = style.getProperties();
                                for (CSSProperty property : properties) {
                                    map.put(property.getName(),property.getValue());
                                }
                            }
                        }
                    }
                    //替换类样式
                    for (String k : map.keySet()) {
                        String v = map.get(k);
                        cssStyle.replaceStyle(k,v,false);
                    }

                    inlineStyles = cssStyle.getInlineStyles();
                }
                attributes = SAXHelper.addAttrValue( attributes, STYLE_ATTR, inlineStyles );
            }
        }
        return attributes;
    }

    private String fetchTitle(int outlineLvl){
        String headTag = P_ELEMENT;
        switch (outlineLvl){
            case 0:
                headTag = H1_ELT;
                break;
            case 1:
                headTag = H2_ELT;
                break;
            case 2:
                headTag = H3_ELT;
                break;
            case 3:
                headTag = H4_ELT;
                break;
            case 4:
                headTag = H5_ELT;
                break;
            case 5:
                headTag = H6_ELT;
                break;
            default:
                break;
        }
        return headTag;
    }
}
