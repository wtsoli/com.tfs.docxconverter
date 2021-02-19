package com.tfs.sample;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringEscapeUtils;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.ImportXHTMLProperties;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.RFonts;

/**
 * This sample converts XHTML to docx content.
 * 
 * Beware that a file created with a Microsoft text editor
 * will start with a byte order mark (BOM):
 * 
 *    http://msdn.microsoft.com/en-us/library/windows/desktop/dd374101(v=vs.85).aspx
 * 
 * and if this is converted to a String, it can result in 
 * "Content not allowed in prolog" error.
 * 
 * So it is preferable to use one of the XHTMLImporter.convert
 * signatures which doesn't use a String (eg File or InputStream).
 * 
 * Here a string may be used for convenience where the XHTML is escaped 
 * (as required for OpenDoPE input), so it can be unescaped first.
 *
 * For best results, be sure to include src/main/resources on your classpath.
 *  
 */
public class ConvertInXHTMLFile {

    public static void main(String[] args) throws Exception {
        
    	
        String inputfilepath = System.getProperty("user.dir") + "/example.html";
        
     // Images: provide correct baseURL
    	String baseURL = "file:///bvols/@git/repos/docx4j-ImportXHTML/sample-docs/docx/sample-docxv2.docx_files";    	
//        String baseURL = "file:///C:/Users/jharrop/git/docx4j-ImportXHTML/sample-docs/docx/sample-docxv2.docx_files";

        
        String stringFromFile = FileUtils.readFileToString(new File(inputfilepath), "UTF-8");
        
        String unescaped = stringFromFile;
//        if (stringFromFile.contains("&lt;/") ) {
//    		unescaped = StringEscapeUtils.unescapeHtml(stringFromFile);        	
//        }
        
		
//        XHTMLImporter.setTableFormatting(FormattingOption.IGNORE_CLASS);
//        XHTMLImporter.setParagraphFormatting(FormattingOption.IGNORE_CLASS);
        
		System.out.println("Unescaped: " + unescaped);
        
                
        // Setup font mapping
		RFonts rfonts = Context.getWmlObjectFactory().createRFonts();
		rfonts.setAscii("Century Gothic");
        XHTMLImporterImpl.addFontMapping("Century Gothic", rfonts);
        
        // Create an empty docx package
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
//		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "/styled.docx"));

		
		NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
		wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
		ndp.unmarshalDefaultNumbering();		
					
		// Convert the XHTML, and add it into the empty docx we made
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        
        XHTMLImporter.setHyperlinkStyle("Hyperlink");
		wordMLPackage.getMainDocumentPart().getContent().addAll( 
				XHTMLImporter.convert(unescaped, baseURL) );
		
		System.out.println(
				XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));

//		System.out.println(
//				XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getNumberingDefinitionsPart().getJaxbElement(), true, true));
		
		wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/example.docx") );
      
  }
	
}