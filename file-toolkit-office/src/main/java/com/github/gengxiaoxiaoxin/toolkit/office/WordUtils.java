package com.github.gengxiaoxiaoxin.toolkit.office;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import javax.xml.bind.JAXBException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

/**
 * WordUtils
 */
public class WordUtils {
	
	/**
	 * replace the word variable
	 * @param mappings mappings
	 * @param inputStream inputStream
	 * @param outputStream outputStream
	 * @throws JAXBException
	 * @throws Docx4JException
	 */
	public static void writeVariable(Map<String, String> mappings, InputStream inputStream, OutputStream outputStream) throws JAXBException, Docx4JException {
		//docx4j  docx转pdf
		//Mapper fontMapper = new IdentityPlusMapper();
		//PhysicalFonts.put("宋体", PhysicalFonts.get("SimSun"));
		//PhysicalFonts.put("新細明體", PhysicalFonts.get("Microsoft YaHei UI Bold"));
		//fontMapper.put("宋体", PhysicalFonts.get("宋体"));
		//Load之后，PhysicalFonts才会读取系统自带的字体，这时候SimSun才能读到
		
		//fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft YaHei"));
		//fontMapper.registerBoldForm("微软雅黑", PhysicalFonts.get("Microsoft YaHei UI Bold"));
		//fontMapper.registerBoldForm("宋体", PhysicalFonts.get("Microsoft YaHei UI Bold"));
		//解决宋体（正文）和宋体（标题）的乱码问题
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputStream);
		//wordMLPackage.setFontMapper(fontMapper);
		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
		mainDocumentPart.variableReplace(mappings);
		Docx4J.save(wordMLPackage, outputStream);
	}
	
	
}
