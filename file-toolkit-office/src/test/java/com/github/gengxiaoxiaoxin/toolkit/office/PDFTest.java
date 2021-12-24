package com.github.gengxiaoxiaoxin.toolkit.office;

import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class PDFTest {
	
	@Test
	public void testPdfBox() throws Exception {
		// Loading an existing document
		File file = new File("C:\\Users\\EDZ\\Desktop\\结算确认函-月结-模板.pdf");
		PDDocument document = PDDocument.load(file);
		
		// Retrieving the pages of the document, and using PREPEND mode
		PDPage page = document.getPage(0);
		PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true);
		
		TrueTypeCollection trueTypeCollection = new TrueTypeCollection(new File("c:/windows/fonts/simsun.ttc"));
		PDType0Font font = PDType0Font.load(document, trueTypeCollection.getFontByName("SimSun"), false);
		// Setting the font to the Content stream
		//小三=15
		contentStream.setFont(font, 15);
		
		PdfBoxUtils.writeTextByCm(contentStream, "北京品诺优创科技有限公司", 4.24f, 23.12f, false);
		PdfBoxUtils.writeTextByCm(contentStream, "2021年10月账单结算金额73,276.82元，", 8.34f, 21f, false);
//		PdfBoxUtils.writeTextByCm(contentStream, "北京品诺优创科技有限公司", 4.24f, 23.12f, false);
//		PdfBoxUtils.writeTextByCm(contentStream, "北京品诺优创科技有限公司", 4.24f, 23.12f, false);
		
		// Closing the content stream
		contentStream.close();
		// Saving the document
		document.save(new File("C:\\Users\\EDZ\\Desktop\\666.pdf"));
		// Closing the document
		document.close();
	}
	
	@Test
	public void testDoc4j2Doc() {
		// 获取模板
		Map<String, String> mappings = new HashMap<>();
		mappings.put("subjectName", "北京品诺优创科技有限公司");
		mappings.put("periodYear", String.valueOf(2021));
		mappings.put("periodMonth", String.valueOf(10));
		mappings.put("settleAmount", "73,276.82");
		mappings.put("billingDate", "2021-11-5");
		mappings.put("supplierName", "陕西振彰食品有限公司");
		mappings.put("supplierAccountName", "陕西振彰食品有限公司");
		mappings.put("supplierBankAccount", "292083207710001");
		mappings.put("supplierBankName", "招行南大街支行");
		//subject
		mappings.put("subjectAccountName", "北京品诺优创科技有限公司");
		mappings.put("subjectTaxId", "91110105MA009T9F5J");
		mappings.put("subjectAddressPhone", "北京市朝阳区建国路89号院15号楼13层1602室010-85967767");
		mappings.put("subjectBankNameBankAccount", "招商银行北京华贸中心支行110944450110501");
		mappings.put("invoiceDetailAddress", "北京市朝阳区建国路89号华贸商务楼15号楼1503，马凌云，15004678552");
		
		mappings.put("beginningDate", "2021年11月01日");
		mappings.put("periodDateEnd", "2021-11-30");
		mappings.put("beginningPrepaidBalance", "989,321,211.10");
		mappings.put("replenishPrepaidAmount", "2131.13");
		mappings.put("prepaidBalance", "-88.88");
		try {
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
			InputStream templateInputStream = new FileInputStream("D:\\project\\pinuc\\merak\\reconciliation-business\\src\\main\\resources\\word\\结算确认函-预付-模板.docx");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);
			//wordMLPackage.setFontMapper(fontMapper);
			MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
			mainDocumentPart.variableReplace(mappings);
			new File("C:\\Users\\EDZ\\Desktop\\test2.docx").delete();
			FileOutputStream outputStream = new FileOutputStream("C:\\Users\\EDZ\\Desktop\\test2.docx");
			Docx4J.save(wordMLPackage, outputStream);
//			new File("C:\\Users\\EDZ\\Desktop\\test2.pdf").delete();
//			FileOutputStream outputStream2 = new FileOutputStream("C:\\Users\\EDZ\\Desktop\\test2.pdf");
//			Docx4J.toPDF(wordMLPackage, outputStream2);
//			FOSettings foSettings = Docx4J.createFOSettings();
//			foSettings.setOpcPackage(wordMLPackage);
			//foSettings.setApacheFopMime("application/pdf");
//			Docx4J.toFO(foSettings, outputStream2, Docx4J.FLAG_EXPORT_PREFER_XSL);
		}
		catch (Exception e) {
			//log.error(e.getMessage(), e);
		}
	}
	
}
