package com.github.gengxiaoxiaoxin.toolkit.office;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.InputStream;

class WordUtilsTest {
	
	@Test
	void appendContent() throws Docx4JException {
		File file = new File("test-output/appendContentResult.doc");
		InputStream destInputStream = this.getClass().getResourceAsStream("/word/dest.docx");
		InputStream sourceInputStream = this.getClass().getResourceAsStream("/word/source.docx");
		WordprocessingMLPackage dest = WordprocessingMLPackage.load(destInputStream);
		WordprocessingMLPackage source = WordprocessingMLPackage.load(sourceInputStream);
		WordUtils.appendContent(dest, source);
		Docx4J.save(dest, file);
	}
}
