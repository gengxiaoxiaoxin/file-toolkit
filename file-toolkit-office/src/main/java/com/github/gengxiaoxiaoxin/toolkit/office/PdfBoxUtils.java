package com.github.gengxiaoxiaoxin.toolkit.office;

import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;

/**
 * 基于PdfBox的工具类
 */
public class PdfBoxUtils {
	
	
	/**
	 * write text into pdf
	 *
	 * @param pdPageContentStream pdPageContentStream
	 * @param text                write content
	 * @param x                   offsetX
	 * @param y                   offsetY
	 * @param isBold              is bold?
	 * @throws Exception
	 */
	public static void writeText(PDPageContentStream pdPageContentStream, String text, float x, float y, boolean isBold) throws Exception {
		// Begin the Content stream
		pdPageContentStream.beginText();
		// Setting the position for the line
		pdPageContentStream.newLineAtOffset(x, y);
		if (isBold) {
			pdPageContentStream.setRenderingMode(RenderingMode.FILL_STROKE);
		}
		// Adding text in the form of string
		pdPageContentStream.showText(text);
		// Ending the content stream
		pdPageContentStream.endText();
	}
	
	public static void writeTextByCm(PDPageContentStream pdPageContentStream, String text, float x, float y, boolean isBold) throws Exception {
		writeText(pdPageContentStream, text, cm2Pixel(x), cm2Pixel(y), isBold);
	}
	
	public static float cm2Pixel(float cm) {
		return cm / 2.54f * 72;
	}
	
}
