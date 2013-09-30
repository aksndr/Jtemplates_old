import com.lowagie.text.rtf.document.RtfDocument;
import net.sourceforge.rtf.IRTFDocumentTransformer;
import net.sourceforge.rtf.RTFTemplate;
import net.sourceforge.rtf.handler.RTFDocumentHandler;
import net.sourceforge.rtf.template.velocity.RTFVelocityTransformerImpl;
import net.sourceforge.rtf.template.velocity.VelocityTemplateEngineImpl;
import org.apache.velocity.app.VelocityEngine;

import java.io.*;
import java.text.FieldPosition;
import java.text.Format;
import java.text.ParsePosition;
import java.util.HashMap;
import java.util.Map;

//import com.itextpdf.text.Image;

public class GeneratorUtils {

	public static String getRtfFromTemplate(String templatePath, Map params) {
		InputStream templateIS;
		try {
			templateIS = new FileInputStream(new File(templatePath));
		} catch (FileNotFoundException e) {
			return "Cannot find template";
		}
		RTFTemplate rtfTemplate = new RTFTemplate();
		// Parser
		RTFDocumentHandler parser = new RTFDocumentHandler();
		rtfTemplate.setParser(parser);

		// Transformer
		IRTFDocumentTransformer transformer = new RTFVelocityTransformerImpl();
		rtfTemplate.setTransformer(transformer);

		// Template engine
		VelocityTemplateEngineImpl templateEngine = new VelocityTemplateEngineImpl();
		// Initialize velocity engine
		VelocityEngine velocityEngine = new VelocityEngine();
//		velocityEngine.setProperty("input.encoding ", "UTF-8"); 
//        velocityEngine.setProperty("output.encoding", "UTF-8"); 
//        velocityEngine.setProperty ("response.encoding", "UTF-8");
		templateEngine.setVelocityEngine(velocityEngine);
		rtfTemplate.setTemplateEngine(templateEngine);

		// Set the RTF model source
		rtfTemplate.setTemplate(templateIS);

		rtfTemplate.put("beans", params);
		rtfTemplate.put("skode", params);

		rtfTemplate.setDefaultFormat(java.lang.String.class, new Format() {
			public StringBuffer format(Object obj, StringBuffer toAppendTo, FieldPosition pos) {
//                 String str = "" + obj;
//                 str = escape((str));
                String str = escape((String)obj);
				toAppendTo.append(str.replaceAll("\n", "\\\\line "));
				return toAppendTo;
			}

			public Object parseObject(String source, ParsePosition pos) {
				return null;
			}
		});
		
		File file;
		try {
			file = File.createTempFile("rep_", ".doc");
		} catch (IOException e1) {
			return "Cannot create temp file";
		}
		try {
			rtfTemplate.merge(file);
		} catch (Exception e) {
			return "Cannot merge template";
		}

		return file.getAbsolutePath();
	}

	public static void main(String[] args) throws Exception {
        HashMap<String, Object> params = new HashMap<String, Object>();
		params.put("type", "test");


		params.put("skode", "barcode_here");

        String outPath = getRtfFromTemplate("Регистрационная карточка.doc",params);
		System.out.println(outPath);
	}
	
	public static String escape(String sentence) {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		try {
			new RtfDocument().filterSpecialChar(baos, sentence, true, true);
		} catch (IOException e) {
			// will never happen for ByteArrayOutputStream
		}
		return new String(baos.toByteArray());
	}

}
