import java.util.HashMap;

import javax.xml.bind.JAXBContext;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Document;

public class UnmarshallFromTemplate {
    public static JAXBContext context = org.docx4j.jaxb.Context.jc;

    public static void main(String[] args) throws Exception {

        String inputfilepath = "unmarshallFromTemplateExample.docx";

        boolean save = true;
        String outputfilepath = "test-out.docx";


        // Open a document from the file system
        // 1. Load the Package
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File(inputfilepath));

        // 2. Fetch the document part
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();

        //xml --> string
        String xml = XmlUtils.marshaltoString(wmlDocumentEl, true);

        HashMap<String, String> mappings = new HashMap<String, String>();

        mappings.put("colour", "green");
        mappings.put("icecream", "chocolate");

        //valorize template
        Object obj = XmlUtils.unmarshallFromTemplate(xml, mappings);

        //change  JaxbElement
        documentPart.setJaxbElement((Document) obj);

        // Save it
        if (save) {
            SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
            saver.save(outputfilepath);
            System.out.println( "Saved output to:" + outputfilepath );
        } else {
            // Display the Main Document Part.

        }
    }
}
