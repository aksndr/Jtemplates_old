/**
 * User: a.arzamastsev Date: 29.07.13 Time: 16:18
 */

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;

import com.lowagie.text.pdf.BarcodeEAN;
import org.apache.poi.util.IOUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;

import javax.imageio.ImageIO;


public class ImageAdd {

    public static void main(String[] args) throws Exception {
        addBcodeImage("4512345678906");
    }

    private static void addBcodeImage (String code)throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

        Image bcode = getBcodeImageForString(code);

        final ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            final BufferedImage bufferedImage = toBufferedImage(bcode);
            ImageIO.write(bufferedImage, "png", baos);
        } finally {
            IOUtils.closeQuietly(baos);
        }
        byte[] bytes  = baos.toByteArray();

        String filenameHint = null;
        String altText = null;
        int id1 = 0;
        int id2 = 1;


        // Image 1: no width specified
        org.docx4j.wml.P p = newImage( wordMLPackage, bytes,
                filenameHint, altText,
                id1, id2 );
        wordMLPackage.getMainDocumentPart().addObject(p);

        // Image 2: width 3000
        org.docx4j.wml.P p2 = newImage( wordMLPackage, bytes,
                filenameHint, altText,
                id1, id2, 3000 );
        wordMLPackage.getMainDocumentPart().addObject(p2);

        // Image 3: width 6000
        org.docx4j.wml.P p3 = newImage( wordMLPackage, bytes,
                filenameHint, altText,
                id1, id2, 6000 );
        wordMLPackage.getMainDocumentPart().addObject(p3);

        wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/OUT_AddImage.docx") );
    }

    private static Image getBcodeImageForString(String code) {
        BarcodeEAN codeEAN = new BarcodeEAN();
        codeEAN.setCode(code);
        return codeEAN.createAwtImage(Color.white, Color.black);
    }

    private static BufferedImage toBufferedImage(Image src) {
        int w = src.getWidth(null);
        int h = src.getHeight(null);
        int type = BufferedImage.TYPE_INT_RGB;  // other options
        BufferedImage dest = new BufferedImage(w, h, type);
        Graphics2D g2 = dest.createGraphics();
        g2.drawImage(src, 0, 0, null);
        g2.dispose();
        return dest;
    }


    /**
     * Create image, without specifying width
     */
    public static org.docx4j.wml.P newImage( WordprocessingMLPackage wordMLPackage,
                                             byte[] bytes,
                                             String filenameHint, String altText,
                                             int id1, int id2) throws Exception {

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline( filenameHint, altText,
                id1, id2, false);

        // Now add the inline in w:p/w:r/w:drawing
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;
    }

    /**
     * Create image, specifying width in twips
     */
    public static org.docx4j.wml.P newImage( WordprocessingMLPackage wordMLPackage,
                                             byte[] bytes,
                                             String filenameHint, String altText,
                                             int id1, int id2, long cx) throws Exception {

        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);

        Inline inline = imagePart.createImageInline( filenameHint, altText,
                id1, id2, cx, false);

        // Now add the inline in w:p/w:r/w:drawing
        org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        org.docx4j.wml.P p = factory.createP();
        org.docx4j.wml.R run = factory.createR();
        p.getContent().add(run);
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;

    }
}
