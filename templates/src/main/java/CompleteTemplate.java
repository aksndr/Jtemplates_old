import com.lowagie.text.pdf.BarcodeEAN;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.util.IOUtils;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.jvnet.jaxb2_commons.ppp.Child;

import javax.imageio.ImageIO;
import javax.xml.bind.JAXBException;
import java.awt.Color;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * User: a.arzamastsev Date: 30.07.13 Time: 9:47
 */
public class CompleteTemplate {

    private WordprocessingMLPackage wordMLPackage;
    private ObjectFactory factory;
    private Logger logger;

    public CompleteTemplate() {
    }

    public static CompleteTemplate instance() {
        return new CompleteTemplate();
    }

    public String completeTemplate(String filepath, Map<String, String> params) {
        logger = Logger.getLogger(CompleteTemplate.class);

        String pathGenerateTemplate;
        StringBuilder s = new StringBuilder();

        if (params.containsKey(BARCODE)) {
            params.put(BARCODE, getValidateBarcode(params.get(BARCODE)));
        }

        for (Map.Entry entry : params.entrySet()) {
            s.append("[").append(entry.getKey()).append(":").append(entry.getValue()).append("]");
        }
        logger.debug("Show params length: " + params.size());
        logger.debug("Show params: " + s);

        BasicConfigurator.configure();

        if (filepath == null || EMPTY.equals(filepath)) {
            logger.error("Failed to completeTemplate ", new NullPointerException("Filepath cannot be empty"));
            return EMPTY;
        }
        if (params.isEmpty()) {
            logger.error("Failed to completeTemplate ", new NullPointerException("Params is empty"));
            return EMPTY;
        }

        try {
            wordMLPackage = openDocument(filepath);
        } catch (Docx4JException e) {
            logger.error("Failed to openDocument ", e);
            return EMPTY;
        }
        factory = Context.getWmlObjectFactory();

        List<Object> list;
        try {
            list = getMergeFieldsParentsList();
        } catch (JAXBException e) {
            logger.error("Failed to getMergeFieldsParentsList ", e);
            return EMPTY;
        }
        for (Object o : list) {
            Collection<Child> finalNodes = new ArrayList<Child>();
            Object d = ((org.docx4j.wml.R) o).getParent();

            PPr ppr = ((org.docx4j.wml.P) d).getPPr();
            RFonts rprFonts = null;
            RStyle rprStyle = null;

            finalNodes.add(ppr);
            List<Object> pChildren = ((org.docx4j.wml.P) d).getContent();
            String mergeFieldName = null;
            for (Object pChild : pChildren) {
                pChild = XmlUtils.unwrap(pChild);
                if (pChild instanceof org.docx4j.wml.R) {
                    List<Object> children = ((org.docx4j.wml.R) pChild).getContent();
                    for (Object rChild : children) {
                        rChild = XmlUtils.unwrap(rChild);
                        if (rChild instanceof org.docx4j.wml.Text) {
                            String fieldName = ((Text) rChild).getValue();
                            if (isMergeField(fieldName)) {
                                mergeFieldName = getMergeFieldName(fieldName);
                            } else if (fieldName.equals(mergeFieldName)) {
                                String mergeFieldValue = params.containsKey(fieldName) ? params.get(fieldName) : EMPTY;
                                if (mergeFieldValue == null) {
                                    mergeFieldValue = EMPTY;
                                }
                                org.docx4j.wml.R run = factory.createR();
                                logger.debug("Proceeding field: " + mergeFieldName + " with value: " + mergeFieldValue);
                                if (fieldName.contains(BARCODE) && !mergeFieldValue.isEmpty()) {
                                    try {
                                        logger.debug("Before getBarcodeDrawning");
                                        run.getContent().add(getBarcodeDrawning(mergeFieldValue));
                                        logger.debug("After getBarcodeDrawning");
                                    } catch (Exception e) {
                                        logger.error("Failed to getBarcodeDrawning ", e);
                                        return EMPTY;
                                    }
                                } else {
                                    ((Text) rChild).setValue(mergeFieldValue);
                                    RPr rPr = factory.createRPr();
                                    rprFonts = ((R) pChild).getRPr().getRFonts();
                                    rprStyle = ((R) pChild).getRPr().getRStyle();
                                    rPr.setRFonts(rprFonts);
                                    run.getContent().add(rPr);
                                    run.getContent().add(rChild);
                                }
                                finalNodes.add(run);
                            }
                            System.out.println("Found Text child value" + fieldName);
                        }
                    }
                }
            }
            ((P) d).getContent().clear();
            ppr.getRPr().setRFonts(rprFonts);
            ppr.getRPr().setRStyle(rprStyle);
            ((P) d).getContent().addAll(0, finalNodes);
            logger.debug("Done.");
        }
        try {
            pathGenerateTemplate = resultFileName(filepath);
            wordMLPackage.save(new File(pathGenerateTemplate));
        } catch (Docx4JException e) {
            logger.error("Failed to save new file ", e);
            return EMPTY;
        }
        return pathGenerateTemplate;
    }

    private String getValidateBarcode(String barcode) {
        StringBuilder nBarcode = new StringBuilder().append(FORMAT);
        if (barcode != null) {
            if (!barcode.isEmpty()) {
                int startIndex = nBarcode.length() - barcode.length();
                int endIndex = nBarcode.length();

                nBarcode.replace(startIndex, endIndex, barcode);
            }
        }
        logger.debug("nBarcode." + nBarcode.toString());
        return nBarcode.toString();
    }

    private boolean isMergeField(String fieldValue) {
        return (fieldValue.contains(MERGEFIELD));
    }

    private Drawing getBarcodeDrawning(String s) throws Exception {
        byte[] bytes = getBarcodeBytes(s);
        logger.debug("getBarcodeDrawning received bytes " + bytes.length);
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
        logger.debug("BinaryPartAbstractImage created ");
        Inline inline = imagePart.createImageInline(null, null, 0, 1, 1500, false);
        logger.debug("BinaryPartAbstractImage created ");
        org.docx4j.wml.Drawing drawing = factory.createDrawing();
        logger.debug("org.docx4j.wml.Drawing created ");
        drawing.getAnchorOrInline().add(inline);
        logger.debug("drawing.getAnchorOrInline added");
        return drawing;
    }

    private byte[] getBarcodeBytes(String s) throws IOException {
        logger.debug("Started getBarcodeBytes for." + s);
        BarcodeEAN codeEAN = new BarcodeEAN();
        codeEAN.setCode(s);
        Image bCode = codeEAN.createAwtImage(new Color(0, 0, 0), new Color(255, 255, 255));
        logger.debug("bCode image created.");
        byte[] imageBytes = null;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            BufferedImage bufferedImage = toBufferedImage(bCode);
            ImageIO.write(bufferedImage, PNG, baos);
            baos.flush();
            imageBytes = baos.toByteArray();
            logger.debug("bufferedImage written.");
        } catch (Exception e) {
            logger.error("bufferedImage error. " + e);
        } finally {
            IOUtils.closeQuietly(baos);
            logger.debug("bufferedImage closed.");
        }
        return imageBytes;
    }

    private BufferedImage toBufferedImage(Image src) {
        logger.debug("toBufferedImage started.");
        int w = src.getWidth(null);
        int h = src.getHeight(null);
        int type = BufferedImage.TYPE_INT_RGB;
        BufferedImage d = new BufferedImage(w, h, type);
        logger.debug("new BufferedImage created.");
        Graphics2D g2 = d.createGraphics();
        logger.debug("createGraphics done.");
        g2.drawImage(src, 0, 0, null);
        logger.debug("drawImage done.");
//        g2.dispose();
        return d;
    }

    private List<Object> getMergeFieldsParentsList() throws JAXBException {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        String xpath = "//w:r[w:instrText]";
//        String xpath = "//w:p[w:r/w:instrText]";
        return documentPart.getJAXBNodesViaXPath(xpath, true);
    }

    private String getMergeFieldName(String text) {
        text = text.replaceAll("MERGEFIELD", "");
        text = text.substring(0, text.indexOf("\\"));
        return text.trim();
    }

    private WordprocessingMLPackage openDocument(String inputfilepath) throws Docx4JException {
        return WordprocessingMLPackage.load(new java.io.File(inputfilepath));
    }

    private String resultFileName(String str) {
        int dot = str.indexOf(DOT);
        String ext = str.substring(dot);
        return str.substring(0, dot) + System.currentTimeMillis() + ext;
    }


    private final static String EMPTY = "";
    private final static String FORMAT = "0000000000000";
    private final static String BARCODE = "barcode";
    private final static String MERGEFIELD = "MERGEFIELD";
    private final static String PNG = "png";
    private final static String DOT = ".";
}