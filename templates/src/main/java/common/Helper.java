package common;

import org.docx4j.XmlUtils;
import org.jvnet.jaxb2_commons.ppp.Child;

import javax.xml.bind.JAXBElement;

/**
 * User: a.arzamastsev Date: 30.07.13 Time: 14:00
 */
public class Helper {

    public Helper() {

    }

    public static Helper instance (){
        return new Helper();
    }


    public org.docx4j.wml.P getParentNode(Child t) {
        JAXBElement para = (JAXBElement)t.getParent();
        org.docx4j.wml.P p = null;
        if ( para.getDeclaredType().getName().equals("org.docx4j.wml.P") ) {
            p = (org.docx4j.wml.P)para.getValue();
        } else {
            p = Helper.instance().getParentNode((Child)XmlUtils.unwrap(para));
        }
        return p;
    }
}
