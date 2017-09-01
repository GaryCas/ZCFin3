package com.company;

import java.io.IOException;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.Marshaller;
import javax.xml.namespace.QName;

import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.NamespacePrefixMapperUtils;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class TextUtils {

    private static Logger log = LoggerFactory.getLogger(TextUtils.class);

    /**
     * Extract contents of descendant <w:t> elements.
     *
     * @param o
     * @return
     */
    public static void extractText(Object o, Writer w) throws Exception {

        extractText(o, w, Context.jc);
    }

    /**
     * Extract contents of descendant <w:t> elements.
     *
     * @param o
     * @param jc JAXBContext
     * @return
     */
    public static void extractText(Object o, Writer w, JAXBContext jc) throws Exception {

        Marshaller marshaller=jc.createMarshaller();
        NamespacePrefixMapperUtils.setProperty(marshaller,
                NamespacePrefixMapperUtils.getPrefixMapper());
        marshaller.marshal(o, new TextExtractor(w));

    }

    /**
     * Extract contents of descendant <w:t> elements.
     * Use this for objects which don't have @XmlRootElement
     *
     * @param o
     * @param w
     * @param jc
     * @param uri
     * @param local
     * @param declaredType
     * @throws Exception
     */
    public static void extractText(Object o, Writer w, JAXBContext jc,
                                   String uri, String local, Class declaredType) throws Exception {

        Marshaller marshaller=jc.createMarshaller();
        NamespacePrefixMapperUtils.setProperty(marshaller,
                NamespacePrefixMapperUtils.getPrefixMapper());
        marshaller.marshal(
                new JAXBElement(new QName(uri,local), declaredType, o ),
                new TextExtractor(w));
    }



    /**
     * A SAX ContentHandler that writes all #PCDATA onto a java.io.Writer
     *
     * From http://www.cafeconleche.org/books/xmljava/chapters/ch06s03.html
     *
     */
    static class TextExtractor extends DefaultHandler {

        private Writer out;

        public TextExtractor(Writer out) {
            this.out = out;
        }

        public void characters(char[] text, int start, int length)
                throws SAXException {

            try {
                out.write(text, start, length);
            }
            catch (IOException e) {
                throw new SAXException(e);
            }

        }

    } // end TextExtractor


    /**
     *
     * @param contents The main document part of the docx file in question
     * @param templatesAndReplacements A string string hashmap that contains the string templates to be replaced and
     *                                 their replacements
     *
     * Replace template strings with actual company and general information - if there is a simpler way of doing this,
     *                                 find it
     *
     * To be moved into another class - Refactor
     */

    public void replacePlaceholders(Document contents, HashMap<String, String> templatesAndReplacements) {

        replaceBodyPlaceHolders(contents, templatesAndReplacements);


        System.out.println("Done replacing template words ");
    }

    private void replaceBodyPlaceHolders(Document contents, HashMap<String, String> templatesAndReplacements) {
        for (Object p : contents.getBody().getContent()) {
            if(p instanceof P){
                for (Object r : ((P) p).getContent()) {
                    if(r instanceof R) {
                        for (Object o : ((R) r).getContent()) {
                            doReplacementOperation(o, templatesAndReplacements);
                        }
                    }
                }
            }
        }
    }

    private void doReplacementOperation(Object o, HashMap<String, String> templatesAndReplacements) {
        String placeholderString = "";

        if (o instanceof JAXBElement) {
            Object jaxbElement = ((JAXBElement) o).getValue();

            for (Map.Entry<String, String> stringStringEntry : templatesAndReplacements.entrySet()) {
                if (((JAXBElement) o).getValue() instanceof Text) {
                    placeholderString = ((Text) jaxbElement).getValue();

                    // seperate into own method for testing
                    if(placeholderString.toLowerCase().contains(stringStringEntry.getKey().toLowerCase())) {
                        placeholderString = placeholderString.replace(placeholderString, stringStringEntry.getValue());
                        ((Text) jaxbElement).setValue(placeholderString);
                        break;
                    }
                }
            }
        }

    }

}