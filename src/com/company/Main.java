package com.company;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.*;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;

public class Main {
    static ObjectFactory factory = Context.getWmlObjectFactory();


    public static void main(String[] args) throws Docx4JException {
        TextUtils textUtils = new TextUtils();


        HashMap templateReplacements = new HashMap<String, String>();
        templateReplacements.put("[Small company limited]", "zcFin");

        WordprocessingMLPackage template = WordprocessingMLPackage.load(new File((System.getProperty("user.dir"))+ "/NonGoogleDocTemp.docx"));
        MainDocumentPart mainDocumentPart = template.getMainDocumentPart();

        try {
            inspectTemplate(template);
            textUtils.replacePlaceholders(mainDocumentPart.getContents(), templateReplacements);
        } catch (JAXBException e) {
            e.printStackTrace();
        }

        template.save(new File((System.getProperty("user.dir") + "/GeneratedDoc.docx")));

    }


    protected static void inspectTemplate(WordprocessingMLPackage template) throws JAXBException {

        RelationshipsPart relationships = template.getRelationshipsPart();

        for (Map.Entry<PartName, Part> partNamePartEntry : template.getParts().getParts().entrySet()) {
            partNamePartEntry.getValue();
        }
    }



    private static void inspectParts(Document jaxbElement) throws JAXBException {


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

}