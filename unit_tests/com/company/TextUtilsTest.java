package com.company;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;
import org.junit.Test;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

import static junit.framework.TestCase.assertTrue;
import static org.assertj.core.api.AssertionsForInterfaceTypes.assertThat;


/**
 * Created by Desktop on 8/17/2017.
 */
public class TextUtilsTest {

    TextUtils textUtils = new TextUtils();

    @Test
    public void shouldReplaceTemplateText() throws Docx4JException {
        //given
        HashMap<String, String> templateReplacements = new HashMap<String, String>();
        templateReplacements.put("[Small company limited]", "zcFin");
        templateReplacements.put("[year]", "2017");
        templateReplacements.put("[31 December 2016]","31 December 2017");
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new File(System.getProperty("user.dir") + "/testResources/NonGoogleDocTemp.docx"));

        //when
        textUtils.replacePlaceholders(template.getMainDocumentPart().getContents(), templateReplacements);

        //then
        String testStr = checkBodyStr(template.getMainDocumentPart().getContents());

        for (Map.Entry<String, String> stringStringEntry : templateReplacements.entrySet()) {
            assertThat(testStr).doesNotContain(stringStringEntry.getKey());
            assertThat(testStr).contains(stringStringEntry.getValue());
        }


    }


    @Test
    public void should(){
        //given
        //when
        //then
    }


    public String checkBodyStr(Document contents) {
        String arrestionString = "";

        for (Object p : contents.getBody().getContent()) {
            if(p instanceof P){
                for (Object r : ((P) p).getContent()) {
                    if(r instanceof R) {
                        for (Object o : ((R) r).getContent()) {
                            if (o instanceof JAXBElement) {
                                Object jaxbElement = ((JAXBElement) o).getValue();
                                if (((JAXBElement) o).getValue() instanceof Text) {
                                    String string = ((Text) jaxbElement).getValue();
                                    arrestionString = arrestionString.concat(string + " ");
                                }
                            }
                        }
                    }
                }
            }
        }

        return arrestionString;
    }

}
