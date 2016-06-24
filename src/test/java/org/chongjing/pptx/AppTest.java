package org.chongjing.pptx;

import org.apache.commons.io.filefilter.RegexFileFilter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Before;
import org.junit.Test;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.xpath.XPathExpressionException;
import java.io.File;
import java.io.FileFilter;
import java.io.IOException;

/**
 * Unit test for simple App.
 */
public class AppTest {

    private static final Logger logger = LogManager.getLogger(AppTest.class);

    private App app;

    @Before
    public void setup() {
        app = new App();
    }

    @Test
    public void testUnzip() throws IOException {
        //unzip("C:\\BaiduYunDownload\\01_javaweb的WEB开发入门.pptx", "C:\\BaiduYunDownload\\01\\");
    }


    @Test
    public void iterateFiles() throws ParserConfigurationException, TransformerException, SAXException, XPathExpressionException, IOException {
        //String dir = "C:\\BaiduYunDownload\\01\\ppt\\slides";
        String dir = "/Users/student/Haizong/01/ppt/slides";
        FileFilter fileFilter = new RegexFileFilter("^.*slide(\\d+)?\\.xml$");
        File file = new File(dir);
        File[] files = file.listFiles(fileFilter);
        for (int i = 0; i < files.length; i++) {
            System.out.println(files[i]);
            app.modify(files[i]);
        }
        //C:\BaiduYunDownload\01\ppt\slides\slide1.xml
    }

    /**
     * http://www.ibm.com/developerworks/cn/xml/x-nmspccontext/
     * http://my.oschina.net/cloudcoder/blog/223359
     * http://www.ibm.com/developerworks/cn/xml/x-javaxpathapi.html
     * http://www.cnblogs.com/xing901022/p/3916511.html
     *
     * @throws ParserConfigurationException
     * @throws IOException
     * @throws SAXException
     * @throws XPathExpressionException
     */
    @Test
    public void unmarshall() throws ParserConfigurationException, IOException, SAXException, XPathExpressionException, TransformerException {
        //String file = "C:\\BaiduYunDownload\\01\\ppt\\slides\\slide2.xml";
        String file = "/Users/student/Documents/课件/jsp方立勋/ppt/slides/slide5.xml";
        app.modify(file);
    }


}
