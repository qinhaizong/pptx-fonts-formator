package org.chongjing.pptx;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.RegexFileFilter;
import org.junit.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.*;
import java.io.File;
import java.io.FileFilter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Unit test for simple App.
 */
public class AppTest {

    @Test
    public void testUnzip() throws IOException {
        unzip("C:\\BaiduYunDownload\\01_javaweb的WEB开发入门.pptx", "C:\\BaiduYunDownload\\01\\");
    }

    /**
     * @param source
     * @param target
     * @throws IOException
     */
    public void unzip(String source, String target) throws IOException {
        //TODO validate parameters
        ZipFile zipFile = new ZipFile(source);
        File targetDir = new File(target);
        if (!targetDir.exists()) {
            targetDir.mkdirs();
        }
        String targetDirAbsolutePath = targetDir.getAbsolutePath();
        Enumeration<? extends ZipEntry> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = entries.nextElement();
            String name = entry.getName();
            File file = new File(targetDirAbsolutePath + File.separator + name);
            if (entry.isDirectory()) {
                file.mkdirs();
            } else {
                //file.createNewFile();
                FileUtils.touch(file);
                IOUtils.copy(zipFile.getInputStream(entry), new FileOutputStream(file));
            }
            System.out.println(file.getAbsoluteFile());
        }
    }

    @Test
    public void iterateFiles() {
        String dir = "C:\\BaiduYunDownload\\01\\ppt\\slides";
        FileFilter fileFilter = new RegexFileFilter("^.*slide(\\d+)?\\.xml$");
        File file = new File(dir);
        File[] files = file.listFiles(fileFilter);
        for (int i = 0; i < files.length; i++) {
            System.out.println(files[i]);
        }
        //C:\BaiduYunDownload\01\ppt\slides\slide1.xml
    }

    /**
     * http://www.ibm.com/developerworks/cn/xml/x-nmspccontext/
     * http://my.oschina.net/cloudcoder/blog/223359
     * http://www.ibm.com/developerworks/cn/xml/x-javaxpathapi.html
     * http://www.cnblogs.com/xing901022/p/3916511.html
     * @throws ParserConfigurationException
     * @throws IOException
     * @throws SAXException
     * @throws XPathExpressionException
     */
    @Test
    public void unmarshall() throws ParserConfigurationException, IOException, SAXException, XPathExpressionException {
        String file = "C:\\BaiduYunDownload\\01\\ppt\\slides\\slide2.xml";
        DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
        builderFactory.setNamespaceAware(true);
        DocumentBuilder documentBuilder = builderFactory.newDocumentBuilder();
        Document document = documentBuilder.parse(new File(file));
        XPathFactory factory = XPathFactory.newInstance();
        XPath xPath = factory.newXPath();
        xPath.setNamespaceContext(new NamespaceContext() {
            @Override
            public String getNamespaceURI(String prefix) {
                if (prefix == null) {
                    throw new IllegalArgumentException("No prefix provided!");
                    //} else if (prefix.equals(XMLConstants.DEFAULT_NS_PREFIX)) {
                } else if ("a".equals(prefix)) {
                    return "http://schemas.openxmlformats.org/drawingml/2006/main";
                } else if ("r".equals(prefix)) {
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                } else if ("p".equals(prefix)) {
                    return "http://schemas.openxmlformats.org/presentationml/2006/main";
                }
                return null;
            }

            @Override
            public String getPrefix(String namespaceURI) {
                return null;
            }

            @Override
            public Iterator getPrefixes(String namespaceURI) {
                return null;
            }
        });
        XPathExpression spCompile = xPath.compile("//p:sp");
        XPathExpression eaCompile = xPath.compile("./p:txBody/a:p/a:r/a:rPr/a:ea");
        NodeList sps = (NodeList) spCompile.evaluate(document, XPathConstants.NODESET);
        for (int i = 0; i < sps.getLength(); i++) {
            Node node = sps.item(i);
            System.out.println(node.getNodeName());
            NodeList eas = (NodeList) eaCompile.evaluate(node, XPathConstants.NODESET);
            for (int j = 0; j < eas.getLength(); j++) {
                Node item = eas.item(j);
                System.out.println(i + "\t" + item.getNodeName() + "\t" + item.getAttributes().getNamedItem("typeface").getNodeValue());
            }
        }

    }
}
