package org.chongjing.pptx;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.RegexFileFilter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.junit.Test;
import org.w3c.dom.*;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
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

    private static final Logger logger = LogManager.getLogger(AppTest.class);

    @Test
    public void testUnzip() throws IOException {
        //unzip("C:\\BaiduYunDownload\\01_javaweb的WEB开发入门.pptx", "C:\\BaiduYunDownload\\01\\");
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
    public void iterateFiles() throws ParserConfigurationException, TransformerException, SAXException, XPathExpressionException, IOException {
        //String dir = "C:\\BaiduYunDownload\\01\\ppt\\slides";
        String dir = "/Users/student/Haizong/01/ppt/slides";
        FileFilter fileFilter = new RegexFileFilter("^.*slide(\\d+)?\\.xml$");
        File file = new File(dir);
        File[] files = file.listFiles(fileFilter);
        for (int i = 0; i < files.length; i++) {
            System.out.println(files[i]);
            modify(files[i]);
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
        String file = "/Users/student/Haizong/01/ppt/slides/slide2.xml";
        modify(file);

    }

    private void modify(String file) throws IOException, ParserConfigurationException, SAXException, XPathExpressionException, TransformerException {
        File sourceXml = new File(file);
        modify(sourceXml);
    }

    private void modify(File sourceXml) throws ParserConfigurationException, IOException, SAXException, XPathExpressionException, TransformerException {
        DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
        builderFactory.setNamespaceAware(true);
        DocumentBuilder documentBuilder = builderFactory.newDocumentBuilder();
        Document document = documentBuilder.parse(sourceXml);
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
        /*NodeList ns = (NodeList) xPath.evaluate("//p:sp/p:txBody/a:p/a:r/a:rPr/a:ea", document, XPathConstants.NODESET);
        for (int i = 0; i < ns.getLength(); i++) {
            ns.item(i).getAttributes().getNamedItem("typeface").setNodeValue("Microsoft YaHei");
            System.out.println(ns.item(i).getNodeName());
        }*/
        XPathExpression spCompile = xPath.compile("//p:sp");
        XPathExpression txBodyCompile = xPath.compile("./p:txBody");
        XPathExpression spPrCompile = xPath.compile("./p:spPr");
        XPathExpression offCompile = xPath.compile("./a:xfrm/a:off");
        XPathExpression extCompile = xPath.compile("./a:xfrm/a:ext");
        XPathExpression rprCompile = xPath.compile("./p:txBody/a:p/a:r/a:rPr");
        XPathExpression eaCompile = xPath.compile("./p:txBody/a:p/a:r/a:rPr/a:ea");
        NodeList sps = (NodeList) spCompile.evaluate(document, XPathConstants.NODESET);
        for (int i = 0; i < sps.getLength(); i++) {
            Node node = sps.item(i);
            System.out.println(node.getNodeName());
            if (i == 0) {
                //添加偏移
                Element n = (Element) spPrCompile.evaluate(node, XPathConstants.NODE);
                logger.info("{}", n);
                if (n != null && n.hasChildNodes()) {
                    logger.info("编辑子节点情况");
                    //<p:spPr> <a:xfrm> <a:off x="539750" y="44450"/> <a:ext cx="8064500" cy="936625"/> </a:xfrm> <a:noFill/> </p:spPr>
                    Node off = (Node) offCompile.evaluate(n, XPathConstants.NODE);
                    if (null != off) {
                        NamedNodeMap offAttributes = off.getAttributes();
                        offAttributes.getNamedItem("x").setNodeValue("539750");
                        offAttributes.getNamedItem("y").setNodeValue("44450");
                    }
                    Node ext = (Node) extCompile.evaluate(n, XPathConstants.NODE);
                    if (null != ext) {
                        NamedNodeMap extAttributes = ext.getAttributes();
                        extAttributes.getNamedItem("cx").setNodeValue("8064500");
                        extAttributes.getNamedItem("cy").setNodeValue("936625");
                    }

                } else if (!n.hasChildNodes()) { // <p:spPr/>
                    logger.info("添加子节点");
                    Element xfrm = document.createElement("a:xfrm");
                    Element aOff = document.createElement("a:off");
                    aOff.setAttribute("x", "539750");
                    aOff.setAttribute("y", "44450");
                    Element aExt = document.createElement("a:ext");
                    aExt.setAttribute("cx", "8064500");
                    aExt.setAttribute("cy", "936625");
                    xfrm.appendChild(aOff);
                    xfrm.appendChild(aExt);
                    n.appendChild(xfrm);
                } else { //
                    logger.info("创建全新的了节点");
                    Element spPr = document.createElement("p:spPr");
                    Element xfrm = document.createElement("a:xfrm");
                    Element aOff = document.createElement("a:off");
                    aOff.setAttribute("x", "539750");
                    aOff.setAttribute("y", "44450");
                    Element aExt = document.createElement("a:ext");
                    aExt.setAttribute("cx", "8064500");
                    aExt.setAttribute("cy", "936625");
                    xfrm.appendChild(aOff);
                    xfrm.appendChild(aExt);
                    spPr.appendChild(xfrm);
                    Node txBody = (Node) txBodyCompile.evaluate(node, XPathConstants.NODE);
                    document.insertBefore(spPr, txBody);
                }
                // 修改标头斜体, 大小
                NodeList rprs = (NodeList) rprCompile.evaluate(node, XPathConstants.NODESET);
                for (int j = 0; j < rprs.getLength(); j++) {
                    Node rpr = rprs.item(j);
                    NamedNodeMap attributes = rpr.getAttributes();
                    Node sz = attributes.getNamedItem("sz");
                    if (null != sz) {
                        sz.setNodeValue("2800");
                    }
                    System.out.println(rpr.getNodeName() + "\t" + rpr.getAttributes().getNamedItem("sz"));
                    try {
                        rpr.getAttributes().removeNamedItem("i");
                    } catch (DOMException e) {
                        System.err.println(e.getMessage());
                    }
                }
            }
            // 修改所有字体
            NodeList eas = (NodeList) eaCompile.evaluate(node, XPathConstants.NODESET);
            for (int j = 0; j < eas.getLength(); j++) {
                Node item = eas.item(j); //Microsoft YaHei
                item.getAttributes().getNamedItem("typeface").setNodeValue("Microsoft YaHei");
                System.out.println(i + "\t" + item.getNodeName() + "\t" + item.getAttributes().getNamedItem("typeface").getNodeValue());
            }
        }
        //File targetXml = new File(dir + "xml/slide2_.xml");
        //FileUtils.touch(targetXml);
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.transform(new DOMSource(document), new StreamResult(sourceXml));
    }
}
