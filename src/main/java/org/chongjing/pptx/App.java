package org.chongjing.pptx;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.io.filefilter.DirectoryFileFilter;
import org.apache.commons.io.filefilter.FileFileFilter;
import org.apache.commons.io.filefilter.RegexFileFilter;
import org.apache.commons.io.filefilter.SuffixFileFilter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.w3c.dom.*;
import org.xml.sax.SAXException;

import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * 格式化pptx文件
 *
 * @author qinhaizong
 */
public class App {
    private static final Logger logger = LogManager.getLogger(App.class);

    private List<String> deleteText = Arrays.asList("北京传智播客教育 www.itcast.cn");

    public static void main(String[] args) throws IOException, ParserConfigurationException, SAXException, XPathExpressionException, TransformerException {
        String pptxRootDir = "/Users/student/Documents/课件/";
        App app = new App();
        app.handlePptx(pptxRootDir);
    }

    /**
     * @param pptxRootDir 有pptx文件的文件夹
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws XPathExpressionException
     * @throws TransformerException
     */
    private void handlePptx(String pptxRootDir) throws IOException, ParserConfigurationException, SAXException, XPathExpressionException, TransformerException {
        File file = new File(pptxRootDir);
        if (!file.exists()) {
            logger.error("{}路径不存在", pptxRootDir);
            System.exit(-1);
        }
        Iterator<File> pptxs = FileUtils.iterateFilesAndDirs(file, new SuffixFileFilter("pptx"), DirectoryFileFilter.INSTANCE);
        while (pptxs.hasNext()) {
            File pptx = pptxs.next();
            if (pptx.isDirectory()) {
                continue;
            }
            logger.info(">>>\t{}", pptx.getAbsoluteFile());
            String base = FilenameUtils.getFullPath(pptx.getAbsolutePath()) + FilenameUtils.getBaseName(pptx.getAbsolutePath());
            //解压文件
            unzip(pptx.getAbsolutePath(), base);
            //修改文件
            handleSlides(base);
            //压缩文件
            zip(base, new SimpleDateFormat("_yyyyMMddHHmmss").format(new Date()));
        }
        /*
        FileFilter fileFilter = new RegexFileFilter("^.*\\.pptx$");
        File[] files = file.listFiles(fileFilter);
        for (int i = 0; i < files.length; i++) {
            File pptx = files[i];
            logger.info("{}", pptx.getAbsolutePath());
            String base = FilenameUtils.getFullPath(pptx.getAbsolutePath()) + FilenameUtils.getBaseName(pptx.getAbsolutePath());
            //解压文件
            unzip(pptx.getAbsolutePath(), base);
            //修改文件
            handleSlides(base);
            //压缩文件
            zip(base, new SimpleDateFormat("_yyyyMMddHHmmss").format(new Date()));
        }
        */
    }

    /**
     * 压缩文件
     *
     * @param base 解压出来的目录
     * @param ext  为避免与原文件重复,后边加入扩展。如:xx.pptx -> xx_年月日时分秒.pptx
     * @throws IOException
     */
    private void zip(String base, String ext) throws IOException {
        logger.info("压缩文件:{} --> {}{}.pptx", base, base, ext);
        File file = new File(base);
        /*File[] files = file.listFiles();
        for (int i = 0; i < files.length; i++) {
            logger.info("{}",files[i]);
        }*/
        File zipFile = new File(base + ext + ".pptx");
        ZipOutputStream zos = null;
        try {
            zos = new ZipOutputStream(new FileOutputStream(zipFile));
            Iterator<File> iterator = FileUtils.iterateFilesAndDirs(file, FileFileFilter.FILE, DirectoryFileFilter.DIRECTORY);
            while (iterator.hasNext()) {
                File next = iterator.next();
                if (next.isDirectory()) {
                    logger.info("跳过目录");
                    continue;
                }
                String name = next.getAbsolutePath().substring(base.length() + 1);
                logger.info("添加文件:{}", name);
                zos.putNextEntry(new ZipEntry(name));
                FileInputStream fis = null;
                try {
                    fis = new FileInputStream(next);
                    int temp = 0;
                    while ((temp = fis.read()) != -1) {
                        zos.write(temp);
                    }
                    fis.close();
                } finally {
                    IOUtils.closeQuietly(fis);
                }
            }
            zos.close();
        } finally {
            IOUtils.closeQuietly(zos);
        }
    }


    /**
     * 解压pptx文件
     *
     * @param source pptx文件
     * @param target 解压目录
     * @throws IOException
     */
    public void unzip(String source, String target) throws IOException {
        logger.info("source:{}\ttarget:{}", source, target);
        ZipFile zipFile = new ZipFile(source);
        File targetDir = new File(target);
        if (!targetDir.exists()) {
            targetDir.mkdirs();
        } else {
            FileUtils.deleteQuietly(targetDir);
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
            //logger.info(file.getAbsoluteFile());
        }
    }

    /**
     * 修改pptx页面
     *
     * @param basepath /ppt/slides/slide[].xml
     * @return
     */
    public void handleSlides(String basepath) throws ParserConfigurationException, TransformerException, SAXException, XPathExpressionException, IOException {
        String dir = basepath + "/ppt/slides/";
        logger.info("修改pptx slide 目录下文件:{}", dir);
        FileFilter fileFilter = new RegexFileFilter("^.*slide(\\d+)?\\.xml$");
        File file = new File(dir);
        File[] files = file.listFiles(fileFilter);
        for (int i = 0; i < files.length; i++) {
            logger.info(files[i]);
            modify(files[i]);
        }
    }

    /**
     * 修改xml文件
     *
     * @param file 文件名
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws XPathExpressionException
     * @throws TransformerException
     */
    public void modify(String file) throws IOException, ParserConfigurationException, SAXException, XPathExpressionException, TransformerException {
        File sourceXml = new File(file);
        modify(sourceXml);
    }

    /**
     * ppt schemas
     */
    class PptxNamespace implements NamespaceContext {
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
    }

    /**
     * 修改xml文件
     *
     * @param sourceXml
     * @throws ParserConfigurationException
     * @throws IOException
     * @throws SAXException
     * @throws XPathExpressionException
     * @throws TransformerException
     */
    public void modify(File sourceXml) throws ParserConfigurationException, IOException, SAXException, XPathExpressionException, TransformerException {
        DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
        builderFactory.setNamespaceAware(true);
        DocumentBuilder documentBuilder = builderFactory.newDocumentBuilder();
        Document document = documentBuilder.parse(sourceXml);
        XPathFactory factory = XPathFactory.newInstance();
        XPath xPath = factory.newXPath();
        xPath.setNamespaceContext(new PptxNamespace());
        XPathExpression spCompile = xPath.compile("//p:sp");
        XPathExpression eaCompile = xPath.compile("./p:txBody/a:p/a:r/a:rPr/a:ea");
        XPathExpression tCompile = xPath.compile("./p:txBody/a:p/a:r/a:t");
        NodeList sps = (NodeList) spCompile.evaluate(document, XPathConstants.NODESET);
        handlerHeaderTitle(document, xPath, sps.item(0));
        for (int i = 0; i < sps.getLength(); i++) {
            Node sp = sps.item(i);
            // 处理要删除的指定段落内容
            Node t = (Node) tCompile.evaluate(sp, XPathConstants.NODE);
            if (null != t) {
                String content = t.getTextContent();
                if (deleteText.contains(content)) {
                    logger.info("删除节点:{}", content);
                    sp.getParentNode().removeChild(sp);
                    continue;
                }
            }
            // 修改所有字体
            NodeList eas = (NodeList) eaCompile.evaluate(sp, XPathConstants.NODESET);
            for (int j = 0; j < eas.getLength(); j++) {
                Node item = eas.item(j); //Microsoft YaHei
                item.getAttributes().getNamedItem("typeface").setNodeValue("Microsoft YaHei");
                logger.info(i + "\t" + item.getNodeName() + "\t" + item.getAttributes().getNamedItem("typeface").getNodeValue());
            }
            if (i == 0) {
                continue;
            }
            handleContent(document, xPath, sp);

        }
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.transform(new DOMSource(document), new StreamResult(sourceXml));
    }

    /**
     * 修改pptx内容
     * //<a:pPr eaLnBrk="1" hangingPunct="1" lvl="1"> first
     * //<a:r><a:rPr altLang="en-US" dirty="0" lang="zh-CN" sz="1900">
     *
     * @param document
     * @param xPath
     * @param node
     */
    private void handleContent(Document document, XPath xPath, Node node) throws XPathExpressionException {
        XPathExpression pCompile = xPath.compile("./p:txBody/a:p");//<a:p>

        NodeList ps = (NodeList) pCompile.evaluate(node, XPathConstants.NODESET);
        int length = ps.getLength();
        if (length == 0) {
            logger.warn("无段落");
            return;
        }
        for (int i = 0; i < length; i++) {//a:rPr
            handleSub(document, xPath, ps.item(i));
        }
    }

    /**
     * 修改pptx内容子段落
     *
     * @param document
     * @param xPath
     * @param p
     * @throws XPathExpressionException
     */
    private void handleSub(Document document, XPath xPath, Node p) throws XPathExpressionException {
        XPathExpression rPrCompile = xPath.compile("./a:r/a:rPr");
        if (null == p && !p.hasChildNodes()) {
            logger.warn("段落为空或无子节点");
            return;
        }
        NodeList rprs = (NodeList) rPrCompile.evaluate(p, XPathConstants.NODESET);
        if (null == rprs && rprs.getLength() == 0) {
            logger.warn("无内容");
            return;
        }
        Node pPr = p.getFirstChild(); //<a:pPr eaLnBrk="1" hangingPunct="1" lvl="1">
        Node lvl = pPr.getAttributes().getNamedItem("lvl");
        boolean isLvl = (null != lvl && lvl.getNodeValue().equals("1"));
        int rprsLength = rprs.getLength();
        for (int j = 0; j < rprsLength; j++) {//a:rPr
            setupSz(document, rprs.item(j), isLvl);
        }
    }

    /**
     * 设置pptx内容段落字体和大小
     *
     * @param document
     * @param rPr
     * @param isLvl
     */
    private void setupSz(Document document, Node rPr, boolean isLvl) {
        if (null == rPr) {
            logger.warn("rPr is Null");
            return;
        }
        if (!rPr.hasChildNodes()) { //<a:latin charset="0" typeface="Microsoft YaHei"/><a:ea charset="0" typeface="Microsoft YaHei"/>
            Element latin = document.createElement("a:latin");
            latin.setAttribute("charset", "0");
            latin.setAttribute("typeface", "Microsoft YaHei");
            Element ea = document.createElement("a:ea");
            ea.setAttribute("charset", "0");
            ea.setAttribute("typeface", "Microsoft YaHei");
            rPr.appendChild(latin);
            rPr.appendChild(ea);
        }
        NamedNodeMap attributes = rPr.getAttributes();
        if (null == attributes || attributes.getLength() == 0) {
            logger.warn("attributes is Null");
            return;
        }
        Node sz = attributes.getNamedItem("sz");
        if (null == sz) {
            logger.info("sz为空,设置字体大小");
            Element e = (Element) rPr;
            e.setAttribute("sz", isLvl ? "2000" : "2400");
            return;
        }
        if (isLvl) {
            logger.info("设置子段落字体大小为2000");
            sz.setNodeValue("2000");
        } else {
            logger.info("设置段落字体大小为2400");
            sz.setNodeValue("2400");
        }
    }

    /**
     * 处理pptx标题
     *
     * @param document
     * @param xPath
     * @param node
     * @throws XPathExpressionException
     */
    private void handlerHeaderTitle(Document document, XPath xPath, Node node) throws XPathExpressionException {
        XPathExpression txBodyCompile = xPath.compile("./p:txBody");
        XPathExpression spPrCompile = xPath.compile("./p:spPr");
        XPathExpression offCompile = xPath.compile("./a:xfrm/a:off");
        XPathExpression extCompile = xPath.compile("./a:xfrm/a:ext");
        XPathExpression rprCompile = xPath.compile("./p:txBody/a:p/a:r/a:rPr");
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
            n.appendChild(createXfrm(document));
        } else { //
            logger.info("创建全新的了节点");
            Element spPr = document.createElement("p:spPr");
            spPr.appendChild(createXfrm(document));
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
            logger.info(rpr.getNodeName() + "\t" + rpr.getAttributes().getNamedItem("sz"));
            try {
                rpr.getAttributes().removeNamedItem("i");
            } catch (DOMException e) {
                System.err.println(e.getMessage());
            }
        }
    }

    /**
     * 标题头偏移
     *
     * @param document
     * @return
     */
    public Element createXfrm(Document document) {
        Element xfrm = document.createElement("a:xfrm");
        Element aOff = document.createElement("a:off");
        aOff.setAttribute("x", "539750");
        aOff.setAttribute("y", "44450");
        Element aExt = document.createElement("a:ext");
        aExt.setAttribute("cx", "8064500");
        aExt.setAttribute("cy", "936625");
        xfrm.appendChild(aOff);
        xfrm.appendChild(aExt);
        return xfrm;
    }
}
