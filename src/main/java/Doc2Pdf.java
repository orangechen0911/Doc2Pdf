import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontProvider;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.select.Elements;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.charset.StandardCharsets;

/**
 * 使用poi+itextpdf进行word转pdf
 * 先将word转成html，再将html转成pdf
 */
public class Doc2Pdf {

    /**
     * 将doc格式文件转成html
     *
     * @param docPath  doc文件路径
     * @param imageDir doc文件中图片存储目录
     * @return html
     */
    static String doc2Html(String docPath, String imageDir) throws IOException, TransformerException, ParserConfigurationException {
        File file = new File(imageDir);
        if (!file.exists()) {
            file.mkdirs();
        }
        HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(docPath));
        org.w3c.dom.Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
        //保存图片，并返回图片的相对路径
        wordToHtmlConverter.setPicturesManager((content, pictureType, name, width, height) -> {
            try (FileOutputStream out = new FileOutputStream(imageDir + name)) {
                out.write(content);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return imageDir + name;
        });
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        wordToHtmlConverter.processDocument(wordDocument);
        org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(baos);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);

        String content = new String(baos.toByteArray());
        baos.close();
        return content;
    }

    /**
     * 将docx格式文件转成html
     *
     * @param docxPath docx文件路径
     * @param imageDir docx文件中图片存储目录
     * @return html
     */
    static String docx2Html(String docxPath, String imageDir) {
        String content = null;

        FileInputStream in = null;
        ByteArrayOutputStream baos = null;
        try {
// 1> 加载文档到XWPFDocument
            in = new FileInputStream(new File(docxPath));
            XWPFDocument document = new XWPFDocument(in);
// 2> 解析XHTML配置（这里设置IURIResolver来设置图片存放的目录）
            XHTMLOptions options = XHTMLOptions.create();
// 存放word中图片的目录
            options.setExtractor(new FileImageExtractor(new File(imageDir)));
            options.URIResolver(new BasicURIResolver(imageDir));
            options.setIgnoreStylesIfUnused(false);
            options.setFragment(true);
// 3> 将XWPFDocument转换成XHTML
            baos = new ByteArrayOutputStream();
            XHTMLConverter.getInstance().convert(document, baos, options);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (in != null) {
                    in.close();
                }
                if (baos != null) {
                    content = new String(baos.toByteArray(), StandardCharsets.UTF_8);
                    baos.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return content;
    }

    /**
     * 使用jsoup规范化html
     *
     * @param html html内容
     * @return 规范化后的html
     */
    static String formatHtml(String html) {
        org.jsoup.nodes.Document doc = Jsoup.parse(html);
// 去除过大的宽度
        String style = doc.attr("style");
        if (StringUtils.isNotEmpty(style) && style.contains("width")) {
            doc.attr("style", "");
        }
        Elements divs = doc.select("div");
        for (Element div : divs) {
            String divStyle = div.attr("style");
            if (StringUtils.isNotEmpty(divStyle) && divStyle.contains("width")) {
                div.attr("style", "");
            }
        }
// jsoup生成闭合标签
        doc.outputSettings().syntax(org.jsoup.nodes.Document.OutputSettings.Syntax.xml);
        doc.outputSettings().escapeMode(Entities.EscapeMode.xhtml);
        return doc.html();
    }

    /**
     * html转成pdf
     *
     * @param html          html
     * @param outputPdfPath 输出pdf路径
     */
    static void htmlToPdf(String html, String outputPdfPath) {
        com.itextpdf.text.Document document = null;
        try {
// 纸
            document = new com.itextpdf.text.Document(PageSize.A4);
// 笔
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));
            document.open();
// html转pdf
            XMLWorkerHelper.getInstance().parseXHtml(writer, document, new ByteArrayInputStream(html.getBytes()),
                    StandardCharsets.UTF_8, new FontProvider() {
                        @Override
                        public boolean isRegistered(String s) {
                            return false;
                        }

                        @Override
                        public Font getFont(String s, String s1, boolean embedded, float size, int style, BaseColor baseColor) {
// 配置字体
                            Font font = null;
                            try {
                                BaseFont bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", BaseFont.EMBEDDED);
                                font = new Font(bf, size, style, baseColor);
                                font.setColor(baseColor);
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                            return font;
                        }
                    });
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (document != null) {
                document.close();
            }
        }
    }

    /**
     * 删除临时文件
     *
     * @param file
     */
    static void deleteFiles(File file) {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            for (File value : files) {
                if (value.isDirectory()) {
                    deleteFiles(value);
                } else {
                    value.delete();
                }
            }
        }
        file.delete();
    }

    /**
     * 将doc转成pdf
     *
     * @param filePath
     */
    public static void docToPdf(String filePath) throws ParserConfigurationException, TransformerException, IOException {

        String file = doc2Html(filePath, filePath.substring(0, filePath.lastIndexOf(File.separator)) + File.separator + "image" + File.separator);
        file = formatHtml(file);
        htmlToPdf(file, filePath.substring(0, filePath.indexOf(".")) + ".pdf");
        deleteFiles(new File(filePath.substring(0, filePath.lastIndexOf(File.separator)) + File.separator + "image" + File.separator));
    }

    /**
     * 将docx转成pdf
     *
     * @param filePath
     */
    public static void docxToPdf(String filePath) {
        String file = docx2Html(filePath, filePath.substring(0, filePath.lastIndexOf(File.separator)) + File.separator + "image" + File.separator);
        file = formatHtml(file);
        htmlToPdf(file, filePath.substring(0, filePath.indexOf(".")) + ".pdf");
        deleteFiles(new File(filePath.substring(0, filePath.lastIndexOf(File.separator)) + File.separator + "image" + File.separator));
    }
}
