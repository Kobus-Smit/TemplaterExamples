package hr.ngs.templater.example;

import hr.ngs.templater.Configuration;
import hr.ngs.templater.DocumentFactoryBuilder;
import hr.ngs.templater.TemplateDocument;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.w3c.dom.*;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.*;

public class HtmlWordExample {

    private static Element convert(String html, DocumentBuilder dBuilder, Map<String, String> links) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
            XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
            List<Object> ooxml = importer.convert(html, null);
            wordMLPackage.getMainDocumentPart().getContent().addAll(ooxml);
            String xml = XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement().getBody(), true, false);
            Document doc = dBuilder.parse(new InputSource(new StringReader(xml)));
            if (links != null) {
                for(Relationship rel : wordMLPackage.getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship()) {
                    if ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink".equals(rel.getType())) {
                        links.put(rel.getId(), rel.getTarget());
                    }
                }
            }
            return doc.getDocumentElement();
        } catch (Docx4JException e) {
            throw new RuntimeException(e);
        } catch (SAXException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static class SimpleHtmlConverter implements DocumentFactoryBuilder.Formatter {
        private final DocumentBuilder dBuilder;
        public SimpleHtmlConverter(DocumentBuilder dBuilder) {
            this.dBuilder = dBuilder;
        }

        @Override
        public Object format(Object value, String metadata) {
            if (metadata.equals("simple-html")) {
                return convert(value.toString(), dBuilder, null).getFirstChild();
            }
            return value;
        }
    }

    private static class ComplexHtmlConverter implements DocumentFactoryBuilder.Formatter {
        private final DocumentBuilder dBuilder;
        public ComplexHtmlConverter(DocumentBuilder dBuilder) {
            this.dBuilder = dBuilder;
        }

        private void rewriteHyperlinks(Element element, Map<String, String> links) {
            if (element.getNodeName().equals("w:hyperlink")) {
                String url = links.get(element.getAttribute("r:id"));
                if (url == null) return;
                Element parent = (Element) element.getParentNode();
                Element fldSimple = element.getOwnerDocument().createElement("w:fldSimple");
                fldSimple.setAttribute("w:instr", " HYPERLINK " + url +" ");
                Node child = element.getFirstChild();
                while (child != null) {
                    element.removeChild(child);
                    fldSimple.appendChild(child);
                    child = child.getNextSibling();
                }
                parent.replaceChild(fldSimple, element);
            } else {
                Node next = element.getFirstChild();
                while (next != null) {
                    if (next instanceof Element) {
                        rewriteHyperlinks((Element) next, links);
                    }
                    next = next.getNextSibling();
                }
            }
        }

        @Override
        public Object format(Object value, String metadata) {
            if (metadata.equals("complex-html")) {
                HashMap<String, String> links = new HashMap<>();
                NodeList bodyNodes = convert(value.toString(), dBuilder, links).getChildNodes();
                List<Element> elements = new ArrayList<>(bodyNodes.getLength());
                for (int i = 0; i < bodyNodes.getLength(); i++) {
                    Element element = (Element) bodyNodes.item(i);
                    if (!links.isEmpty()) {
                        rewriteHyperlinks(element, links);
                    }
                    elements.add(element);
                }
                //lets put special attribute directly on XML so we don't need to put it on tag
                elements.get(0).setAttribute("templater-xml", "remove-old-xml");
                return elements.toArray(new Element[0]);
            }
            return value;
        }
    }

    public static void main(final String[] args) throws Exception {
        DocumentBuilder dBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();

        InputStream templateStream = HtmlWordExample.class.getResourceAsStream("/template.docx");
        File tmp = File.createTempFile("html", ".docx");

        Path embeddedHtml = Files.createTempFile("embed", ".html");
        Files.copy(HtmlWordExample.class.getResourceAsStream("/example.html"), embeddedHtml, StandardCopyOption.REPLACE_EXISTING);

        FileOutputStream fos = new FileOutputStream(tmp);
        TemplateDocument tpl =
                Configuration.builder()
                        .include(new SimpleHtmlConverter(dBuilder))
                        .include(new ComplexHtmlConverter(dBuilder))
                        .build().open(templateStream, "docx", fos);
        tpl.process(new HashMap<String, Object>() {{
            put("Html1", "<html>\n" +
                    "<head>title</head>\n" +
                    "<body>\n" +
                    "some <strong>text</strong> in <span style=\"color:red\">red!</span>\n" +
                    "</body>\n" +
                    "</html>");
            put("Html2", "<html>\n" +
                    "<head>title</head>\n" +
                    "<body>\n" +
                    "<ul>\n" +
                    "        <li>Number 1</li>\n" +
                    "        <li>Number 2</li>\n" +
                    "</ul>\n" +
                    "<a href=\"https://templater.info/\">Templater</a>\n" +
                    "</body>\n" +
                    "</html>");
            put("Html3", embeddedHtml.toFile());
        }});
        tpl.close();
        fos.close();
        Files.delete(embeddedHtml);//once finished we can delete the file
        java.awt.Desktop.getDesktop().open(tmp);
    }
}
