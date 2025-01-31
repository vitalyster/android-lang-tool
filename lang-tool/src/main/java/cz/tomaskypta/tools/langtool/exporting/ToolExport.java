package cz.tomaskypta.tools.langtool.exporting;

import java.io.*;
import java.util.*;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;


public class ToolExport {

    private static final String DIR_VALUES = "values";
    private static final String[] POTENTIAL_RES_DIRS = new String[]{"res", "src/main/res"};

    private DocumentBuilder builder;
    private File outExcelFile;
    private String project;
    private Map<String, Integer> keysIndex;
    private PrintStream out;
    private ExportConfig mConfig;
    private Set<String> sAllowedFiles = new HashSet<String>();

    {
        sAllowedFiles.add("strings.xml");
    }

    public ToolExport(PrintStream out) throws ParserConfigurationException {
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        builder = dbf.newDocumentBuilder();
        this.out = out == null ? System.out : out;
    }

    public static void run(ExportConfig config) throws SAXException,
        IOException, ParserConfigurationException {
        run(null, config);
    }

    public static void run(PrintStream out, ExportConfig config) throws SAXException, IOException, ParserConfigurationException {
        ToolExport tool = new ToolExport(out);
        if (StringUtils.isEmpty(config.inputExportProject)) {
            tool.out.println("Cannot export, missing config");
            return;
        }
        File project = new File(config.inputExportProject);
        if (StringUtils.isEmpty(config.outputFile)) {
            config.outputFile = "exported_strings_" + System.currentTimeMillis() + ".xls";
        }
        tool.outExcelFile = new File(config.outputFile);
        tool.project = project.getName();
        tool.mConfig = config;
        tool.sAllowedFiles.addAll(config.additionalResources);
        tool.export(project);
    }

    private void export(File project) throws SAXException, IOException {
        File res = findResourceDir(project);
        if (res == null) {
            System.err.println("Cannot find resource directory.");
            return;
        }
        Optional<File> defValuesDir = Arrays.stream(res.listFiles()).filter(i -> i.getName().equals(DIR_VALUES))
                .findFirst();
        if (defValuesDir.isPresent()) {
            keysIndex = exportDefLang(defValuesDir.get());
        }
        for (File dir : res.listFiles()) {
            if (!dir.isDirectory() || !dir.getName().startsWith(DIR_VALUES)) {
                continue;
            }
            String dirName = dir.getName();
            if (!dirName.equals(DIR_VALUES)) {
                int index = dirName.indexOf('-');
                if (index == -1)
                    continue;
                String lang = dirName.substring(index + 1);
                exportLang(lang, dir);
            }
        }
    }

    private File findResourceDir(File project) {
        List<File> availableResDirs = new LinkedList<File>();
        for (String potentialResDir : POTENTIAL_RES_DIRS) {
            File res = new File(project, potentialResDir);
            if (res.exists()) {
                availableResDirs.add(res);
            }
        }
        if (!availableResDirs.isEmpty()) {
            return availableResDirs.get(0);
        }
        return null;
    }

    private void exportLang(String lang, File valueDir) throws IOException, SAXException {
        for (String fileName : sAllowedFiles) {
            File stringFile = new File(valueDir, fileName);
            if (!stringFile.exists()) {
                continue;
            }
            exportLangToExcel(project, lang, stringFile, getStrings(stringFile), outExcelFile, keysIndex);
        }
    }

    private Map<String, Integer> exportDefLang(File valueDir) throws IOException, SAXException {
        Map<String, Integer> keys = new HashMap<String, Integer>();
        Workbook wb = WorkbookFactory.create(outExcelFile.getName().endsWith("x"));

        Sheet sheet;
        sheet = wb.createSheet(project);
        int rowIndex = 0;
        sheet.createRow(rowIndex++);
        createTilte(wb, sheet);
        addLang2Tilte(wb, sheet, "default");
        sheet.createFreezePane(1, 1);

        FileOutputStream outFile = new FileOutputStream(outExcelFile);
        wb.write(outFile);
        outFile.close();

        for (String fileName : sAllowedFiles) {
            File stringFile = new File(valueDir, fileName);
            if (!stringFile.exists()) {
                continue;
            }
            keys.putAll(exportDefLangToExcel(rowIndex, project, stringFile, getStrings(stringFile), outExcelFile));
        }


        return keys;
    }

    private NodeList getStrings(File f) throws SAXException, IOException {
        Document dom = builder.parse(f);
        return dom.getDocumentElement().getChildNodes();
    }

    private static CellStyle createTitleStyle(Workbook wb) {
        Font bold = wb.createFont();
        bold.setBold(true);

        CellStyle style = wb.createCellStyle();
        style.setFont(bold);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        return style;
    }

    private static CellStyle createCommentStyle(Workbook wb) {

        Font commentFont = wb.createFont();
        commentFont.setColor(IndexedColors.GREEN.getIndex());
        commentFont.setItalic(true);
        commentFont.setFontHeightInPoints((short) 12);

        CellStyle commentStyle = wb.createCellStyle();
        commentStyle.setFont(commentFont);
        return commentStyle;
    }

    private static CellStyle createPlurarStyle(Workbook wb) {

        Font commentFont = wb.createFont();
        commentFont.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        commentFont.setItalic(true);
        commentFont.setFontHeightInPoints((short)12);

        CellStyle commentStyle = wb.createCellStyle();
        commentStyle.setFont(commentFont);
        return commentStyle;
    }

    private static CellStyle createKeyStyle(Workbook wb) {
        Font bold = wb.createFont();
        bold.setBold(true);
        bold.setFontHeightInPoints((short)11);

        CellStyle keyStyle = wb.createCellStyle();
        keyStyle.setFont(bold);

        return keyStyle;
    }

    private static CellStyle createTextStyle(Workbook wb) {
        Font plain = wb.createFont();
        plain.setFontHeightInPoints((short)12);

        CellStyle textStyle = wb.createCellStyle();
        textStyle.setFont(plain);

        return textStyle;
    }

    private static CellStyle createMissedStyle(Workbook wb) {

        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return style;
    }

    private static void createTilte(Workbook wb, Sheet sheet) {
        Row titleRow = sheet.getRow(0);

        Cell cell = titleRow.createCell(0);
        cell.setCellStyle(createTitleStyle(wb));
        cell.setCellValue("KEY");

        sheet.setColumnWidth(cell.getColumnIndex(), (40 * 256));
    }

    private static void addLang2Tilte(Workbook wb, Sheet sheet, String lang) {
        Row titleRow = sheet.getRow(0);
        Cell lastCell = titleRow.getCell((int) titleRow.getLastCellNum() - 1);
        if (lang.equals(lastCell.getStringCellValue())) {
            // language column already exists
            return;
        }
        Cell cell = titleRow.createCell((int)titleRow.getLastCellNum());
        cell.setCellStyle(createTitleStyle(wb));
        cell.setCellValue(lang);

        sheet.setColumnWidth(cell.getColumnIndex(), (60 * 256));
    }


    private Map<String, Integer> exportDefLangToExcel(int rowIndex, String project, File src, NodeList strings, File f) throws IOException {
        out.println();
        out.println("Start processing DEFAULT language " + src.getName());

        Map<String, Integer> keys = new HashMap<String, Integer>();

        Workbook wb = WorkbookFactory.create(new FileInputStream(f));

        CellStyle commentStyle = createCommentStyle(wb);
        CellStyle plurarStyle = createPlurarStyle(wb);
        CellStyle keyStyle = createKeyStyle(wb);
        CellStyle textStyle = createTextStyle(wb);

        Sheet sheet = wb.getSheet(project);


        for (int i = 0; i < strings.getLength(); i++) {
            Node item = strings.item(i);
            if (item.getNodeType() == Node.TEXT_NODE) {

            }
            if (item.getNodeType() == Node.COMMENT_NODE) {
                Row row = sheet.createRow(rowIndex++);
                Cell cell = row.createCell(0);
                cell.setCellValue(String.format("/** %s **/", item.getTextContent()));
                cell.setCellStyle(commentStyle);

                sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 255));
            }

            if ("string".equals(item.getNodeName())) {
                Node translatable = item.getAttributes().getNamedItem("translatable");
                if (translatable != null && "false".equals(translatable.getNodeValue())) {
                    continue;
                }

                String key = getKey(item);
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                keys.put(key, rowIndex);

                Row row = sheet.createRow(rowIndex++);

                Cell cell = row.createCell(0);
                cell.setCellValue(key);
                cell.setCellStyle(keyStyle);

                cell = row.createCell(1);
                cell.setCellStyle(textStyle);
                cell.setCellValue(item.getTextContent().replace("\\'", "'").replace("\\\"", "\""));
            } else if ("plurals".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                String pluralName = key;

                Row row = sheet.createRow(rowIndex++);
                Cell cell = row.createCell(0);
                cell.setCellValue(String.format("//plurals: %s", pluralName));
                cell.setCellStyle(plurarStyle);

                NodeList items = item.getChildNodes();
                for (int j = 0; j < items.getLength(); j++) {
                    Node plurarItem = items.item(j);
                    if ("item".equals(plurarItem.getNodeName())) {
                        String itemKey = pluralName + "#" + plurarItem.getAttributes().getNamedItem("quantity").getNodeValue();
                        keys.put(itemKey, rowIndex);

                        Row itemRow = sheet.createRow(rowIndex++);

                        Cell itemCell = itemRow.createCell(0);
                        itemCell.setCellValue(itemKey);
                        itemCell.setCellStyle(keyStyle);

                        itemCell = itemRow.createCell(1);
                        itemCell.setCellStyle(textStyle);
                        itemCell.setCellValue(plurarItem.getTextContent());
                    }
                }
            } else if ("string-array".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                if (mConfig.isIgnoredKey(key)) {
                    continue;
                }
                NodeList arrayItems = item.getChildNodes();
                for (int j = 0, k = 0; j < arrayItems.getLength(); j++) {
                    Node arrayItem = arrayItems.item(j);
                    if ("item".equals(arrayItem.getNodeName())) {
                        String itemKey = key + "[" + k++ + "]";
                        keys.put(itemKey, rowIndex);

                        Row itemRow = sheet.createRow(rowIndex++);

                        Cell itemCell = itemRow.createCell(0);
                        itemCell.setCellValue(itemKey);
                        itemCell.setCellStyle(keyStyle);

                        itemCell = itemRow.createCell(1);
                        itemCell.setCellStyle(textStyle);
                        itemCell.setCellValue(arrayItem.getTextContent());
                    }
                }
            }
        }

        FileOutputStream outFile = new FileOutputStream(f);
        wb.write(outFile);
        outFile.close();

        out.println("DEFAULT language was precessed");
        return keys;
    }

    private void exportLangToExcel(String project, String lang, File src, NodeList strings, File f, Map<String, Integer> keysIndex) throws IOException {
        out.println();
        out.println(String.format("Start processing: '%s' %s", lang, src.getName()));
        Set<String> missedKeys = new HashSet<String>(keysIndex.keySet());

        Workbook wb = WorkbookFactory.create(new FileInputStream(f));

        CellStyle textStyle = createTextStyle(wb);

        Sheet sheet = wb.getSheet(project);
        addLang2Tilte(wb, sheet, lang);

        Row titleRow = sheet.getRow(0);
        int lastColumnIdx = (int)titleRow.getLastCellNum() - 1;

        for (int i = 0; i < strings.getLength(); i++) {
            Node item = strings.item(i);

            if ("string".equals(item.getNodeName())) {
                Node translatable = item.getAttributes().getNamedItem("translatable");
                if (translatable != null && "false".equals(translatable.getNodeValue())) {
                    continue;
                }
                String key = getKey(item);
                Integer index = keysIndex.get(key);
                if (index == null) {
                    out.println("\t" + key + " - row does not exist");
                    continue;
                }

                missedKeys.remove(key);
                Row row = sheet.getRow(index);

                Cell cell = row.createCell(lastColumnIdx);
                cell.setCellValue(item.getTextContent());
                cell.setCellStyle(textStyle);
            } else if ("plurals".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                String plurarName = key;

                NodeList items = item.getChildNodes();
                for (int j = 0; j < items.getLength(); j++) {
                    Node pluralItem = items.item(j);
                    if ("item".equals(pluralItem.getNodeName())) {
                        key = plurarName + "#" + pluralItem.getAttributes().getNamedItem("quantity").getNodeValue();
                        Integer index = keysIndex.get(key);
                        if (index == null) {
                            out.println("\t" + key + " - row does not exist");
                            continue;
                        }
                        missedKeys.remove(key);

                        Row row = sheet.getRow(index);

                        Cell cell = row.createCell(lastColumnIdx);
                        cell.setCellValue(pluralItem.getTextContent());
                        cell.setCellStyle(textStyle);
                    }
                }
            } else if ("string-array".equals(item.getNodeName())) {
                String key = item.getAttributes().getNamedItem("name").getNodeValue();
                NodeList arrayItems = item.getChildNodes();
                for (int j = 0, k = 0; j < arrayItems.getLength(); j++) {
                    Node arrayItem = arrayItems.item(j);
                    if ("item".equals(arrayItem.getNodeName())) {
                        String itemKey = key + "[" + k++ + "]";
                        Integer rowIndex = keysIndex.get(itemKey);
                        if (rowIndex == null) {
                            out.println("\t" + key + " - row does not exist");
                            continue;
                        }
                        missedKeys.remove(key);

                        Row itemRow = sheet.getRow(rowIndex);

                        Cell cell = itemRow.createCell(lastColumnIdx);
                        cell.setCellValue(arrayItem.getTextContent());
                        cell.setCellStyle(textStyle);
                    }
                }
            }
        }

        CellStyle missedStyle = createMissedStyle(wb);

        if (!missedKeys.isEmpty()) {
            out.println("  MISSED KEYS:");
        }
        for (String missedKey : missedKeys) {
            out.println("\t" + missedKey);
            Integer index = keysIndex.get(missedKey);
            Row row = sheet.getRow(index);
            Cell cell = row.createCell((int)row.getLastCellNum());
            cell.setCellStyle(missedStyle);
        }

        FileOutputStream outStream = new FileOutputStream(f);
        wb.write(outStream);
        outStream.close();

        if (missedKeys.isEmpty()) {
            out.println(String.format("'%s' was processed", lang));
        } else {
            out.println(String.format("'%s' was processed with MISSED KEYS - %d", lang, missedKeys.size()));
        }
    }

    private String getKey(Node item) {
        String key = item.getAttributes().getNamedItem("name").getNodeValue();
        NodeList nodes = item.getChildNodes();
        if (nodes.getLength() == 0) {
            throw new IllegalArgumentException("Unpredictable node format at " + item);
        }
        Node text = nodes.item(0);
        if (text.getNodeType() == Node.CDATA_SECTION_NODE) {
            key += "!cdata";
        }

        return key;
    }
}
