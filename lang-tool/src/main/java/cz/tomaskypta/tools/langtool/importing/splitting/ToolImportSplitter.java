package cz.tomaskypta.tools.langtool.importing.splitting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import cz.tomaskypta.tools.langtool.importing.ImportConfig;
import cz.tomaskypta.tools.langtool.importing.ToolImport;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

/**
 * Created by Tomas Kypta on 19.09.14.
 */
public class ToolImportSplitter {

    private TreeMap<Integer, String> mSplittingMap;
    private HashMap<String, String> mOutputFileNames;
    private File mIntermediateXlsDir;

    public static void run(SplittingConfig config) throws IOException,
        ParserConfigurationException, TransformerException {
        if (config == null) {
            System.err.println("Cannot split, missing config");
            return;
        }

        if (StringUtils.isEmpty(config.inputFile)) {
            System.err.println("Cannot split, missing input file");
            return;
        }

        if (StringUtils.isEmpty(config.splittingConfigFile)) {
            System.err.println("Cannot split, missing splitting config. Importing instead.");
            ToolImport.run(config);
            return;
        }

        Workbook wb = WorkbookFactory.create(new FileInputStream(new File(config.inputFile)));
        Sheet sheet = wb.getSheetAt(0);

        Workbook wbConfig = WorkbookFactory.create(new FileInputStream(new File(config.splittingConfigFile)));
        Sheet sheetConfig = wbConfig.getSheetAt(0);

        ToolImportSplitter tool = new ToolImportSplitter();
        tool.mIntermediateXlsDir = new File("intermediate");
        tool.mIntermediateXlsDir.mkdirs();

        tool.prepareSplittingMap(sheetConfig);
        tool.split(sheet);

        for (String file : tool.mSplittingMap.values()) {
            File outputFile = new File(tool.mIntermediateXlsDir, file);
            System.out.println("Importing file: " + file);

            ImportConfig partConfig = new ImportConfig(config);
            partConfig.inputFile = outputFile.getPath();
            partConfig.outputDirName = outputFile.getName().substring(0, outputFile.getName().indexOf('.'));
            partConfig.outputFileName = tool.mOutputFileNames.get(file);
            ToolImport.run(partConfig);
        }
    }

    private void prepareSplittingMap(Sheet sheetConfig) throws IOException, TransformerException {
        mSplittingMap = new TreeMap<Integer, String>();
        mOutputFileNames = new HashMap<String, String>();
        Iterator<Row> it = sheetConfig.rowIterator();
        while (it.hasNext()) {
            Row row = it.next();
            if (row == null || row.getCell(0) == null || row.getCell(1) == null) {
                return;
            }
            String splitName = row.getCell(1).getStringCellValue();
            mSplittingMap.put((int)row.getCell(0).getNumericCellValue(), splitName);
            if (row.getCell(2) != null) {
                mOutputFileNames.put(splitName, row.getCell(2).getStringCellValue());
            }
        }
    }

    private void split(Sheet inSheet) throws IOException, TransformerException {
        Row inTitleRow = inSheet.getRow(0);
        for (Map.Entry<Integer, String> entry : mSplittingMap.entrySet()) {
            System.out.println("Splitting into file: " + entry.getValue());
            File outputFile = new File(mIntermediateXlsDir, entry.getValue());
            FileOutputStream fos = null;

            try {
                fos = new FileOutputStream(outputFile);

                Workbook wb = WorkbookFactory.create(outputFile.getName().endsWith("x"));
                Sheet outSheet = wb.createSheet(inSheet.getSheetName());
                copyTitleRow(inTitleRow, outSheet);

                Integer actFileStart = entry.getKey();
                Integer nextFileStart = mSplittingMap.higherKey(entry.getKey());
                if (nextFileStart == null) {
                    nextFileStart = inSheet.getLastRowNum() + 2;
                }

                copyRowRange(inSheet, outSheet, actFileStart, nextFileStart);

                wb.write(fos);
            } finally {
                if (fos != null) {
                    fos.close();
                }
            }
        }
    }

    private void copyTitleRow(Row inTitleRow, Sheet outSheet) {
        Row outTitleRow = outSheet.createRow(0);
        copyRow(inTitleRow, outTitleRow);
    }

    private void copyRowRange(Sheet inSheet, Sheet outSheet, int rowStart, int rowEnd) {
        for (int rowIdx = rowStart, outRowIdx = 1; rowIdx < rowEnd; rowIdx++, outRowIdx++) {
            Row outRow = outSheet.createRow(outRowIdx);
            Row inRow = inSheet.getRow(rowIdx-1);
            copyRow(inRow, outRow);
        }
    }

    private void copyRow(Row inTitleRow, Row outRow) {
        // TODO copy formatting
        Iterator<Cell> it = inTitleRow.cellIterator();
        while (it.hasNext()) {
            Cell srcCell = it.next();
            outRow.createCell(srcCell.getColumnIndex(), srcCell.getCellType());
            outRow.getCell(srcCell.getColumnIndex()).setCellValue(srcCell.getStringCellValue());
        }
    }
}
