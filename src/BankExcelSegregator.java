import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class BankExcelSegregator {
    public static void main(String[] args) throws Exception {
        //String inputFile = "C:\Users\mishsur\Downloads\Ticket_Status_18082025.xlsx"; // Use .csv

        Scanner scan = new Scanner(System.in);
        System.out.println("Enter input xlsx file complete path:");
        String inputFile = scan.nextLine();

        String outputFilePath = inputFile.substring(0, inputFile.lastIndexOf("\\") + 1);
        System.out.println("Enter output file complete location. Press enter to continue with default path: " + outputFilePath);
        String outputFilePathEntered = scan.nextLine();
        if(null != outputFilePathEntered && !outputFilePathEntered.isEmpty())
            outputFilePath=outputFilePathEntered;

        String workbookSheetName = "Banker-SPOC";
        System.out.println("Enter sheet name which has data. Press enter to continue with default sheet: " + workbookSheetName);
        String workbookSheetNameEntered = scan.nextLine();
        if(null != workbookSheetNameEntered && !workbookSheetNameEntered.isEmpty())
            workbookSheetName=workbookSheetNameEntered;


        FileInputStream fis = new FileInputStream(inputFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(workbookSheetName);

        // Find "Bank Name" column index
        Row headerRow = sheet.getRow(0);
        int bankColIdx = -1;
        for (Cell cell : headerRow) {
            System.out.println("Header columns:" + cell.getStringCellValue());
            if (cell.getStringCellValue().trim().equalsIgnoreCase("Bank Name")) {
                bankColIdx = cell.getColumnIndex();
                break;
            }
        }
        if (bankColIdx == -1) throw new RuntimeException("Bank Name column not found");

        // Collect unique bank names and their rows
        Map<String, List<Row>> bankRows = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            Cell bankCell = row.getCell(bankColIdx);
            if (bankCell == null) {
                System.out.println("Skipping row " + i + " as Bank Name cell is empty");
                continue;
            }
            String bankName = bankCell.getStringCellValue().trim();
            bankRows.computeIfAbsent(bankName, k -> new ArrayList<>()).add(row);
        }
        System.out.println("Found total " + bankRows.size() + " banks: " + bankRows.keySet());

        System.out.println("\nNumber of rows per bank:");


        Workbook outWorkbook = new XSSFWorkbook();

        // Write separate Excel files for each bank
        for (Map.Entry<String, List<Row>> entry : bankRows.entrySet()) {
            String bankName = entry.getKey();
            List<Row> rows = entry.getValue();

            System.out.println(bankName + "=" + rows.size());

            //Creating new excel workbook and sheet
            //Workbook outWorkbook = new XSSFWorkbook();
            Sheet outSheet = outWorkbook.createSheet(bankName);

            // Copy header
            Row outHeader = outSheet.createRow(0);
            for (Cell cell : headerRow) {
                Cell newCell = outHeader.createCell(cell.getColumnIndex());
                newCell.setCellValue(cell.getStringCellValue());
                CellStyle srcStyle = cell.getCellStyle();
                if (srcStyle != null) {
                    CellStyle destStyle = outWorkbook.createCellStyle();
                    destStyle.cloneStyleFrom(srcStyle);
                    newCell.setCellStyle(destStyle);
                }

            }

            // Copy rows
            int outRowNum = 1;
            Map<CellStyle, CellStyle> styleMap = new HashMap<>();
            for (Row row : rows) {
                Row outRow = outSheet.createRow(outRowNum++);
                for (Cell cell : row) {
                    Cell outCell = outRow.createCell(cell.getColumnIndex());
                    CellStyle srcStyle = cell.getCellStyle();
                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        outCell.setCellValue(cell.getDateCellValue());
                    } else {
                        switch (cell.getCellType()) {
                            case STRING -> outCell.setCellValue(cell.getStringCellValue());
                            case NUMERIC -> outCell.setCellValue(cell.getNumericCellValue());
                            case BOOLEAN -> outCell.setCellValue(cell.getBooleanCellValue());
                            default -> outCell.setCellValue(cell.toString());
                        }
                    }
                    if (srcStyle != null) {
                        CellStyle destStyle = styleMap.computeIfAbsent(srcStyle, s -> {
                            CellStyle newStyle = outWorkbook.createCellStyle();
                            newStyle.cloneStyleFrom(s);
                            return newStyle;
                        });
                        outCell.setCellStyle(destStyle);
                    }
                }
            }
        }
        // Save file
        String outFileName = outputFilePath + "output.xlsx";
        try (FileOutputStream fos = new FileOutputStream(outFileName)) {
            outWorkbook.write(fos);
        }
        outWorkbook.close();

        System.out.println("\nSegregation completed.");
        System.out.println("Total banks: " + bankRows.size());
        System.out.println("Total rows processed excluding header: " + bankRows.values().stream().mapToInt(List::size).sum());
        workbook.close();
        fis.close();
    }
}
