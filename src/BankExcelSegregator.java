import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

public class BankExcelSegregator {

    private static final Map<String, List<Row>> BANK_NAME_ROWS_MAP = new TreeMap<>();
    private static final Map<String, Integer> TEAM_NAME_TICKET_COUNT_MAP = new TreeMap<>();

    private static final List<TicketAgeBoundDetails> TICKET_AGE_BOUND_DETAILS_BOUND_DETAILS_LIST = List.of( new TicketAgeBoundDetails(0,30,"upto 30 Days"),
            new TicketAgeBoundDetails(31,60,"31 to 60 Days"),
            new TicketAgeBoundDetails(61,90,"61 to 90 Days"),
            new TicketAgeBoundDetails(91,180,"91 to 180 Days"),
            new TicketAgeBoundDetails(181,365,"181 to 365 Days"),
            new TicketAgeBoundDetails(366,Integer.MAX_VALUE,"More than 1 year")
    );

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
        Map<String, Integer> headerMap = fetchHeaders(sheet);
        int bankColIdx = headerMap.get("Bank Name");
        if (bankColIdx == -1) throw new RuntimeException("Bank Name column not found");

        collectUniqueBankTeamAndAgeBound(sheet, headerMap);

        System.out.println("Found total " + BANK_NAME_ROWS_MAP.size() + " banks: " + BANK_NAME_ROWS_MAP.keySet());

        System.out.println("\nNumber of rows per bank:");


        Workbook outWorkbook = new XSSFWorkbook();

        createDashBoard(outWorkbook);

        Row headerRow = sheet.getRow(0);

        // Write separate Excel files for each bank
        for (Map.Entry<String, List<Row>> entry : BANK_NAME_ROWS_MAP.entrySet()) {
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
            int ageColIndex = headerMap.get("Age");
            rows.sort(new RowsAgeComparator(ageColIndex));
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
        System.out.println("Total banks: " + BANK_NAME_ROWS_MAP.size());
        System.out.println("Total rows processed excluding header: " + BANK_NAME_ROWS_MAP.values().stream().mapToInt(List::size).sum());
        workbook.close();
        fis.close();
    }

    public static Map<String, Integer> fetchHeaders(Sheet sheet) {
        Map<String, Integer> headerMap = new HashMap<>();
        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            System.out.println("Header columns:" + cell.getStringCellValue());
            headerMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
        }

        return headerMap;
    }

    public static void collectUniqueBankTeamAndAgeBound(Sheet sheet, Map<String, Integer> headerMap) {

        int bankColIdx = headerMap.get("Bank Name");
        int teamColIdx = headerMap.get("Team");
        int ageColIndex = headerMap.get("Age");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            // Collect unique bank names and their rows
            Cell bankCell = row.getCell(bankColIdx);
            if (bankCell == null || bankCell.getCellType() == CellType.BLANK) {
                System.out.println("Skipping row " + (i+1) + " as Bank Name cell is empty. Hence will not be processing Team name or Age also.");
                continue;
            } else {
                String bankName = bankCell.getStringCellValue().trim();
                BANK_NAME_ROWS_MAP.computeIfAbsent(bankName, k -> new ArrayList<>()).add(row);
            }

            // Collect unique Team Name
            Cell teamNameCell = row.getCell(teamColIdx);
            if (teamNameCell == null || teamNameCell.getCellType() == CellType.BLANK) {
                System.out.println("Team Name cell is empty for row " + (i+1));
            } else {
                String teamName = teamNameCell.getStringCellValue().trim();
                TEAM_NAME_TICKET_COUNT_MAP.putIfAbsent(teamName, 0);
                TEAM_NAME_TICKET_COUNT_MAP.put(teamName, TEAM_NAME_TICKET_COUNT_MAP.get(teamName) + 1);
            }

            Cell ageCell = row.getCell(ageColIndex);
            if (ageCell == null || ageCell.getCellType() == CellType.BLANK) {
                System.out.println("Age cell is empty for row " + (i+1));
                TICKET_AGE_BOUND_DETAILS_BOUND_DETAILS_LIST.forEach(ticket -> ticket.checkAndIncrementCount(0));
            } else {
                int ticketAge = Integer.parseInt(ageCell.getStringCellValue().replace("Day(s)","").trim());
                TICKET_AGE_BOUND_DETAILS_BOUND_DETAILS_LIST.forEach(ticket -> ticket.checkAndIncrementCount(ticketAge));
            }
        }
    }

    public static void createDashBoard(Workbook outWorkbook) {

        System.out.println("\nCreating Dashboard:");

        CellStyle cellStyle = outWorkbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        Font headerFont = outWorkbook.createFont();
        headerFont.setBold(true);

        // Create header cell style with bold font and borders
        CellStyle headerStyle = outWorkbook.createCellStyle();
        headerStyle.cloneStyleFrom(cellStyle);
        headerStyle.setFont(headerFont);

        Sheet dashboardSheet = outWorkbook.createSheet("Dashboard");
        Row DashBoardSheetHeader = dashboardSheet.createRow(0);

        Cell teamCell = DashBoardSheetHeader.createCell(0);
        teamCell.setCellValue("Team");
        teamCell.setCellStyle(headerStyle);

        Cell countOfTicketNoCell = DashBoardSheetHeader.createCell(1);
        countOfTicketNoCell.setCellValue("Count of Ticket No");
        countOfTicketNoCell.setCellStyle(headerStyle);

        Cell bankNameCell = DashBoardSheetHeader.createCell(3);
        bankNameCell.setCellValue("BankName");
        bankNameCell.setCellStyle(headerStyle);

        Cell countOfBankTicketNoCell = DashBoardSheetHeader.createCell(4);
        countOfBankTicketNoCell.setCellValue("Count of Ticket No");
        countOfBankTicketNoCell.setCellStyle(headerStyle);

        Cell ticketAgeCell = DashBoardSheetHeader.createCell(6);
        ticketAgeCell.setCellValue("Ticket Age");
        ticketAgeCell.setCellStyle(headerStyle);

        Cell totalTicketCell = DashBoardSheetHeader.createCell(7);
        totalTicketCell.setCellValue("Total Ticket");
        totalTicketCell.setCellStyle(headerStyle);


        int maxRows = Math.max(TEAM_NAME_TICKET_COUNT_MAP.size(), Math.max(BANK_NAME_ROWS_MAP.size(), TICKET_AGE_BOUND_DETAILS_BOUND_DETAILS_LIST.size()));

        for(int i=1; i<=maxRows+1; i++) {
            dashboardSheet.createRow(i);
        }

        int outRowNum = 1;
        int total = 0;

        for (Map.Entry<String, Integer> entry : TEAM_NAME_TICKET_COUNT_MAP.entrySet()) {
            Row teamNameTicketCountRow = dashboardSheet.getRow(outRowNum++);
            teamNameTicketCountRow.createCell(0).setCellValue(entry.getKey());
            teamNameTicketCountRow.createCell(1).setCellValue(entry.getValue());
            teamNameTicketCountRow.getCell(0).setCellStyle(cellStyle);
            teamNameTicketCountRow.getCell(1).setCellStyle(cellStyle);
            total = total + entry.getValue();
        }

        Row teamNameTicketCountTotalRow = dashboardSheet.getRow(outRowNum++);
        teamNameTicketCountTotalRow.createCell(0).setCellValue("Grand Total");
        teamNameTicketCountTotalRow.createCell(1).setCellValue(total);
        teamNameTicketCountTotalRow.getCell(0).setCellStyle(headerStyle);
        teamNameTicketCountTotalRow.getCell(1).setCellStyle(headerStyle);

        outRowNum = 1;
        total = 0;
        for (Map.Entry<String, List<Row>> entry : BANK_NAME_ROWS_MAP.entrySet()) {
            Row bankNameTicketCountRow = dashboardSheet.getRow(outRowNum++);
            bankNameTicketCountRow.createCell(3).setCellValue(entry.getKey());
            bankNameTicketCountRow.createCell(4).setCellValue(entry.getValue().size());
            bankNameTicketCountRow.getCell(3).setCellStyle(cellStyle);
            bankNameTicketCountRow.getCell(4).setCellStyle(cellStyle);
            total = total + entry.getValue().size();
        }

        Row bankNameTicketCountTotalRow = dashboardSheet.getRow(outRowNum++);
        bankNameTicketCountTotalRow.createCell(3).setCellValue("Grand Total");
        bankNameTicketCountTotalRow.createCell(4).setCellValue(total);
        bankNameTicketCountTotalRow.getCell(3).setCellStyle(headerStyle);
        bankNameTicketCountTotalRow.getCell(4).setCellStyle(headerStyle);

        outRowNum = 1;
        total = 0;
        for (TicketAgeBoundDetails ticketAgeBoundDetails : TICKET_AGE_BOUND_DETAILS_BOUND_DETAILS_LIST) {
            Row ticketAgeCountRow = dashboardSheet.getRow(outRowNum++);
            ticketAgeCountRow.createCell(6).setCellValue(ticketAgeBoundDetails.getColumnName());
            ticketAgeCountRow.createCell(7).setCellValue(ticketAgeBoundDetails.getCount());
            ticketAgeCountRow.getCell(6).setCellStyle(cellStyle);
            ticketAgeCountRow.getCell(7).setCellStyle(cellStyle);
            total = total + ticketAgeBoundDetails.getCount();
        }


        Row grandTotalRow = dashboardSheet.getRow(outRowNum++);
        grandTotalRow.createCell(6).setCellValue("Grand Total");
        grandTotalRow.createCell(7).setCellValue(total);
        grandTotalRow.getCell(6).setCellStyle(headerStyle);
        grandTotalRow.getCell(7).setCellStyle(headerStyle);
    }
}

