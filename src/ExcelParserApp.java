import java.io.*;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.hpsf.Array;
import org.apache.poi.hpsf.Date;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelParserApp {

    public static void main(String[] args) throws FileFormatException, IOException {

        String inputFileName = "";
        for (int i = 0; i < args.length; i++) {
            if (i == 0) {
                inputFileName = inputFileName.concat(args[i]);

            } else {
                inputFileName = inputFileName.concat(" " + args[i]);
            }
        }
        if (inputFileName.split(",")[0].contains(".xlsx")) {
            String inputFilePath = inputFileName.split(",")[0];
            String outputFileName = inputFileName.split(",")[1].replace(" ", "");

            File excel = new File(inputFilePath);
            FileInputStream fis = new FileInputStream(excel);

            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheetAt(0);

            int rowNum = ws.getLastRowNum();
            int colNum = ws.getRow(0).getLastCellNum();

            ArrayList<ArrayList<String>> rowData = new ArrayList<ArrayList<String>>();
            int dupIE=0,dupCO=0,dupINC = 0;
            String colValStr = "";
            Iterator<Row> rowIterator = ws.rowIterator();
            //rowIterator.next();
            while (rowIterator.hasNext()) {
                XSSFRow row = (XSSFRow) rowIterator.next();
                // Now let's iterate over the columns of the current row
                Iterator<Cell> cellIterator = row.cellIterator();
                int count = 0;
                ArrayList<String> list = new ArrayList<String>();
                while (cellIterator.hasNext()) {

                    XSSFCell cell = (XSSFCell) cellIterator.next();
                    switch (cell.getCellType()) {
                        //case STRING: colValStr = cell.getStringCellValue(); break;
                        case STRING:
                            colValStr = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            LocalDateTime date = cell.getLocalDateTimeCellValue();
                            colValStr = date.toString().replace("T", " ");
                            //list.add(colValStr);
                            break;
                    }
                    list.add(colValStr);

                }
                if (list.contains("Infrastructure Event")) {
                    dupIE++;
                    list.clear();

                }
                if (list.contains("Customer Owned")) {
                    dupCO++;
                    list.clear();

                }
                if (!list.isEmpty()) {
                    rowData.add(list);
                }
            }

            int countRecord = 0;
            Map<String, ArrayList<String>> filterData = new LinkedHashMap<String, ArrayList<String>>();
            for (ArrayList list : rowData) {
                if (!filterData.containsKey(list.get(2))) {
                    filterData.put((String) list.get(2), list);
                    countRecord++;
                }else{
                    dupINC++;
                }
            }

            int rowNumber = 0;
            // createnew workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            // Create a blank sheet
            XSSFSheet sheet = workbook.createSheet("Incs With Spfc Worklog Entries");
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            Font font = workbook.createFont();

            Set<String> keySet = filterData.keySet();
            for (String key : keySet) {
                Row row = sheet.createRow(rowNumber++);
                List<String> nameList = filterData.get(key);
                if(nameList.contains("Count:")) {
                    rowNum -= 1;
                    countRecord -= 1;
                    continue;
                }
                int cellNumber = 0;
                for (String s : nameList) {
                    Cell cell = row.createCell(cellNumber++);

                    cell.setCellValue(s);
                }
                if(rowNumber == 1){
                    font.setBold(true);
                    cellStyle.setAlignment(HorizontalAlignment.LEFT);
                    cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(96, 200, 225, 255)));
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cellStyle.setFont(font);
                    row.setHeight((short) 700);
                    for(int j=0;j<colNum;j++)
                        row.getCell(j).setCellStyle(cellStyle);
                }
            }
            for (int i = 0; i < 15; i++) {
                sheet.autoSizeColumn(i);
            }
            System.out.println("Total Number Of Records in the Source File "+ (rowNum - 1));
            System.out.println("Number of Infrastructure Event Removed "+ dupIE);
            System.out.println("Number of Customer Owned Removed "+ dupCO);
            System.out.println("Number of Duplicate Incident ID's Removed "+dupINC);
            System.out.println("Number of Unique Records in Output File:" + (countRecord-1));

            FileOutputStream fos = new FileOutputStream(outputFileName);
            workbook.write(fos);
            fos.close();
        } else {
            throw new FileFormatException("File formt must be Excel file");
        }
    }
}
