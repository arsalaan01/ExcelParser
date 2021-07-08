import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.hpsf.Array;
import org.apache.poi.hpsf.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelParserApp {
    public static void main(String[] args) throws Exception {

        //   File excel = new File("E:/VocalinkExcelSheet/Writesheet.xlsx");
        File excel = new File("C:\\Users\\196599\\Desktop\\NTT Team Members - Incidents Aug 2020.xlsx");
        FileInputStream fis = new FileInputStream(excel);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet ws = wb.getSheetAt(0);

        int rowNum = ws.getLastRowNum() + 1;
        int colNum = ws.getRow(0).getLastCellNum();

        ArrayList<ArrayList<String>> rowData = new ArrayList<ArrayList<String>>();

        String colValStr = "";
        Iterator<Row> rowIterator = ws.rowIterator();
        //rowIterator.next();
        while (rowIterator.hasNext()) {
            XSSFRow row = (XSSFRow) rowIterator.next();
            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();
            int count = -1;
            ArrayList<String> list = new ArrayList<String>();
            while (cellIterator.hasNext()) {
                count++;
                XSSFCell cell = (XSSFCell) cellIterator.next();
                switch (cell.getCellType()) {
                    //case STRING: colValStr = cell.getStringCellValue(); break;
                    case STRING:
                        colValStr = cell.getStringCellValue();
                        break;
                    case NUMERIC:
                        LocalDateTime date = cell.getLocalDateTimeCellValue();
                        colValStr = date.toString().replace("T", " ");
                        list.add(colValStr);
                        break;
                }
                list.add(colValStr);

            }
            if (list.contains("Infrastructure Event")
                    || list.contains("Customer Owned")) {
                list.clear();
            }
            if (!list.isEmpty()) {
                rowData.add(list);
            }
        }
        int countRecord = 0;
        Map<String, ArrayList<String>> filterData = new LinkedHashMap<String, ArrayList<String>>();
        ;
        for (ArrayList list : rowData) {
            if (!filterData.containsKey(list.get(2))) {
                filterData.put((String) list.get(2), list);
                countRecord++;
            }
        }
//        System.out.println(filterData);
//        System.out.println(countRecord);

//        Set<String> keySet = filterData.keySet();
//        for(String colData:keySet){
//            List<String> nameList = filterData.get(colData);
//            for(String s : nameList){
//
//            }
//        }



        int r1=0,c1=0;
        // createnew workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        Set<String> keySet = filterData.keySet();
        for(String colData:keySet){
            XSSFRow row = sheet.createRow(r1);
            List<String> nameList = filterData.get(colData);
            c1 =0;
            for(String s : nameList){
                XSSFCell cell = row.createCell(c1);
                c1++;
            }
            r1++;
        }


        FileOutputStream fos = new FileOutputStream(new File("E:/VocalinkExcelSheet/Writesheet3.xlsx"));
        workbook.write(fos);
        fos.close();
    }
    }
}

// iterating over a map
//                // iterating over a map
//                for(Map.Entry<String, ArrayList<String>> listEntry : filterData.entrySet()){
//                    //System.out.println("Iterating list number - " + ++i);
//                    // iterating over a list
//                    for(String colValue : listEntry.getValue()){
//                        System.out.println(listEntry.getKey() +"                "+ colValue);
//                    }
//                }


// for (int i = 0 ; i <= rowNum; i++) {
//         XSSFRow row = sheet.createRow(0);
//         for (int c =0; c < colNum; c++){
//        XSSFCell cell = row.createCell(c);
//        Set<String> keySet = filterData.keySet();
//        for(String colData:keySet){
//        List<String> nameList = filterData.get(colData);
//        for(String s : nameList){
//        cell.setCellValue(s);
//        }
//        }
//        }