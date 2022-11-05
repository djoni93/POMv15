package excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelSupport {

    String prefixPath = "src/test/test_data/";

    public Map<String, String> getDataByRow(String fileName, String sheetName, String rowNum) throws IOException {

        FileInputStream fis = new FileInputStream("src/test/test_data/"+fileName+".xlsx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);

        int row = Integer.parseInt(rowNum);

        int lastColumnNum = sheet.getRow(1).getLastCellNum();

        Map<String, String> data = new HashMap<>();

        for(int i = 0; i < lastColumnNum; i++){
            data.put(
                    sheet.getRow(1).getCell(i).getStringCellValue(),
                    sheet.getRow((row+1)).getCell(i).getStringCellValue()
            );
        }

        return data;

    }

    public Map<String, String> getDataByID(String fileName, String sheetName, String testID) throws IOException {
        FileInputStream fis = new FileInputStream(prefixPath+fileName+".xslx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet(sheetName);

        int lastColumnNum = sheet.getRow(1).getLastCellNum();
        int lasRowNum = sheet.getLastRowNum();

        Map<String, String> data = new HashMap<>();

        for (int i = 0; i<lasRowNum;i++){

            if(sheet.getRow(i).getCell(0).getStringCellValue().equalsIgnoreCase(testID)){
                for(int j = 0; j<lastColumnNum; j++){
                    data.put(
                            sheet.getRow(1).getCell(i).getStringCellValue(),
                            sheet.getRow((i)).getCell(i).getStringCellValue()
                    );
                }
            }
        }








        return data;

    }
}
