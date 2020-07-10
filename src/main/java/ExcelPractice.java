import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelPractice {
    public static void main(String[] args) throws IOException {
        String filepath = "C:\\Exception\\exceltest.xlsx";
        FileInputStream inputStream = new FileInputStream(filepath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream); // 액셀 읽기
        XSSFSheet sheet = workbook.getSheetAt(0); // 시트가져오기 0은 첫번째 시트
        int rows = sheet.getPhysicalNumberOfRows(); // 시트에서 총 행수
        XSSFCell cell00 = sheet.getRow(0).getCell(0);
        XSSFCell cell01 = sheet.getRow(0).getCell(1);
        XSSFCell cell10 = sheet.getRow(1).getCell(0);

        double value00 = 0;
        double value01 = 0;
        String value10 = "";
        value00 = cell00.getNumericCellValue();
        value01 = cell01.getNumericCellValue();
        value10 = cell10.getStringCellValue();

        System.out.println(value00);
        System.out.println(value01);
        System.out.println(value10);
    }
}
