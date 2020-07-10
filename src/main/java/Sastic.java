import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class Sastic {
    /* 파일 파싱(?)하기 전에 변경해야 하는 변수 */
    public static final String NEST_OS = "linux"; // windows, linux
    public static final String CLIENT_OS = "linux"; // windows, linux
    public static final String BEFORE_FILE_NAME = "issues_linuxServer_linuxClient_centos_objc_1838.xls";

    /* 최초 변경해야 하는 변수 */
    // 파싱(?)할 파일이 존재하는 경로
    public static final String BEFORE_FILE_PATH = "C:\\Users\\kimdain\\Downloads\\engine_issue\\"+BEFORE_FILE_NAME;
    // 파싱(?)된 파일을 저장할 경로
    public static final String AFTER_FILE_PATH = "C:\\Users\\kimdain\\Desktop\\incident\\sast\\engine_issue\\"+NEST_OS+"_"+CLIENT_OS;
    // 윈도우 클라이언트로 분석 했을 시 앞부분 제외할 경로
    public static final String WINDOWS_EXCLUDE_PATH = "C:\\Users\\kimdain\\Desktop\\testcode\\2006testcodeUtill\\";
    // 리눅스 클라이언트로 분석 했을 시 앞부분 제외할 경로
    public static final String LINUX_EXCLUDE_PATH = "/home/dain/testcode/2006testcodeUtill/";

    enum SASTCOLUMN{
        CHECKER(5), LINE(6), PATH(9);

        int index;

        SASTCOLUMN (int index){
            this.index  = index;
        }

    }

    public static void main(String[] args) throws IOException {
        String filepath = BEFORE_FILE_PATH;
        FileInputStream inputStream = new FileInputStream(filepath);
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream); // 액셀 읽기
        HSSFSheet sheet = workbook.getSheetAt(0); // 시트가져오기 0은 첫번째 시트

        // 컬럼 삭제
        deleteColumn(sheet);
        cutPath(sheet);

        // 파일 다운로드
        filedown(workbook);

    }

    public static void deleteColumn(HSSFSheet sheet){
        Iterator<Row> rowIterator = sheet.iterator();

        while(rowIterator.hasNext()){
            HSSFRow row = (HSSFRow)rowIterator.next();
            for(int i=0 ; i<17 ; i++){
                if(i != SASTCOLUMN.CHECKER.index && i != SASTCOLUMN.LINE.index && i != SASTCOLUMN.PATH.index) {
                    HSSFCell cell = row.getCell(i);
                    row.removeCell(cell);
                }
            }
            row.shiftCellsLeft(SASTCOLUMN.CHECKER.index,SASTCOLUMN.LINE.index,5);
            row.shiftCellsLeft(SASTCOLUMN.PATH.index, SASTCOLUMN.PATH.index,7);
        }
    }

    public static void cutPath(HSSFSheet sheet){
        Iterator<Row> rowIterator = sheet.iterator();
        // 첫째 행 버리기
        rowIterator.next();
        int osIndex = 0;
        boolean osWindows = false;

        if(CLIENT_OS.equals("windows")) {
            osIndex = WINDOWS_EXCLUDE_PATH.length();
            osWindows = true;
        } else osIndex = LINUX_EXCLUDE_PATH.length();

        while (rowIterator.hasNext()) {
            HSSFRow row = (HSSFRow)rowIterator.next();
            HSSFCell cell = row.getCell(2);
            String path = cell.getStringCellValue();
            // 경로 자르기
            path = path.substring(osIndex);
            // 윈도우 경로라면 /로 변환
            if(osWindows)
                path = path.replace("\\","/");
            cell.setCellValue(path);
        }

    }

    public static void filedown (HSSFWorkbook workbook){
        try {
            SimpleDateFormat formatter = new SimpleDateFormat("mmss");
            Date nowdate = new Date();
            String dateString = " "+formatter.format(nowdate);
            FileOutputStream fileoutputstream = new FileOutputStream(AFTER_FILE_PATH+dateString+".xls");
            workbook.write(fileoutputstream);
            fileoutputstream.close();
            System.out.println("엑셀파일생성성공");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("엑셀파일생성실패");
        }
    }
}
