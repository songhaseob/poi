package kr.or.ddit.basic;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	
public class XlsxUtils {
	
	@SuppressWarnings("resource")
	public static List<List<String>> readToList(String path){
		List<List<String>> list = new ArrayList<List<String>>();
		
		try {
			FileInputStream fi = new FileInputStream(path);//자바에서 파일을 바이트 단위로 입력하기위한 클래스
			XSSFWorkbook workbook = new XSSFWorkbook(fi);//통합 문서를 읽거나 쓸 때 생성 할 개체입니다. 
			XSSFSheet sheet = workbook.getSheetAt(0); //불러올 시트지정
			
			for(int i=0; i<sheet.getLastRowNum(); i++) {
				XSSFRow row = sheet.getRow(i);// 시트에 대한 행을 하나씩 추출
				if(row != null) {
					List<String> cellList = new ArrayList<String>();
					for(int j=0; j<row.getLastCellNum(); j++) {
						XSSFCell cell = row.getCell(j); // 행에대한 셀을 하나씩 추출하여 셀 타입에 따라 처리
						if(cell != null) {
							cellList.add( cellReader(cell) ); //셀을 읽어와서 List에 추가
						}
					}
					list.add(cellList); // 추가된 로우List를 List에 추가
				}
			}
		}catch(FileNotFoundException e) {
			e.printStackTrace();
		}catch(IOException e) {
			e.printStackTrace();
		}
		
		return list;
	}
	
	
	
	private static String cellReader(XSSFCell cell) {// 셀타입 처리부분
		String value = "";
		CellType ct = cell.getCellTypeEnum();
		if(ct != null) {
			switch(cell.getCellTypeEnum()) {
			case FORMULA://수식자체를 가져올때 String 반환
				value = cell.getCellFormula();
				break;
			case NUMERIC:
			    value=cell.getNumericCellValue()+"";
			    break;// 숫자 타입일때 double 반환
			case STRING://String으로 반환
			    value=cell.getStringCellValue()+"";
			    break;
			case BOOLEAN://boolean으로 반환
			    value=cell.getBooleanCellValue()+"";
			    break;
			case ERROR://에러나 났을경우 byte로 반환
			    value=cell.getErrorCellValue()+"";
			    break;
			}
		}
		return value; 
	}
}