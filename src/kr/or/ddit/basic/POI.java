package kr.or.ddit.basic;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POI {

   public static void main(String[] args) {
      
      ArrayList<String> name = new ArrayList<>();
      ArrayList<Integer> kor = new ArrayList<>();
      ArrayList<Integer> eng = new ArrayList<>();
      ArrayList<Integer> math = new ArrayList<>();
      ArrayList<Integer> sci = new ArrayList<>();
      
      name.add("홍길동"); name.add("홍길순"); name.add("박짱구"); name.add("김철수"); name.add("이유리");
      kor.add(80);   kor.add(92);   kor.add(96);   kor.add(77);   kor.add(84);
      eng.add(62);   eng.add(95);   eng.add(100);   eng.add(75);   eng.add(85);
      math.add(77);   math.add(88);   math.add(50);   math.add(95);   math.add(86);
      sci.add(93);   sci.add(96);   sci.add(66);   sci.add(88);   sci.add(87);
      
      ArrayList<Grade> list = new ArrayList<Grade>();
      for(int i=0; i<name.size(); i++) {
         Grade g = new Grade();
         g.setName(name.get(i));
         g.setKor(kor.get(i));
         g.setEng(eng.get(i));      
         g.setMath(math.get(i));
         g.setSci(sci.get(i));
         list.add(g);
      }
         
      XSSFWorkbook xlsWb = new XSSFWorkbook(); //xlsx 통합 문서를 읽거나 쓸 때 생성 할 개체입니다.
//    HSSFWorkbook xlsWb = new HSSFWorkbook(); //xls
      Sheet sheet1 = xlsWb.createSheet("성적표"); //시트명 설정
            
      Row row = null;
      Cell cell = null;
      
      int rowIdx = 0;
      
      row = sheet1.createRow(rowIdx++);
      String[] title = {"이름","국어","수학","영어","과학"};
      for(int i=0; i<title.length; i++) {
         cell = row.createCell(i);//title의 길이만큼 셀 생성
           cell.setCellValue(title[i]);//생성된 셀에 title의 i번째 값 입력
           cell.setCellStyle(cellStyle(xlsWb, "head")); // 셀 스타일 적용
      }

      Iterator<Grade> it = list.iterator();
      //Iterator타입에 변수를 생성하고 컬렉션마다에 iterator메서드를 값으로 넣습니다.
      //List타입에 변수를 Iterator으로 변환합니다.
           
      //hasNext메서드는 Iterator에 메서드입니다.
      //it변수에 다음데이터가 없을 때 까지 실행합니다.
      while(it.hasNext()) {
    	  //Iterator 안에 다음 값이 들어있는지 확인
    	  //들었으면 true, 안들었음 false
    	  
         Grade grd = it.next();//next메서드는 데이터를 반환합니다.
         
         
         row = sheet1.createRow(rowIdx++);//반복이 돌면 아래줄로 가서 입력
         int cellIdx = 0;//몇번째 셀에 넣을것인지 지정 ■ ■ > ㅁ ㅁ
         
         //data 출력하기
           cell = row.createCell(cellIdx++);//값 입력할때마다 옆으로 이동
           cell.setCellValue(grd.getName());//이름값
           cell.setCellStyle(cellStyle(xlsWb, "data")); // 셀 스타일 적용

           cell = row.createCell(cellIdx++);
           cell.setCellValue(grd.getKor());//국어점수값
           cell.setCellStyle(cellStyle(xlsWb, "data")); // 셀 스타일 적용
           
           cell = row.createCell(cellIdx++);
           cell.setCellValue(grd.getEng());//영어점수값
           cell.setCellStyle(cellStyle(xlsWb, "data")); // 셀 스타일 적용
           
           cell = row.createCell(cellIdx++);
           cell.setCellValue(grd.getMath());//수학점수값
           cell.setCellStyle(cellStyle(xlsWb, "data")); // 셀 스타일 적용
           
           cell = row.createCell(cellIdx++);
           cell.setCellValue(grd.getSci());//과학점수값
           cell.setCellStyle(cellStyle(xlsWb, "data")); // 셀 스타일 적용
      }
      
 

        try {// excel 파일 저장
           String path = "D:/excelTest/"; //경로
           String fileName = "grade.xlsx"; //파일명
            File xlsFile = new File(path+fileName); //저장경로 설정
            FileOutputStream fileOut = new FileOutputStream(xlsFile);
            xlsWb.write(fileOut);//위에서 만든 개체로 엑셀파일 쓰기
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
      
      
      
   }
   //셀 스타일 설정하는 함수
   public static CellStyle cellStyle(XSSFWorkbook xlsWb, String kind) {
      CellStyle cellStyle = xlsWb.createCellStyle();//셀스타일 적용할 대상
      cellStyle.setAlignment(HorizontalAlignment.CENTER); //가운데 정렬
      cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //중앙 정렬
      
      if(kind.equals("head")) {
         cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex()); //노란색
         cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //색상 패턴처리
      }else if(kind.equals("data")) {
         cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex()); //회색 25%
         cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //색상 패턴처리
      }
       
      return cellStyle;
   }
   
}