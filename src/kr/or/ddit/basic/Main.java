package kr.or.ddit.basic;

import java.util.List;

public class Main {

	public static void main(String[] args) {
		String filePath = "D:/excelTest/grade.xlsx";
		List<List<String>> readList = XlsxUtils.readToList(filePath);
		
		for(int i=0; i<readList.size(); i++) {
			for(int j=0; j<readList.get(i).size(); j++) {
				System.out.print(readList.get(i).get(j) + "	");
			}
			System.out.println();
		}
	}
}

//public class Main {
//
//	public static void main(String[] args) {
//		String filePath = "C:/test/ttt.xlsx";
//		List<List<String>> readList = XlsxUtils.readToList(filePath);
//		
//		readList.forEach(row->{
//			row.forEach(cell->{
//				System.out.print(cell+", ");
//			});
//			System.out.println();
//		});
//	}
//}