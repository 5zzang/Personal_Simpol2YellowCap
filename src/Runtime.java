import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.TimeZone;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import au.com.bytecode.opencsv.CSVReader;

/**
 * Simpol2YellowCap : 배송송장 변환 프로그램.
 * @author 5zzang
 */
public class Runtime {
	public static void main(String[] args) {
		System.out.println(">>> 심폴 배송CSV를 옐로우캡 택배 엑설로 변환시키는 프로그램. by 5zzang <<<");
		
		try {
			// 오늘 날짜 Simpol 배송 CSV 파일을 읽어온다.
			TimeZone zone = TimeZone.getTimeZone("GMT+9:00");
			Calendar cal = Calendar.getInstance(zone);
			SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
			String today_order = "order_deli_" + formatter.format(cal.getTime()) + ".csv";
			
			System.out.println(">>> 오늘 날짜 심폴 배송CSV 파일을 찾습니다...");
			FileInputStream today_fis = new FileInputStream(today_order);
			
			// 파일이 정상적으로 있는 경우..
			if ( today_fis.available() > 0 ) {
				System.out.println(">>> 오늘 날짜 심폴 배송CSV 파일을 찾았습니다.");
				
				// 엑셀 템플릿 파일을 로드한다.
				System.out.println(">>> 엑셀 템플릿 파일을 로드합니다...");
				String template = "./template.xls";
				POIFSFileSystem fs_template = new POIFSFileSystem(new FileInputStream(template));
				HSSFWorkbook wb_template = new HSSFWorkbook(fs_template);
				HSSFSheet sheet_template = wb_template.getSheetAt(0);
				System.out.println(">>> 엑셀 템플릿 파일이 로드가 완료 되었습니다.");
				
				// 배송 CSV 파일을 CSVReader로 읽는다.
				System.out.println(">>> 심폴 배송CSV 파일을 변환중입니다...");
				CSVReader csv_today = new CSVReader(new InputStreamReader(today_fis, "EUC-KR"));
				List<String[]> todayList = csv_today.readAll();
				csv_today.close();
				System.out.println(">>> 심폴 배송CSV 파일의 변환이 완료되었습니다.");
				
				// 엑셀 템플릿에 배송 CSV의 내용을 넣는다.
				System.out.println(">>> 싱폼 배송CSV의 값을 새로운 옐로우캡 택배 엑셀로 변환 하는중입니다...");
				for (int rowNo = 1; rowNo < todayList.size(); rowNo++) {
					String[] readLine = (String[]) todayList.get(rowNo);
					//if (i == 1) System.out.println(" >>> Column Size : " +readLine.length);
					
					HSSFRow row = sheet_template.createRow(rowNo);
					row.createCell(0).setCellValue(readLine[7]);	// 수화주
					row.createCell(1).setCellValue(readLine[10]);	// 우편번호
					row.createCell(2).setCellValue(readLine[11]);	// 주소
					row.createCell(3).setCellValue(readLine[8]);	// 전화번호
					row.createCell(4).setCellValue(readLine[9]);	// 핸드폰번호
					row.createCell(5).setCellValue(1);	// 택배수량
					row.createCell(6).setCellValue(2500);	// 택배금액
					row.createCell(7).setCellValue("003");	// 선착불
					row.createCell(8).setCellValue(readLine[22]);	// 상품명
					row.createCell(9).setCellValue(readLine[29]);	// 비고(택배메시지)
				}
				System.out.println(">>> 열로우캡 택배 엑셀로 변환이 완료되었습니다.");
				
				// today.xls 파일을 만든다.
				System.out.println(">>> 옐로우캡 택배 엑셀 : today.xls 를 생성중입니다...");
				FileOutputStream today_xls = new FileOutputStream("./today.xls");
				wb_template.write(today_xls);
				today_xls.close();	// 파일쓰기 종료
				System.out.println(">>> 옐로우캡 택배 엑셀 : today.xls 이 정상적으로 생성이 되었습니다.");
				System.out.println(">>> 심폴 배송CSV 파일을 옐로우캡 택배엑셀로 변환하는 작업이 모두 완료 되었습니다. <<<");	
			}
		} catch (FileNotFoundException e) {
			System.out.println(">>> 오늘 날짜 배송CSV 파일이 없습니다. <<<");
		} catch (IOException e) {
			System.out.println(">>> 예상되지 않은 파일 입출력 관련 에러 발생!!");
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println(">>> 변환 프로그램을 종료합니다. 감사합니다. <<<");
		System.out.println(">>> http://www.다육이와한지.com - 031.963.2967 <<<");
	}
}