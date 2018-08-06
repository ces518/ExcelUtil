package util;

import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import egovframework.edosi.view.EdosiExcelVO;

public class ExcelUtil {
	
	protected void buildExcelDocument(Map model, HSSFWorkbook workbook,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		
		
        //파라미터 선언 및 초기화
		List<ExcelVO> list =  (List)model.get("test");

        //다운로드를 위하여 헤더 Setting
		response.setHeader("Content-Disposition",  "attachment;fileName=\"" + "파일명" + ".xls\";");
		
        HSSFSheet sheet = workbook.createSheet("예약통계");
        int row = 0;
        sheet.setDefaultColumnWidth((short)12);

        //데이터 입력
        row = this.setFirstCell(row, sheet);
        Iterator<ExcelVO> it = list.iterator();
        ExcelVO bean = null;
        while(it.hasNext()){
        	bean = it.next();
        	row = this.setCell(row, sheet, bean);
        }
	}

	/**
	 * 제목셀을 설정한다.
	 * @param row
	 * @param sheet
	 * @return
	 */
	private int setFirstCell(int row, HSSFSheet sheet){
		HSSFCell cell;
		
		//일자
		cell = getCell(sheet, row, 0);
		setText(cell, "일자");
		
		//예약건수
		cell = getCell(sheet, row, 1);
		setText(cell, "예약건수");
		
		//취소건수
		cell = getCell(sheet, row, 2);
		setText(cell, "취소건수");
		
		//완료건수
		cell = getCell(sheet, row, 3);
		setText(cell, "완료건수");
		
		//유효예약률
		cell = getCell(sheet, row, 4);
		setText(cell, "유효예약률");
		
		return ++row;
	}
	
	/**
	 * 셀에 내용을 입력한다.
	 * @param row
	 * @param sheet
	 * @param employeeBean
	 * @return
	 */
	private int setCell(int row, HSSFSheet sheet, ExcelVO bean){
		HSSFCell cell;
		
		//이름
		cell = getCell(sheet, row, 0);
		setText(cell, bean.getName());
		
		//주소
		cell = getCell(sheet, row, 1);
		setText(cell, bean.getAddr());
		
		return ++row;
	}
}
