package com.adms.mglplanreport.service;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.adms.mglplanreport.enums.ETemplateWB;
import com.adms.mglplanreport.util.WorkbookUtil;

public class MGLByCampaign {
	
	private final int C_COLUMN_NUM = 2;
	private final int I_COLUMN_NUM = 8;
	
	private final int START_TABLE_HEADER_ROW = 8;
	private final int START_TABLE_DATA_ROW = 9;
	private final int TEMP_TABLE_TOTAL_KEYCODE_ROW = 12;
	private final int TEMP_TABLE_MONTH_TO_DATE_ROW = 13;

	public void generateReport(String outPath, Date dataDate) {
		
		try {
			//Template
			Workbook wb = WorkbookFactory.create(Thread.currentThread().getContextClassLoader().getResourceAsStream(ETemplateWB.MGL_BY_CAMPAIGN_TEMPLATE.getFilePath()));
			Sheet tempSheet = wb.getSheetAt(ETemplateWB.MGL_BY_CAMPAIGN_TEMPLATE.getSheetIndex());
			Sheet toSheet = wb.createSheet("MGL by Campaign");
			
//			receive All data
//			bla bla
			
//			set Grid blank
			toSheet.setDisplayGridlines(false);
			
//			set header
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void setHeader(Sheet tempSheet, Sheet toSheet) throws Exception {
		
		Cell tempCell = null;
		Cell toCell = null;
		
		for(int n = 0; n < 8; n++) {
			for(int c = 0; c < tempSheet.getRow(n).getLastCellNum(); c++) {
				tempCell = tempSheet.getRow(n).getCell(c, Row.CREATE_NULL_AS_BLANK);
				toCell = toSheet.createRow(n).createCell(n, tempCell.getCellType());

				switch(n) {
				case 2 :
					if(c == C_COLUMN_NUM) {
						setHeaderCellValue(toCell, "CAMPAIGN_CODE");
					} else if(c == I_COLUMN_NUM) {
						setHeaderCellValue(toCell, "Calling Date");
					}
				break;
				case 3 :
					if(c == C_COLUMN_NUM) {
						setHeaderCellValue(toCell, "KEY_CODE");
					} else if(c == I_COLUMN_NUM) {
						setHeaderCellValue(toCell, "CALLING_SITE");
					}
				break;
				case 4 :
					if(c == C_COLUMN_NUM) {
						setHeaderCellValue(toCell, "RECORD_RECEIVED");
					}
				break;
				case 5 :
					if(c == C_COLUMN_NUM) {
						setHeaderCellValue(toCell, "PRINT_DATE");
					}
				break;
				default : WorkbookUtil.getInstance().copyCellValue(tempCell, toCell); break;
				}
				
			}
		}
	}
	
	private void setHeaderCellValue(Cell cell, Object value) throws Exception {
		WorkbookUtil.getInstance().setObjectValueToCell(cell, value);
	}
}
