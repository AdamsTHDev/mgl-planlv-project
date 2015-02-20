package com.adms.mglplanreport.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;


public class WorkbookUtil extends org.apache.poi.ss.util.WorkbookUtil {

	private static WorkbookUtil instance;
	
	public static WorkbookUtil getInstance() {
		if(instance == null) {
			instance = new WorkbookUtil();
		}
		return instance;
	}
	
	public void addPicture(Sheet sheet, byte[] bytes, int rowIdx, int colIdx, int pictureType) {
		addPicture(sheet, bytes, rowIdx, colIdx, 1, pictureType);
	}
	
	public void addPicture(Sheet sheet, byte[] bytes, int rowIdx, int colIdx, int resizePercent, int pictureType) {
		int pictureIdx = sheet.getWorkbook().addPicture(bytes, pictureType);
		
		CreationHelper helper = sheet.getWorkbook().getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setRow1(rowIdx);
		anchor.setCol1(colIdx);
		
		Picture picture = drawing.createPicture(anchor, pictureIdx);
		picture.resize(resizePercent);
	}
	
	public void copyCellValue(Cell origCell, Cell toCell) {
		switch(origCell.getCellType()) {
		case Cell.CELL_TYPE_BLANK :
			toCell.setCellValue(origCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_BOOLEAN :
			toCell.setCellValue(origCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR :
			toCell.setCellValue(origCell.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA :
			toCell.setCellValue(origCell.getCellFormula());
			break;
		case Cell.CELL_TYPE_NUMERIC :
			toCell.setCellValue(origCell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING :
			toCell.setCellValue(origCell.getRichStringCellValue());
			break;
		}
	}
	
	public void copyColumnWidth(Sheet origSheet, int origColumnNum, Sheet toSheet, int toColumnNum) {
		toSheet.setColumnWidth(toColumnNum, origSheet.getColumnWidth(origColumnNum));
	}
}
