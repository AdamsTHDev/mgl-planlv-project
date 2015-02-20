package com.adms.mglplanreport.service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;

import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.MGLSummaryObj;
import com.adms.utils.DateUtil;

public class MGLSummaryReport {
	
	private final int START_TABLE_HEADER_ROW = 7;
	private final int START_TABLE_DATA_ROW = 9;
	private final int TEMP_TABLE_TOTAL_ROW = 12;
	private final String MONTH_PATTERN = DateUtil.getDefaultMonthPattern();
	
	private Map<String, Double[]> sumOfMtdMap = new HashMap<>();
	private Double[] sumAllOfYTD = new Double[]{0D, 0D};
	
	public void generateReport(String outPath, Date dataDate) {

		try {
			//Template
			Workbook wb = WorkbookFactory.create(Thread.currentThread().getContextClassLoader().getResourceAsStream("template/MGL-Summary.xlsx"));
			Sheet tempSheet = wb.getSheetAt(0);
			Sheet sheet = wb.createSheet("MGL_SUM");
			
//			set Grid blank
			sheet.setDisplayGridlines(false);
			
//			set caption
			Cell captionCell = sheet.createRow(5).createCell(0, tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getCellType());
			captionCell.setCellStyle(tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getCellStyle());
			captionCell.setCellValue(tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getStringCellValue());
			
//			set table column header
			List<MGLSummaryObj> mglSumList = getMGLSummary();
			int noOfMonth = getNumberOfMonthInYear(mglSumList);
			doTableHeader(tempSheet, sheet, noOfMonth);
			
//			insert picture
			byte[] bytes = IOUtils.toByteArray(Thread.currentThread().getContextClassLoader().getResourceAsStream("template/ADAMS_logo_th.png"));
			WorkbookUtil.getInstance().addPicture(sheet, bytes, 0, 0, Workbook.PICTURE_TYPE_PNG);
			
//			set table data
			doTableData(tempSheet, sheet, noOfMonth, mglSumList);
			
//			set table total
			doTableTotal(tempSheet, sheet, noOfMonth);
			
//			remove sheet(s)
			wb.removeSheetAt(0);
			wb.removeSheetAt(0);

//			write out
			OutputStream os = new FileOutputStream(new File(outPath + "/test.xlsx"));
			wb.write(os);
			os.close();
			wb.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
		
	}

	private List<MGLSummaryObj> getMGLSummary() {
//		for Test
		List<MGLSummaryObj> list = new ArrayList<>();
		
		Map<String, Double[]> map = null;
		MGLSummaryObj obj = null;
		
		obj = new MGLSummaryObj();
		obj.setCampaignCode("CampaignCode1");
		obj.setCampaignName("CampaignName1");
		obj.setIssuedRate(100D);
		obj.setPaidRate(200D);
		obj.setIapMTD(300D);
		obj.setIapYTD(400D);
		map = new HashMap<>();
		map.put(DateUtil.getStringOfMonth(5), new Double[]{1D, 2D});
		map.put(DateUtil.getStringOfMonth(6), new Double[]{2D, 4D});
		map.put(DateUtil.getStringOfMonth(7), new Double[]{2D, 4D});
		map.put(DateUtil.getStringOfMonth(8), new Double[]{2D, 4D});
		map.put(DateUtil.getStringOfMonth(9), new Double[]{2D, 4D});
		map.put(DateUtil.getStringOfMonth(10), new Double[]{2D, 4D});
		map.put(DateUtil.getStringOfMonth(11), new Double[]{2D, 4D});
		obj.setMTD(map);
		list.add(obj);
		
		obj = new MGLSummaryObj();
		obj.setCampaignCode("CampaignCode2");
		obj.setCampaignName("CampaignName2");
		obj.setIssuedRate(100D);
		obj.setPaidRate(200D);
		obj.setIapMTD(300D);
		obj.setIapYTD(400D);
		map = new HashMap<>();
		map.put(DateUtil.getStringOfMonth(3), new Double[]{1D, 2D});
		map.put(DateUtil.getStringOfMonth(4), new Double[]{3D, 4D});
		map.put(DateUtil.getStringOfMonth(5), new Double[]{3D, 4D});
		map.put(DateUtil.getStringOfMonth(6), new Double[]{3D, 4D});
		map.put(DateUtil.getStringOfMonth(7), new Double[]{3D, 4D});
		map.put(DateUtil.getStringOfMonth(8), new Double[]{3D, 4D});
		obj.setMTD(map);
		list.add(obj);
		
		obj = new MGLSummaryObj();
		obj.setCampaignCode("CampaignCode3");
		obj.setCampaignName("CampaignName3");
		obj.setIssuedRate(100D);
		obj.setPaidRate(200D);
		obj.setIapMTD(300D);
		obj.setIapYTD(400D);
		map = new HashMap<>();
		map.put(DateUtil.getStringOfMonth(3), new Double[]{1D, 2D});
		map.put(DateUtil.getStringOfMonth(4), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(5), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(6), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(7), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(8), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(9), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(10), new Double[]{4D, 10D});
		map.put(DateUtil.getStringOfMonth(11), new Double[]{4D, 10D});
		obj.setMTD(map);
		list.add(obj);
//		
		
		
		return list;
	}
	
	private int getNumberOfMonthInYear(List<MGLSummaryObj> mglSumList) {
		int max = 0;
		for(MGLSummaryObj obj : mglSumList) {
			Map<String, Double[]> map = obj.getMTD();
			for(String key: map.keySet()) {
				int m = 0;
				try {
					m = DateUtil.getMonthNo(key);
				} catch(Exception e) {
					System.err.println("cannot convert: " + key);
					e.printStackTrace();
				}
				max = max < m ? max = new Integer(m) : max;
			}
		}
		return max > 0 ? max + 1 : 1;
	}
	
	private void doTableHeader(Sheet tempSheet, Sheet toSheet, int noOfMonth) throws Exception {
		
		int startRow = new Integer(START_TABLE_HEADER_ROW).intValue();
		int mtdIdx = noOfMonth * 2;
		
		for(int rn = startRow; rn < startRow + 2; rn++) {
			Row tempRow = tempSheet.getRow(rn);
			Row toRow = toSheet.createRow(rn);
			
			Cell tempCell = null;
			Cell toCell = null;
			
			int currMonth = 0;
			int maxCol = tempRow.getLastCellNum() + mtdIdx;
			for(int cn = 0; cn < maxCol; cn++) {
				
//				MTD
				if(cn > 1 && cn < (mtdIdx + 2)) {
					int temp = cn % 2 == 0 ? 2 : 3;
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					
					String mmm = toCell.getStringCellValue();
					if(mmm.indexOf(MONTH_PATTERN) > 0) {
						toCell.setCellValue(mmm.replace(MONTH_PATTERN, DateUtil.getStringOfMonth(currMonth)));
						currMonth++;
					}
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, temp, toSheet, cn);
					
					if(rn == startRow && cn % 2 != 0) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, startRow, cn - 1, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
				
//				after MTD
				} else if(((cn + 2) - mtdIdx) > 3) {
					int temp = (cn + 2) - mtdIdx;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, temp, toSheet, cn);
					
					if(rn == startRow && cn % 2 != 0 && temp < 6) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, startRow, cn - 1, cn);
						toSheet.addMergedRegion(mergedRegion);
					} else if(rn == (startRow + 1) && temp > 6) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, rn, cn, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
					
//				campaign
				} else {
					tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, cn, toSheet, cn);
					
					if(rn == (startRow + 1)) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, rn, cn, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
				}
			}
		}
	}
	
	private void doTableData(Sheet tempSheet, Sheet toSheet, int noOfMonth, List<MGLSummaryObj> mglSumList) throws Exception {
//		remark*: flow is same as table header
		
		int startRow = new Integer(START_TABLE_DATA_ROW).intValue();
		int n = 0;
		int mtdIdx = noOfMonth * 2;
		
		for(int rn = startRow; rn < startRow + mglSumList.size(); rn++) {
			MGLSummaryObj mgl = mglSumList.get(n);
			
			Double[] sumOfYtd = new Double[]{0D, 0D};
			
			Row tempRow = tempSheet.getRow(startRow);
			Row toRow = toSheet.createRow(rn);
			
			Cell tempCell = null;
			Cell toCell = null;
			
			int maxCol = tempRow.getLastCellNum() + mtdIdx;
			
			for(int cn = 0; cn < maxCol; cn++) {
				
//				MTD
				if(cn > 1 && cn < (mtdIdx + 2)) {
					boolean isFirstColOfMTD = cn % 2 == 0;
					int temp = isFirstColOfMTD ? 2 : 3;
					double val = 0D;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					
					toCell.setCellStyle(tempCell.getCellStyle());
					
					String mtdColMonth = null;
					String monthFromCell = null;
					if(isFirstColOfMTD) {
						monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
					} else {
						monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn - 1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
					}
					mtdColMonth = monthFromCell.substring(monthFromCell.indexOf("(") + 1, monthFromCell.indexOf(")"));
					
					Double[] mtdVal = mgl.getMTD().get(mtdColMonth);
					if(mtdVal != null && mtdVal.length > 0) {
						int idx = isFirstColOfMTD ? 0 : 1;
						val = mtdVal[idx].doubleValue();
						sumOfYtd[idx] += val;
						
						Double[] mtdByMMM = sumOfMtdMap.get(mtdColMonth);
						if(mtdByMMM == null) {
							mtdByMMM = new Double[]{0D, 0D};
						}
						mtdByMMM[idx] += val;
						sumOfMtdMap.put(mtdColMonth, mtdByMMM);
					}
					
					toCell.setCellValue(val);
					
//				after MTD
				} else if(((cn + 2) - mtdIdx) > 3) {
					int temp = (cn + 2) - mtdIdx;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					
//					YTD
					if(temp < 6) {
						int idx = cn % 2 == 0 ? 0 : 1;
						toCell.setCellValue(sumOfYtd[idx]);
						sumAllOfYTD[idx] += sumOfYtd[idx];
						
					} /* others */ else if(temp > 6) {
						WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					}
					
//				campaign
				} else {
					tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					
					toCell.setCellValue(cn % 2 == 0 ? mgl.getCampaignName() : mgl.getCampaignCode());
				}
			}
			n++;
		}
	}
	
	private void doTableTotal(Sheet tempSheet, Sheet toSheet, int noOfMonth) {
		int startRow = toSheet.getLastRowNum() + 1;
		int mtdIdx = noOfMonth * 2;
		
		Row tempRow = tempSheet.getRow(TEMP_TABLE_TOTAL_ROW);
		Row toRow = toSheet.createRow(startRow);
		
		Cell tempCell = null;
		Cell toCell = null;
		int maxCol = mtdIdx + tempRow.getLastCellNum();
		
		for(int cn = 0; cn < maxCol; cn++) {
			
//			MTD
			if(cn > 1 && cn < (mtdIdx + 2)) {
				boolean isFirstColOfMTD = cn % 2 == 0;
				int temp = isFirstColOfMTD ? 2 : 3;
				
				tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				
				toCell.setCellStyle(tempCell.getCellStyle());
				
				String mtdColMonth = null;
				String monthFromCell = null;
				if(isFirstColOfMTD) {
					monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
				} else {
					monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn - 1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
				}
				mtdColMonth = monthFromCell.substring(monthFromCell.indexOf("(") + 1, monthFromCell.indexOf(")"));
				
				Double[] mtd = sumOfMtdMap.get(mtdColMonth);
				if(mtd != null) {
					toCell.setCellValue(sumOfMtdMap.get(mtdColMonth)[isFirstColOfMTD ? 0 : 1]);
				} else {
					toCell.setCellValue(0);
				}
				
				
//			after MTD
			} else if(((cn + 2) - mtdIdx) > 3) {
				int temp = (cn + 2) - mtdIdx;
				
				tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				toCell.setCellStyle(tempCell.getCellStyle());
				
//				YTD
				if(temp < 6) {
					toCell.setCellValue(sumAllOfYTD[cn % 2 == 0 ? 0 : 1]);
					
				} /* others */ else if(temp > 6) {
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
				}
				
//			campaign
			} else {
				tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				toCell.setCellStyle(tempCell.getCellStyle());
				
				if(cn == 0) {
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
				} else {
					toSheet.addMergedRegion(new CellRangeAddress(startRow, startRow, cn - 1, cn));
				}
			}
		}
		
	}
}
