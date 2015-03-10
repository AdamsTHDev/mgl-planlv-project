package com.adms.mglplanreport.app;

import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.net.URLClassLoader;

import com.adms.imex.excelformat.DataHolder;
import com.adms.imex.excelformat.ExcelFormat;
import com.adms.mglplanlv.entity.Campaign;
import com.adms.mglplanlv.entity.MglTarget;
import com.adms.mglplanlv.service.campaign.CampaignService;
import com.adms.mglplanlv.service.mgltarget.MglTargetService;
import com.adms.mglplanreport.util.ApplicationContextUtil;

public class ImportMGLTargetApp {

	public static void main(String[] args) {
		
		InputStream fileFormat = null;
		InputStream excel = null;
		try {
			fileFormat = URLClassLoader.getSystemResourceAsStream("fileformat/MGL_Target_Format.xml");
			excel = new FileInputStream("D:/project/reports/MGL/in/Summary_MGL_Target.xlsx");
			
			ExcelFormat ex = new ExcelFormat(fileFormat);
			DataHolder fileDataHolder = ex.readExcel(excel);
			
			DataHolder sheet = fileDataHolder.get(fileDataHolder.getKeyList().get(0));
			CampaignService campaignService = (CampaignService) ApplicationContextUtil.getApplicationContext().getBean("campaignService");
			MglTargetService mglTargetService = (MglTargetService) ApplicationContextUtil.getApplicationContext().getBean("mglTargetService");
			
			for(DataHolder data : sheet.getDataList("dataList")) {
				String campaignCode = data.get("campaignCode").getStringValue();
				BigDecimal issuedRate = data.get("issuedRate").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP);
				BigDecimal paidRate = data.get("paidRate").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP);
				String targetYear = data.get("targetYear").getStringValue();
				
				MglTarget mgl = new MglTarget();
				mgl.setCampaign(campaignService.find(new Campaign(campaignCode)).get(0));
				mgl.setIssuedRate(issuedRate);
				mgl.setPaidRate(paidRate);
				mgl.setTargetYear(targetYear);
				mgl = mglTargetService.add(mgl, "Import System");
				System.out.println("ADDED: " + mgl.toString());
			}
			
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}
}
