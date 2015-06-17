package com.adms.mglplanreport.app;

import java.util.Date;

import com.adms.mglplanreport.service.MGLByCampaignMonthlyReport;
import com.adms.mglplanreport.service.MGLSummaryReport;
import com.adms.mglplanreport.service.PlanLVReport;
import com.adms.utils.DateUtil;

public class MGLApplication {

	public static void main(String[] args) {
		String dir = "D:/project/reports/MGL/out";
		try {
			String processDateStr = "20150531";
			Date processDate = DateUtil.convStringToDate("yyyyMMdd", processDateStr);
			
			new MGLSummaryReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/summary", processDate);
			
			new MGLByCampaignMonthlyReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/product", processDate);
			
			new PlanLVReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/planlv", processDate);
			
			System.out.println("END");
		} catch(Exception e) {
			e.printStackTrace();    
		}
	}
}
