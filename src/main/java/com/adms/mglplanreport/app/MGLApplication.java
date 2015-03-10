package com.adms.mglplanreport.app;

import com.adms.mglplanreport.service.MGLSummaryReport;
import com.adms.utils.DateUtil;

public class MGLApplication {

	public static void main(String[] args) {
		String dir = "D:/project/reports/MGL/out";
		try {
			System.out.println("START MGL Summary Report");
			
			MGLSummaryReport mglSummaryReport = new MGLSummaryReport();
			mglSummaryReport.generateReport(dir, DateUtil.convStringToDate("yyyyMMdd", "20150131"));
			
			System.out.println("END");
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
