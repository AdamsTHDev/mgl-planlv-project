package com.adms.mglplanreport.enums;

public enum ETemplateWB {

	MGL_SUMMARY_TEMPLATE("MGL Summary", "template/MGL-Summary.xlsx", 0),
	MGL_BY_CAMPAIGN_TEMPLATE("MGL By Campaign", "template/MGL-Summary.xlsx", 1),
	MGL_BY_CAMPAIGN_TOTAL_TEMPLATE("MGL By Campaign Total", "template/MGL-Summary.xlsx", 2),
	PLAN_LV_TEMPLATE("Plan LV Template", "template/MGL-Summary.xlsx", 3),
	MGL_CAMPAIGN_NAME_TEMPLATE("Campaign Name for MGL", "template/MGL-Summary.xlsx", 4);
	
	private String name;
	private String filePath;
	private int sheetIndex;
	
	private ETemplateWB(String name, String filePath, int sheetIndex) {
		this.name = name;
		this.filePath = filePath;
		this.sheetIndex = sheetIndex;
	}

	public String getName() {
		return name;
	}

	public String getFilePath() {
		return filePath;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}
	
}
