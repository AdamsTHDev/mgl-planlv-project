package com.adms.mglplanreport.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.math.BigDecimal;
import java.net.URLClassLoader;
import java.text.DecimalFormat;
import java.util.Date;
import java.util.List;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import com.adms.imex.excelformat.DataHolder;
import com.adms.imex.excelformat.ExcelFormat;
import com.adms.mglplanlv.entity.ListLot;
import com.adms.mglplanlv.entity.ProductionByLot;
import com.adms.mglplanlv.service.listlot.ListLotService;
import com.adms.mglplanlv.service.productionbylot.ProductionByLotService;
import com.adms.support.FileWalker;
import com.adms.utils.Logger;

public class ImportProduction {
	
	public static final String BATCH_ID = "BATCH_ID";
	public static final String APPLICATION_CONTEXT_FILE = "config/application-context.xml";

	public static final String FILE_FORMAT_PRODUCTION_BY_LOT_TELE = "fileformat/FileFormat_SSIS_DailyProductionByLot-input-TELE.xml";
	public static final String FILE_FORMAT_PRODUCTION_BY_LOT_OTO = "fileformat/FileFormat_SSIS_DailyProductionByLot-input-OTO.xml";

	public static final String PRODUCTION_BY_LOT_SERVICE_BEAN = "productionByLotService";
	public static final String LIST_LOT_SERVICE_BEAN = "listLotService";

	private ApplicationContext applicationContext;
	protected Logger log = Logger.getLogger(Logger.DEBUG);

	protected void setLogLevel(int logLevel)
	{
		this.log.setLogLevel(logLevel);
	}


	protected Object getBean(String beanId)
	{
		if (this.applicationContext == null)
		{
			this.applicationContext = new ClassPathXmlApplicationContext(APPLICATION_CONTEXT_FILE);
		}

		return this.applicationContext.getBean(beanId);
	}

	protected ProductionByLotService getProductionByLotService()
	{
		return (ProductionByLotService) getBean(PRODUCTION_BY_LOT_SERVICE_BEAN);
	}

	protected ListLotService getListLotService()
	{
		return (ListLotService) getBean(LIST_LOT_SERVICE_BEAN);
	}

	private ProductionByLot extractRecord(DataHolder dataHolder, ProductionByLot productionByLot)
			throws Exception
	{
		log.debug("extractRecord " + dataHolder.printValues());

		BigDecimal minutes = dataHolder.get("minutes").getDecimalValue();
		String minutesTxt = new DecimalFormat("0.000000000000000").format(minutes);
//		log.warn(Integer.valueOf(minutesTxt.split("\\.")[0]) / 60);
//		log.warn(Integer.valueOf(minutesTxt.split("\\.")[0]) % 60);
//		log.warn(Math.round(Float.valueOf("0." + minutesTxt.split("\\.")[1]) * 60));
		
		productionByLot.setHour(Long.valueOf(minutesTxt.split("\\.")[0]) / 60);
		productionByLot.setMinute(Long.valueOf(minutesTxt.split("\\.")[0]) % 60);
		productionByLot.setSecond(Long.valueOf(Math.round(Float.valueOf("0." + minutesTxt.split("\\.")[1]) * 60)));
		
		productionByLot.setDialing(Long.valueOf(dataHolder.get("dialing").getIntValue()));
		productionByLot.setCompleted(Long.valueOf(dataHolder.get("completed").getIntValue()));
		productionByLot.setContact(Long.valueOf(dataHolder.get("contact").getIntValue()));
		productionByLot.setSales(Long.valueOf(dataHolder.get("sales").getIntValue()));
		productionByLot.setAbandons(Long.valueOf(dataHolder.get("abandons").getIntValue()));
		productionByLot.setUwReleaseSales(Long.valueOf(dataHolder.get("uwReleaseSales").getIntValue()));
		productionByLot.setTyp(dataHolder.get("typ").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setTotalCost(dataHolder.get("totalCost").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setReleaseSales(Long.valueOf(dataHolder.get("releaseSales").getIntValue()));
		productionByLot.setAmpPostUw(dataHolder.get("ampPostUw").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setDeclines(Long.valueOf(dataHolder.get("declines").getIntValue()));

		return productionByLot;
	}

	private void importDataHolder(ListLot listLot, Integer totalLead, Integer remainingLead, DataHolder dataHolder)
			throws Exception
	{
		Date productionDate = (Date) dataHolder.get("productionDate").getValue();
		ProductionByLot productionByLot = getProductionByLotService().findProductionByLotByListLotCodeAndProductionDate(listLot.getListLotCode(), productionDate);

		boolean newProductionByLot = false;
		if (productionByLot == null)
		{
			productionByLot = new ProductionByLot();
			newProductionByLot = true;
		}
		else
		{
			log.debug("found productionByLot record [" + productionByLot + "]");
		}

		productionByLot.setProductionDate(productionDate);
		productionByLot.setListLot(listLot);
		productionByLot.setTotalLead(Long.valueOf(totalLead));
		productionByLot.setRemainingLead(Long.valueOf(remainingLead));

		try
		{
			extractRecord(dataHolder, productionByLot);

			if (newProductionByLot)
			{
				getProductionByLotService().addProductionByLot(productionByLot, BATCH_ID);
			}
			else
			{
				getProductionByLotService().updateProductionByLot(productionByLot, BATCH_ID);
			}
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	private void importDataHolderList(ListLot listLot, Integer totalLead, Integer remainingLead, List<DataHolder> dataHolderList)
			throws Exception
	{
		for (DataHolder dataHolder : dataHolderList)
		{
			importDataHolder(listLot, totalLead, remainingLead, dataHolder);
		}
	}

	private void importFile(String fileFormatFileName, String dataFileLocation)
			throws Exception
	{
		log.info("importFile: " + dataFileLocation);
		InputStream format = null;
		InputStream input = null;
		try
		{
			format = URLClassLoader.getSystemResourceAsStream(fileFormatFileName);
			ExcelFormat excelFormat = new ExcelFormat(format);

			input = new FileInputStream(dataFileLocation);
			DataHolder fileDataHolder = excelFormat.readExcel(input);

			List<String> sheetNames = fileDataHolder.getKeyList();

			for (String sheetName : sheetNames)
			{
				DataHolder sheetDataHolder = fileDataHolder.get(sheetName);

				ListLot listLot = null;
				Integer totalLead = null;
				Integer remainingLead = null;
				List<DataHolder> listLotList = sheetDataHolder.getDataList("listLotList");
				if (listLotList.size() != 1)
				{
					throw new Exception("listLotCode invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder listLotDataHolder = listLotList.get(0);
					String listLotCode = listLotDataHolder.get("listLotCode").getStringValue();
					listLot = getListLotService().findListLotByListLotCode(listLotCode);
					if (listLot == null)
					{
						log.warn("not found listLot for listLotCode: " + listLotCode);
						continue;
					}
				}

				List<DataHolder> totalLeadList = sheetDataHolder.getDataList("totalLeadList");
				if (totalLeadList.size() != 1)
				{
					throw new Exception("totalLead invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder totalLeadDataHolder = totalLeadList.get(0);
					totalLead = totalLeadDataHolder.get("totalLead").getIntValue();
				}
				
				List<DataHolder> remainingLeadList = sheetDataHolder.getDataList("remainingLeadList");
				if (remainingLeadList.size() != 1)
				{
					throw new Exception("remainingLead invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder remainingLeadDataHolder = remainingLeadList.get(0);
					remainingLead = remainingLeadDataHolder.get("remainingLead").getIntValue();
				}

				List<DataHolder> dataHolderList = sheetDataHolder.getDataList("dataRecords");
				importDataHolderList(listLot, totalLead, remainingLead, dataHolderList);
			}
		}
		catch (Exception e)
		{
			throw e;
		}
		finally
		{
			try
			{
				format.close();
			}
			catch (Exception e)
			{
			}
			try
			{
				input.close();
			}
			catch (Exception e)
			{
			}
		}
	}

	public static void main(String[] args) throws Exception
	{
		String fileFormatFileName = /* args[0]; */ null;
//		String rootPath = /* args[1]; */ "T:/Business Solution/Share/N_Mos/Daily Sales Report/201502";
		String rootPath = "D:/project/upload file/auto report/201505_didi/OTO/addon02";

		ImportProduction batch = new ImportProduction();
		batch.setLogLevel(Logger.DEBUG);

		FileWalker fw = new FileWalker();
		fw.walk(rootPath, new FilenameFilter()
		{
			public boolean accept(File dir, String name)
			{
				return !name.contains("~$") 
						&& !name.contains("TSR") 
						&& !name.contains("SalesReportByRecords") 
						&& (name.contains("Production.xls") 
								|| name.contains("Production Report.xlsx") 
								|| (name.contains("Production Report")
										&& name.contains(".xls"))
								|| name.contains("PRODUC"));
			}
		});

		for (String filename : fw.getFileList())
		{
			if (filename.contains("TELE"))
			{
				fileFormatFileName = FILE_FORMAT_PRODUCTION_BY_LOT_TELE;
			}
			else if (filename.contains("OTO"))
			{
				fileFormatFileName = FILE_FORMAT_PRODUCTION_BY_LOT_OTO;
			}
			else {
				System.err.println("File format not found for: " + filename);
			}

			batch.importFile(fileFormatFileName, filename);
		}
	}
}
