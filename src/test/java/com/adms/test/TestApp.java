package com.adms.test;

import java.util.Arrays;
import java.util.Calendar;

import com.adms.utils.DateUtil;



public class TestApp {

	public static void main(String[] args) {

		try {
			System.out.println("start");
			
			
			
			Calendar cal = Calendar.getInstance();
			System.out.println(DateUtil.convDateToString("MMM_yyyyMM", cal.getTime()));
			
			String[] plan1s = new String[]{"PLAN 1", "600000"};
			
			String planVal = "40000-Plan 1";
			
			System.out.println(Arrays.binarySearch(plan1s, planVal.toUpperCase()));
			
			
			
			System.out.println("finish");
		} catch(Exception e) {
			e.printStackTrace();
		}

	}
}
