package com.gss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class RunGSSMailList {

	/**
	 * 整理日誌
	 * 
	 * @throws IOException
	 * @throws ParseException
	 */
	protected static void run(String path, Map<String, String> mapProp) throws Exception {
		List<Map<String, String>> list;

		String MonthReportExcel = path + mapProp.get("MonthReportExcel");
		System.out.println("月報Excel: " + MonthReportExcel);
		
		Workbook workbook = null;
		OutputStream output = null;
		try {
			// 整理 MAIL內容
			list = parserMailContent(path, mapProp);

			for (Map<String, String> map : list) {
				for (Entry<String, String> ent : map.entrySet()) {
					System.out.print(ent.getKey() + " : " + ent.getValue() + " , ");
				}
				System.out.println("");
			}
			
			
			File f = new File(MonthReportExcel);
			workbook = Tools.getWorkbook(MonthReportExcel, f);

			write(workbook, list);

			System.out.println("Done!");

			output = new FileOutputStream(f);
			workbook.write(output);

		} catch (Exception ex) {
			if (ex.getMessage().contains("Current browser version is")) {
				System.out.println("############################################################ \r\n"
						+ "Please change your ChromeDriver version\r\n"
						+ "############################################################ \r\n");
			}
			throw ex;
		} finally {
			try {
				if (workbook != null)
					workbook.close();
				if (output != null)
					output.close();
			} catch (IOException ex) {
				System.out.println("runMonthReport finally Error:");
				ex.printStackTrace();
			}
		}
		
	}

	/**
	 * 整理 MAIL內容
	 * 
	 * @return 整理後的list
	 * @throws Exception 
	 */
	private static List<Map<String, String>> parserMailContent(String path , Map<String, String> mapProp) throws Exception {
		boolean isPm = false;
		int hhInt = 0, arrLen = 0;//, mailTitleArrLen = 0;
		String mailTitle = "" , mailDate = "" , mailTime = "" , time = "", sender = "";
		String[] mailTitleArr;
		Map<String, String> map;
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		List<Map<String, String>> listMail;
		List<String> readStatus = Arrays.asList(new String[] {"已讀","已回覆","未讀"});
		
		/**
		 *  取 昨、今 兩天的日期
		 *  因mail中的日期會寫"昨天"、"今天"而非日期
		 */
		Calendar cal = Calendar.getInstance();
		String today = Tools.getCalendar2String(cal, "yyyyMMdd");
		cal.add(Calendar.DATE, -1);
		String yesterday = Tools.getCalendar2String(cal, "yyyyMMdd");

		// 取回mail的title
		listMail = Selenium_Crawler.getMailContent(path, mapProp);
		String[] inboxName = mapProp.get("inboxName").split(",");

		for (Map<String, String> mailMap : listMail) {

			System.out.println(mailMap.get("title"));

			/**
			 * 
			 * (收件匣)
			 * 		已標記, 高優先順序, 已讀, NHIA 林志忠, 有附件, FW: 需求單  PA111050018 , 751 KB, 22/5/5 四, 下午 2:02
			 * 		高優先順序, 已讀, NHIA 林志忠, 有附件, FW: 需求單  PA111050018 , 751 KB, 22/5/5 四, 下午 2:02
			 * 		已標記, 已回覆, Nicole Tsou (鄒文瑩), 有附件, Re: 請協助將加密程式放置正式機上_請協助通知文境至測試機測試加密功能, 288 KB, 22/5/14 六, 下午 5:29
			 * 		已讀, 張孫瑋, RE: 74.28重啟後問題, 46 KB, 22/5/18 三, 上午 8:53
			 * (0 問題單)
			 * 		已回覆, NHIA 林志忠, 有附件, FW: 請評估工時, 收件匣/1 NHIA/0 問題單, 252 KB, 22/4/27 三, 下午 1:37
			 * 		已讀, NHIA 林志忠, RE: 請評估工時_新收載QP6E_FST_TRACK_DATA、QP6E_FST_DATA, 收件匣/1 NHIA/0 問題單, 26 KB, 22/4/27 三, 上午 10:31
			 * 		已標記, 高優先順序, 已讀, NHIA 林志忠, 有附件, FW: 應用系統需求單(單號:NA110110153) , 收件匣/1 NHIA/0 問題單, 118 KB, 22/1/13 四, 上午 7:57
			 * 
			 * mailTitleArr 最後一個逗號有時會多空隔有時不會，因此先將空格移除再做split 
			 * 0: 已讀
			 * 1: NHIA 
			 * 2: 有附件 
			 * 3: 執行成功=>(386062)彙總－疾病就醫利用彙整檔(補歷史資料)(DWF)_2007/01(附檔) 
			 * 4: 收件匣/NHIA/2JOB成功 
			 * 5: 18KB
			 * 6: "今天" or "21/1/16六" 
			 * 7: 上午10:22
			 */
			mailTitleArr = mailMap.get("title").replace(" ", "").replaceAll("\t", "").split(",");
			// 有時Title會有逗號，會影響到陣列的總數
			arrLen = mailTitleArr.length;

			/**
			 * mail 收到的實際時間 格式 年月日(各兩碼) 切割時間點為上午9點 (9:00前屬當天，9:00後屬隔天)
			 */
			time = mailTitleArr[arrLen - 1].trim(); // 下午10:22
			isPm = time.substring(0, 2).equals("下午"); // 上午false 下午true
			time = time.substring(2); // 10:22
			hhInt = Integer.parseInt(time.substring(0, time.indexOf(":"))); // 10
			hhInt = hhInt == 12 ? 0 : hhInt; // 0 ~ 11
			hhInt = hhInt + (isPm ? 12 : 0); // 0 ~ 23
			// mail收到的時間(24小時制)
			mailTime = Tools.getLen2(String.valueOf(hhInt)) + time.substring(time.indexOf(":")); // 22:22

			// mail日期
			mailDate = mailTitleArr[arrLen - 2].trim();
			mailDate = mailDate.equals("昨天") ? yesterday
					: mailDate.equals("今天") ? today
							: Tools.getDate2String(
									Tools.getString2Date(mailDate.substring(0, mailDate.length() - 1), "yy/M/d"),
									"yyyyMMdd");
			if("20220519 09:30".equals(mailDate + " " +mailTime)) {
				System.out.println("wait");
			}
			// mail主旨
			mailTitle = "";
			sender = "";
			for (int i = 1; i <= arrLen; i++) {
				if (mailTitleArr[arrLen - i].equals("收件匣/1 NHIA/" + inboxName[1])) {
					mailTitle = mailTitleArr[arrLen - i - 1];
				}
				if (mailTitleArr[arrLen - i].equals("有附件")) {
					mailTitle = mailTitleArr[arrLen - i + 1];
				}
				if (readStatus.contains(mailTitleArr[arrLen - i])) {
					sender = mailTitleArr[arrLen - i + 1];
					mailTitle = mailTitle.length() > 0 ? mailTitle : mailTitleArr[arrLen - i + 2];
				}
			}

			map = new TreeMap<String, String>();
//			map.put("MailDate", mailDate);
//			map.put("MailTime", mailTime);
			map.put("MailDateTime", mailDate + " " + mailTime);
			map.put("sender", sender);
			map.put("MailTitle", mailTitle);
			list.add(map);

//			System.out.println("======================== Start ========================");
//			System.out.println("MailDateTime 日期時間 => " + mailDate + " " + mailTime);
//			System.out.println("MailTitle 主旨 => " + mailTitle);
//			System.out.println("sender 寄件者 => " + sender);
//			System.out.println("======================== End ========================");
		}
		
		return list;
	}

	/**
	 * write
	 * 
	 * @throws ParseException
	 */
	private static void write(Workbook Workbook, List<Map<String, String>> listAll) throws ParseException {
		int setColNum = 0, rowNum = 0;
		Row row;
		Cell cell = null;
//		Workbook.removeSheetAt(0);
		Sheet sheet = Workbook.createSheet(Tools.getToDay("yyyyMMdd"));
		
		for (Map<String, String> map : listAll) {
			rowNum++;
			sheet.createRow(rowNum);
			row = sheet.getRow(rowNum);
			// 設定第一欄
			setColNum = 0;
			Tools.setCellStyle(setColNum++, cell, row, map.get("MailDateTime"));
			// 設定第二欄
			Tools.setCellStyle(setColNum++, cell, row, map.get("sender"));
			// 設定第三欄
			Tools.setCellStyle(setColNum++, cell, row, map.get("MailTitle"));
		}
	}

	private static String mailTitleTrim(String str) {
		return str.toUpperCase().replace("RE:", "").replace("FW:", "").trim();
	}
	/**
	 * 檢查此項目是否已被列過或者已移至歷史清單
	 * 
	 * @param chkSheet
	 * @param map
	 * @return
	 */
//	private static boolean chkSheetForJobF(Sheet chkSheet, Map<String, String> map) {
//		String cellValue = "";
//		for (Row row : chkSheet) {
//			targetChkCell = row.getCell(1);
//			
//			if (targetChkCell != null && row.getRowNum() > 0) {
//				cellValue = "";
//				if (targetChkCell.getCellType() == Cell.CELL_TYPE_STRING)
//					cellValue = targetChkCell.getStringCellValue();
//				if (targetChkCell.getCellType() == Cell.CELL_TYPE_NUMERIC)
//					cellValue = String.valueOf((int) targetChkCell.getNumericCellValue());
//
//				if (cellValue.equals(map.get("jobSeq")))
//					return false;
//			}
//		}
//		return true;
//	}
}
