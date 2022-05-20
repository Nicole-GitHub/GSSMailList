package com.gss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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

		String MonthReportExcelCName = mapProp.get("MonthReportExcelCName");
		String MonthReportExcelExt = mapProp.get("MonthReportExcelExt");
		String MonthReportExcelSource = path + MonthReportExcelCName + MonthReportExcelExt;
		System.out.println("月報Excel: " + MonthReportExcelSource);
		
		String MonthReportExcelTarget = path + MonthReportExcelCName + Tools.getToDay("yyyyMMdd") + MonthReportExcelExt;

		Workbook workbook = null;
		OutputStream output = null;
		try {
			// 整理 MAIL內容
			list = parserMailContent(path, mapProp);

			for (Map<String, String> map : list) {
				for (Entry<String, String> ent : map.entrySet()) {
					System.out.println(ent.getKey() + " : " + ent.getValue() + " , ");
				}
			}
			
//			
//			File f = new File(MonthReportExcelSource);
//			workbook = Tools.getWorkbook(MonthReportExcelSource, f);
////			Sheet sheet1 = workbook.getSheetAt(0);
//
//			// 寫入 "JobList" 頁籤的狀態，並整理出失敗的Job
////			writeSheet3(workbook, list);
//
//			System.out.println("Done!");
//
//			output = new FileOutputStream(f);
//			workbook.write(output);
//
//			f.renameTo(new File(MonthReportExcelTarget)); //改名
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
		List readStatus = Arrays.asList(new String[] {"已讀","已回覆","未讀"});
		
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

		for (Map<String, String> mailMap : listMail) {

			System.out.println(mailMap.get("title"));

			/**
			 * 
			 * (收件匣)
			 * 		已標記, 高優先順序, 已讀, NHIA 林志忠, 有附件, FW: 需求單  PA111050018 , 751 KB, 22/5/5 四, 下午 2:02
			 * 		高優先順序, 已讀, NHIA 林志忠, 有附件, FW: 需求單  PA111050018 , 751 KB, 22/5/5 四, 下午 2:02
			 * 		已標記, 已回覆, Nicole Tsou (鄒文瑩), 有附件, Re: 請協助將加密程式放置正式機上_請協助通知文境至測試機測試加密功能, 288 KB, 22/5/14 六, 下午 5:29
			 * (0 問題單)
			 * 		已回覆, NHIA 林志忠, 有附件, FW: 請評估工時, 收件匣/1 NHIA/0 問題單, 252 KB, 22/4/27 三, 下午 1:37
			 * 		已讀, NHIA 林志忠, RE: 請評估工時_新收載QP6E_FST_TRACK_DATA、QP6E_FST_DATA, 收件匣/1 NHIA/0 問題單, 26 KB, 22/4/27 三, 上午 10:31
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
			mailTitleArr = mailMap.get("title").replace(" ", "").split(",");
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
			// mail主旨
//			mailTitleArrLen = 0;
			mailTitle = "";
			sender = "";
			for (int i = 1; i < arrLen; i++) {
				if (mailTitleArr[arrLen - i].equals("收件匣/1 NHIA/0 問題單")) {
					mailTitle = mailTitleArr[arrLen - i - 1];
//					mailTitleArrLen = arrLen - i - 1;
//					break;
				}
				if (readStatus.contains(mailTitleArr[arrLen - i])) {
					sender = mailTitleArr[arrLen - i + 1];
//					mailTitleArrLen = arrLen - i - 1;
//					break;
				}
				if(sender.length() > 0 && mailTitle.length() > 0)
					break;
			}
//			mailTitle = mailTitleArr[mailTitleArrLen];
//			sender = mailTitleArr[1];

			map = new HashMap<String, String>();
			map.put("MailDate", mailDate);
			map.put("MailTime", mailTime);
			map.put("MailDateTime", mailDate + " " + mailTime);
			map.put("MailTitle", mailTitle);
			map.put("sender", sender);
			list.add(map);

			System.out.println("======================== Start ========================");
			System.out.println("MailDateTime 日期時間 => " + mailDate + " " + mailTime);
			System.out.println("MailTitle 主旨 => " + mailTitle);
			System.out.println("sender 寄件者 => " + sender);
			System.out.println("======================== End ========================");
		}
		
		return list;
	}

	/**
	 * 將失敗的job列進 "待辦JOB" 頁籤
	 * 
	 * @throws ParseException
	 */
//	private static void writeSheet3(Workbook Workbook, List<Map<String, String>> list) throws ParseException {
//		int setColNum = 0, lastRowNum = 0;
//		Row row;
//		Cell cell = null;
//		Map<String, String> map;
//		Sheet sheet3 = Workbook.getSheetAt(2);
//		Sheet sheet4 = Workbook.getSheetAt(3);
//		CellStyle cellStyle = Workbook.createCellStyle();
//		lastRowNum = sheet3.getLastRowNum();
//
//		// 為了讓list能由後往前讀，故使用ListIterator
//		ListIterator<TreeMap<String, String>> listIterator = listFforSheet3.listIterator();
//		// 先讓迭代器的指標移到最尾筆
//		while (listIterator.hasNext()) {
//			System.out.println("待辦 job : " + listIterator.next());
//		}
//		// 再由後往前讀出來
//		while (listIterator.hasPrevious()) {
//			map = listIterator.previous();
//			// 若出現不在當月日誌清單內的失敗job則跳過
//			if(map.get("jobRSDate") == null) {
//				continue;
//			}
//			JobMonth = map.get("jobRSDate").substring(0, 6);
//			isPrint = true;
//			// 判斷是否為當月的日誌
//			if (JobMonth.equals(excelMonth)) {
//				// 檢查此項目是否已被列過或者已移至歷史清單
//				isPrint = chkSheetForJobF(sheet3, map) && chkSheetForJobF(sheet4, map);
//
//				if (isPrint) {
//					lastRowNum++;
//					sheet3.createRow(lastRowNum);
//					row = sheet3.getRow(lastRowNum);
//					String dateStr = map.get("jobRSDate").substring(0, 4) + "/"
//							+ map.get("jobRSDate").substring(4, 6) + "/" + map.get("jobRSDate").substring(6);
//					// 設定第一欄
//					setColNum = 0;
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, dateStr);
//					// 設定第二欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobSeq"));
//					// 設定第三欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobEName"));
//					// 設定第四欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobName"));
//					// 設定第五欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobPeriod"));
//					// 設定第六欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, "");
//					// 設定第七欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4,
//							map.get("RQ_rq_id") + " => " + map.get("RQ_run_flag"));
//					// 設定第八欄
//					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, "");
//				}
//			}
//		}
//	}

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
