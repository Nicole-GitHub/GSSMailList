package com.gss;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jsoup.helper.StringUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

import us.codecraft.webmagic.selector.Html;

public class Selenium_Crawler {

	static ChromeDriver driver = null;
	static Html html;
	static WebElement element;

	protected static List<Map<String, String>> getMailContent(String path, Map<String, String> mapProp) throws Exception {
		
		driver = null;
		List<WebElement> listElement;
		List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
		Map<String, String> map;
		String title = "";

		String os = System.getProperty("os.name");

		// 收件匣名稱
		String[] inboxName = mapProp.get("inboxName").split(",");
		String inboxNameStr = "";
		for (String str : inboxName)
			inboxNameStr += inboxNameStr.length() > 0 ? ", " + str : str;
		System.out.println("收件匣名稱: " + inboxNameStr);

		// 帳號
		String account = mapProp.get("account");

		// 密碼
		String pwd = mapProp.get("pwd");
		
		/** 
		 * 執行日期小於15號則從前一個月的月初開始取，反之則從當月月初開始
		 * 為防止月初沒mail的情況，故設定滑動頁面"最多"滑到信件時間出現欲檢查日期的前10天出現為止
		 */
		Calendar chkDate = Calendar.getInstance();
		int chkMailDateLen = 10;
		
		Integer dd = Integer.parseInt(Tools.getCalendar2String(chkDate, "dd"));
		if(dd < 15) {
			chkDate.add(Calendar.MONTH, -1);
		}
		/**
		 *  設定為上月最後一個工作天的日期
		 *  set DATE -1 : 設為上月最後一個工作天
		 *  set DATE 1 : 設為當月1號
		 *  add DATE -1 : 日期 - 1
		 *  add DATE 1 : 日期 + 1
		 */
		chkDate.set(Calendar.DATE,-1);
		
		String[] calArr = new String[chkMailDateLen];
		for (int i = 0; i < chkMailDateLen; i++) {
			calArr[i] = Tools.getCalendar2String(chkDate, "yy/M/d");
			chkDate.add(Calendar.DATE, -1);
		}
		
		String chromeDriver = path + mapProp.get("chromeDriverPath") + mapProp.get("chromeDriverName") + "_"
				+ mapProp.get("chromeDriverVersion") + (os.contains("Mac") ? "" : ".exe");
		
		// Selenium
		DesiredCapabilities capabilities = DesiredCapabilities.chrome();
		capabilities.setCapability("chrome.switches", Arrays.asList("--start-maximized"));
		System.setProperty("webdriver.chrome.driver", chromeDriver);

		driver = new ChromeDriver(capabilities);
		
		driver.get(mapProp.get("webAddress"));
		System.out.println("##start login ");

		try {
			// 延迟加载，保证JS数据正常加载
			Thread.sleep(1000);

			// 登入Mail
			driver.findElement(By.id("username")).sendKeys(account);
			driver.findElement(By.id("password")).sendKeys(pwd);
			driver.findElement(By.className("DwtButton")).click();

			// 等待三秒以確保頁面加載完整
			Thread.sleep(3000);
			
			for (String str : inboxName) {

				// 進入對應的收信匣
				listElement = driver.findElements(By.xpath("//td[@class='DwtTreeItem-Text']"));
				for (WebElement em : listElement) {
					if (em.getText().contains(str)) {
						System.out.println(em.getText());
						em.click();
						break;
					}
				}

				Thread.sleep(1000);
				/**
				 * 內容加載後截取信件list區塊 再拆分為主旨、內容兩部份放入map中
				 */
				if (scrollDown(calArr, chkMailDateLen)) {
					listElement = driver.findElements(By.className("Row"));

					for (WebElement em : listElement) {
						map = new HashMap<String, String>();
						title = em.getAttribute("aria-label");
						map.put("title", title);
						listMap.add(map);
					}
					
				}
			}

		} catch (Exception e) {
			System.out.println("Selenium_Crawler Error：" + e.getMessage());
			throw e;
		} finally {
			if(driver != null)
				driver.close();
		}
		return listMap;
	}

	/**
	 * 滑動頁面
	 */
	private static boolean scrollDown(String[] calArr, int chkMailDateLen) {
		if (driver != null) {
			try {
				element = driver.findElement(By.id("zl__TV-main__rows"));
				html = new Html(element.getAttribute("outerHTML"));
				String scroll = "go";
				boolean dateisBlank = true;
				
				while ("go".equals(scroll)) {
					// 最少要滾到檢查日期的前chkMailDateLen天
					for (int calArrLen = 1; calArrLen < chkMailDateLen; calArrLen++) {
						dateisBlank = StringUtil
								.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[calArrLen] + "')]").get());
						// 若已滾到檢查日期的前chkMailDateLen天則可停止
						if (!dateisBlank)
							break;
					}
					if (!dateisBlank)
						break;

					html = new Html(element.getAttribute("outerHTML"));
					// 執行頁面滾動的JS語法
					String height1 = ((JavascriptExecutor) driver)
							.executeScript("var element = document.getElementById('zl__TV-main__rows');"
									+ "var height1 = element.scrollHeight;"
									+ "element.scroll(0,height1);"
									+ "return height1;")
							.toString();
					Thread.sleep(1000);
					String height2 = ((JavascriptExecutor) driver)
							.executeScript("var element = document.getElementById('zl__TV-main__rows');"
									+ "var height2 = element.scrollHeight;"
									+ "return height2;")
							.toString();
					/**
					 * height1: 未滾前的高度
					 * height2: 滾動後的高度
					 * 若兩個高度皆相同則表示已滾到底
					 */
					scroll = Integer.parseInt(height1) == Integer.parseInt(height2) ? "stop" : "go";
					// 给页面预留加载时间
					Thread.sleep(1000);
				}
				System.out.println("加載中...");
				return true;
			} catch (Exception e) {
				System.out.println("加載失敗:");
				e.printStackTrace();
			}
		}
		return false;
	}

}
