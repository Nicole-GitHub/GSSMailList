package com.gss;

import java.io.File;
import java.util.Date;
import java.util.Map;

public class GSSMailListMain {

	public static void main(String[] args) {
		try {

			// 取得jar檔的絕對路徑
//			System.out.println("3:"+ ClassLoader.getSystemResource(""));
//			System.out.println("4:"+ DailyReport.class.getResource(""));//DailyReport.class檔案所在路徑
//			System.out.println("5:"+ DailyReport.class.getResource("/")); // Class包所在路徑,得到的是URL物件,用url.getPath()獲取絕對路徑String
//			System.out.println("6:"+ new File("/").getAbsolutePath());
//			System.out.println("7:"+ System.getProperty("user.dir"));
//			System.out.println("9:"+ System.getProperty("java.class.path"));

			String path = System.getProperty("user.dir") + File.separator; // Jar

			String os = System.getProperty("os.name");
			System.out.println("=== NOW TIME ===> " + new Date());
			System.out.println("===os.name===> " + os);

			// Debug
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/DailyReport/" // Mac
					: "C:/Users/Nicole/Dropbox/DailyReport/"; // win
			
			System.out.println("path: " + path);
			Map<String, String> mapProp = Property.getProperties(path);

			boolean done = false;
			do {
				try {
					RunGSSMailList.run(path, mapProp);
					done = true;
				} catch (Exception e) {
					System.out.println(new Date() + " ===> " + e.getMessage());
					if ("getDailyReportExcel Error".equals(e.getMessage())
							|| e.getMessage().contains("This version of ChromeDriver only supports Chrome version"))
						done = true;
				}
			} while (!done);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
