package com.hindu.ExcelSplitter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * Handles requests for the application home page.
 */
@Controller
public class HomeController {
	
	private static final Logger logger = LoggerFactory.getLogger(HomeController.class);
	
	/**
	 * Simply selects the home view to render by returning its name.
	 */
	@RequestMapping(value = "/", method = RequestMethod.GET)
	public String home(Locale locale, Model model) {
		logger.info("Welcome home! The client locale is {}.", locale);
		
		Date date = new Date();
		DateFormat dateFormat = DateFormat.getDateTimeInstance(DateFormat.LONG, DateFormat.LONG, locale);
		
		String formattedDate = dateFormat.format(date);
		
		model.addAttribute("serverTime", formattedDate );
		
		/*String fileLocation = "D:\\TestData\\hind.xlsx";
		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileLocation));
			Workbook workbook = new XSSFWorkbook(file);
			
			Sheet sheet = workbook.getSheetAt(0);
			 
			Map<Integer, List<String>> data = new HashMap<Integer, List<String>>();
			int i = 0;
			for (Row row : sheet) {
			    data.put(i, new ArrayList<String>());
			    for (Cell cell : row) {
			    	switch(cell.getCellTypeEnum()) {
			    		case STRING:
			    			data.get(new Integer(i)).add(cell.getStringCellValue());
			    	}
			    	
			    }
			    i++;
			}
			workbook.close();
			
			model.addAttribute("data", data);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/
		

		return "home";
	}
	
	@RequestMapping(value = "/download")
	public void download(HttpServletResponse response) {
		String baseFileLocation = "D:\\TestData\\hind.xlsx";
		FileInputStream baseFile;
		
		String newFileLocation = "D:\\TestData\\newhind.xlsx";
		FileOutputStream fout;
		
		try {
			baseFile = new FileInputStream(new File(baseFileLocation));
			Workbook baseWorkbook = new XSSFWorkbook(baseFile);
			
			Sheet baseWorksheet = baseWorkbook.getSheetAt(0);
			
			Workbook newWorkbook = new XSSFWorkbook();
			Sheet newSheet = newWorkbook.createSheet();

			int i = 0;
			for (Row row : baseWorksheet) {
			    Row newRow = newSheet.createRow(i);
			    for (Cell cell : row) {
			    	switch(cell.getCellTypeEnum()) {
			    		case STRING:
			    			Cell newCell = newRow.createCell(cell.getColumnIndex());
			    			newCell.setCellValue(cell.getStringCellValue());
			    	}
			    	
			    }
			    i++;
			}
			
			fout = new FileOutputStream(newFileLocation);
			newWorkbook.write(fout);
			newWorkbook.close();
			
			baseWorkbook.close();
	
			response.setContentType("application/vnd.ms-excel");
			response.setHeader("Content-Disposition", "attachment; filename=newHindu" + System.currentTimeMillis() + ".xlsx");
			
			Path newFile = Paths.get(newFileLocation);
			Files.copy(newFile, response.getOutputStream());
			response.getOutputStream().flush();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
}
