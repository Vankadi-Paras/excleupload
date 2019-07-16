package com

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URLConnection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.env.Environment;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.ui.Model;

import com.fasterxml.jackson.core.JsonProcessingException;



@Service
public class ExcleUpload
{
	@Autowired
	ExcelExport excelExport;


	@Autowired
	Environment environment;

	@Autowired
	TextEncryptDecrypt textEncryptDecrypt;

	@Override
	public void exportExcel(HttpServletResponse response, String id) {

		try {
			Integer userid1 = Integer.parseInt(textEncryptDecrypt.decrypt(id.toString()));
			
			String sql = "SELECT columnname FROM tablename WHERE id = :defid";
			Map<String, Object> paramMap = new HashMap<>();
			
			paramMap.put("defid", userid1);
			
			String filename = String.valueOf(dbUtils.returnResultSet("SELECT excelfile FROM `tablename` WHERE `id`=" + userid1, null).get(0));

			List<?> resultList = dbUtils.returnResultSet(sql, paramMap);
			
			paramMap.clear();
			String [] columns = resultList.toArray(new String[resultList.size()]);
			System.out.println(columns);
			String filepath = createExcel(columns,filename);//excelExport.usingQuery(columns, "SELECT excelfile FROM `tablename` where 1=2  ", filename ,	paramMap);

			File file = new File(filepath);

			String mimeType = URLConnection.guessContentTypeFromName(file.getName());
			if (mimeType == null) {
				System.out.println("mimetype is not detectable, will take default");
				mimeType = "application/octet-stream";
			}

			System.out.println("mimetype : " + mimeType);

			response.setContentType(mimeType);

			response.setHeader("Content-Disposition", String.format("attachment; filename=\"%s\"", file.getName()));

			response.setContentLength((int) file.length());

			InputStream inputStream = new BufferedInputStream(new FileInputStream(file));

			if (null != inputStream) {
				IOUtils.copy(inputStream, response.getOutputStream());
			}

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	@Autowired
	private ServletContext servletContext;
	
	public String createExcel(String columns[],String file_name) {
		try {
			
		String myPath = servletContext.getRealPath("/excel");
		String fileName = file_name + ".xlsx";
		
		File f = new File(myPath + File.separator + fileName);
		
		File myPathFile = new File(myPath);
		
		if(!myPathFile.exists())
			myPathFile.mkdir();

		if (!f.exists())
			f.createNewFile();
		
		Cell cell = null; Row row = null;
		SXSSFWorkbook wb = new SXSSFWorkbook(100); 
		wb.setCompressTempFiles(true);
		  
		CreationHelper createHelper = wb.getCreationHelper();
		  
		CellStyle dateCellStyle = wb.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
		  
		CellStyle timeCellStyle = wb.createCellStyle();
		timeCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("h:mm"));
		  
		CellStyle dateTimeCellStyle = wb.createCellStyle();
		dateTimeCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
		  
		Sheet sheet1 = wb.createSheet("Sheet1"); 
		row = sheet1.createRow(0); 
		for (int i = 0; i < columns.length; i++) { 
			cell = row.createCell(i);
			cell.setCellValue(columns[i]); 
		}
		
		FileOutputStream fileOutputStream = new FileOutputStream(f);
		wb.write(fileOutputStream);
		wb.close();  
		fileOutputStream.flush(); 
		fileOutputStream.close();
		
		return f.getAbsolutePath();
		}catch(Exception e) {
			e.printStackTrace();
			return null;
		}
		
		
	}

}
