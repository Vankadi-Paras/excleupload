package com.salesupload.Tabledef.service.Impl;

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
import com.salesupload.Tabledef.enitity.Tabledefinition;
import com.salesupload.Tabledef.repository.TableDefRepository;
import com.salesupload.Tabledef.service.TableDefService;
import com.salesupload.common.AjaxResponse;
import com.salesupload.common.DBUtils;
import com.salesupload.common.DataTablesUtility;
import com.salesupload.common.DropDownUtils;
import com.salesupload.common.ExcelExport;
import com.salesupload.common.TextEncryptDecrypt;

@Service
public class TableDefServiceImpl implements TableDefService {

	@Autowired
	TableDefRepository tableDefRepository;
	@Autowired
	AjaxResponse ajaxResponse;

	@Autowired
	DBUtils dbUtils;
	@Autowired
	HttpServletRequest request;

	@Autowired
	ExcelExport excelExport;

	@Autowired
	DropDownUtils dropDownUtils;

	@Autowired
	DataTablesUtility dataTablesUtility;
	@Autowired
	Environment environment;

	@Autowired
	TextEncryptDecrypt textEncryptDecrypt;

	@Override
	public String getPage() {
		return "master/tabledefinition/table_definition";
	}

	@Override
	public ResponseEntity<Object> SaveOrUpdate(Tabledefinition tabledefinition) {
		try {
			Optional<Tabledefinition> table = null;
			if (tabledefinition.getTabledefinitionid() != 0) {
				table = tableDefRepository.findById(tabledefinition.getTabledefinitionid());
			}
			String str = "SELECT tablename FROM tabledefinition WHERE ";
			Map<String, Object> paramMap = new HashMap<>();
			int id = tabledefinition.getTabledefinitionid();
			if (id > 0) {
				str += " tabledefinitionid != :tabledefinitionid AND ";
				paramMap.put("tabledefinitionid", id);
			}
			String tablename = tabledefinition.getTablename();
//			String excelfile = tabledefinition.getExcelfile();
			if (tablename != null)  {
				str += "   LOWER(tablename) = LOWER(:tablename)";
				paramMap.put("tablename", tablename);
			}
//			if (excelfile != null)  {
//				str += " and   LOWER(excelfile) = LOWER(:excelfile)";
//				paramMap.put("excelfile", excelfile);
//			}
			List<?> isExist = dbUtils.returnResultSet(str, paramMap);
			if (isExist.size() > 0) {
				ajaxResponse.setStatus("401");
				ajaxResponse.setMessage("tablename already available..!");
				return new ResponseEntity<>(ajaxResponse, HttpStatus.UNAUTHORIZED);
			}
			String str1 = "SELECT excelfile FROM tabledefinition WHERE ";
			
			Map<String, Object> paramMap1 = new HashMap<>();
			int id1 = tabledefinition.getTabledefinitionid();
			if (id1 > 0) {
				str1 += " tabledefinitionid != :tabledefinitionid AND ";
				paramMap1.put("tabledefinitionid", id);
			}
			String excelfile = tabledefinition.getExcelfile();
			if (excelfile != null)  {
				str1 += "   LOWER(excelfile) = LOWER(:excelfile)";
				paramMap1.put("excelfile", excelfile);
			}
			List<?> isExist1 = dbUtils.returnResultSet(str1, paramMap1);
			if (isExist1.size() > 0) {
				ajaxResponse.setStatus("401");
				ajaxResponse.setMessage("Excel file  already available..!");
				return new ResponseEntity<>(ajaxResponse, HttpStatus.UNAUTHORIZED);
			}
			else {

				if (table != null) {
					tabledefinition.setTabledefinitionid(tabledefinition.getTabledefinitionid());
					tabledefinition.setTablecreated(table.get().getTablecreated());
					tabledefinition.setActive(table.get().getActive());

					tableDefRepository.save(tabledefinition);
					ajaxResponse.setStatus("200");
					ajaxResponse.setMessage("succesfully data updated");
					ajaxResponse.getDataMap().put("redirectUrl",
							environment.getProperty("server.servlet.context-path") + "/tabledefinition");
					return new ResponseEntity<Object>(ajaxResponse, HttpStatus.OK);

				} else {
					Integer maxid = Integer.parseInt(dbUtils.getMaxId("tabledefinition", "tabledefinitionid", null));
					tabledefinition.setTabledefinitionid(maxid);
					tabledefinition.setActive(1);
					tabledefinition.setTablecreated("No");
					tableDefRepository.save(tabledefinition);
					ajaxResponse.setMessage("succesfully data saved");
					ajaxResponse.setStatus("200");
					ajaxResponse.getDataMap().put("redirectUrl",
							environment.getProperty("server.servlet.context-path") + "/tabledefinition");
					return new ResponseEntity<Object>(ajaxResponse, HttpStatus.OK);

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return new ResponseEntity<>("Internal server ", HttpStatus.INTERNAL_SERVER_ERROR);
	}

	public ResponseEntity<Object> getTableDefinition() {

		String sqlQuery = "SELECT tabledefinitionid,filetypename,tablename,excelfile,tablecreated,(SELECT COUNT(1) FROM tablecreationdetail tc  WHERE td.tabledefinitionid=tc.tabledefinitionid) AS COUNT FROM tabledefinition  td";
		String whereClause = " WHERE active=1 ";
		String countQuery = "SELECT COUNT(*) FROM tabledefinition";
		try {
			return new ResponseEntity<Object>(dataTablesUtility.getDataInJson(request, sqlQuery, countQuery,
					whereClause, "tabledefinition", new String[] { "tabledefinitionid", "filetypename", "tablename",
							"excelfile", "tablecreated" }),
					HttpStatus.OK);
			
		
		} catch (JsonProcessingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	public String tableCreationAdd(String tableId, Model model) {
		try {
			Integer id = Integer.parseInt(textEncryptDecrypt.decrypt(tableId));
			Optional<Tabledefinition> tabledefinition = tableDefRepository.findById(id);
			model.addAttribute("tablenameid", id);
			model.addAttribute("tablename", tabledefinition.get().getTablename());
			return "master/tablecreationadd/table_creation_add";
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;

	}

	@Override
	public ResponseEntity<?> getColumnDetail(String tabledefinitionid) {
		try {
			String sqlQuery = "select columnname,columndescription,ct.name,columntypevalue,ismandotary,(case ismandotary when '1' then 'Yes' else 'No' end) as ismandotary1 from tablecreationdetail td,columntypemaster ct where td.columntypeid=ct.columntypeid and td.tabledefinitionid=:tabledefinitionid";

			Integer userid1 = Integer.parseInt(textEncryptDecrypt.decrypt(tabledefinitionid.toString()));
			Map<String, Object> parameterMap = new HashMap<>();
			if (userid1 != null) {

				ajaxResponse.getDataMap().clear();
				parameterMap.put("tabledefinitionid", userid1);
				List<?> columnDetailMasterList = dbUtils.returnResultSet(sqlQuery, parameterMap);
				ajaxResponse.getDataMap().put("columnDetailList", columnDetailMasterList);
				ajaxResponse.setStatus("200");
				ajaxResponse.setMessage("detail data");
				return new ResponseEntity<>(ajaxResponse, HttpStatus.OK);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		ajaxResponse.setStatus("500");
		ajaxResponse.setMessage("Internal server occured!");
		return new ResponseEntity<>(ajaxResponse, HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@Override
	public void exportTableDefToExcel(HttpServletResponse response, String tabledefid) {

		try {
			Integer userid1 = Integer.parseInt(textEncryptDecrypt.decrypt(tabledefid.toString()));
			//String sql = "SELECT tablename FROM tabledefinition WHERE tabledefinitionid = :defid";
			String sql = "SELECT columnname FROM tablecreationdetail WHERE tabledefinitionid = :defid";
			Map<String, Object> paramMap = new HashMap<>();
			// paramMap.put("defid", Integer.parseInt(tabledefid));
			paramMap.put("defid", userid1);
			
			String filename = String.valueOf(dbUtils.returnResultSet("SELECT excelfile FROM `tabledefinition` WHERE `tabledefinitionid`=" + userid1, null).get(0));

			List<?> resultList = dbUtils.returnResultSet(sql, paramMap);
			
			paramMap.clear();
			String [] columns = resultList.toArray(new String[resultList.size()]);
			System.out.println(columns);
			String filepath = createExcel(columns,filename);//excelExport.usingQuery(columns, "SELECT excelfile FROM `tabledefinition` where 1=2  ", filename ,	paramMap);

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

	@Override
	public ResponseEntity<Object> edittabdef(String id) {
		try {
			Integer userid = Integer.parseInt(textEncryptDecrypt.decrypt(id.toString()));
			Optional<Tabledefinition> tabdef = tableDefRepository.findById(userid);
			if (tabdef != null) {
				ajaxResponse.setStatus("200");
				ajaxResponse.setMessage("Add/Edit TableDeffination Successfully");
				ajaxResponse.getDataMap().clear();
				ajaxResponse.getDataMap().put("tabdeflist", tabdef);
				return new ResponseEntity<>(ajaxResponse, HttpStatus.OK);
			} else {
				ajaxResponse.setStatus("404");
				ajaxResponse.setMessage("data not found.");
				ajaxResponse.getDataMap().clear();
				ajaxResponse.getDataMap().put("redirectUrl", "/");
				return new ResponseEntity<Object>(ajaxResponse, HttpStatus.NOT_FOUND);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
}
