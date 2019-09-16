package com.example.demo.service;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.persistence.Query;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ExcelUtilityService {

	@PersistenceContext
    private EntityManager em;
	
	public void updateExcel(Double key, Double value){
		try{
			FileInputStream inputStream = new FileInputStream(new File("inputFile.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = sheet.iterator();
			boolean found=false;
			while(iterator.hasNext() && !found){
				Row currentRow = iterator.next();
				Double cellValue = currentRow.getCell(0).getNumericCellValue();
				if(cellValue == key){
					currentRow.getCell(1).setCellValue(value);
					found=true;
				}
			}
			inputStream.close();
			 
            FileOutputStream outputStream = new FileOutputStream("JavaBooks.xlsx");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
		}catch (Exception e) {
			// TODO: handle exception
		}
	}
	
	public void process(){
		String csv= "";
		String[] array = csv.split(",");
		for(String val : array){
			String response = getPatientIdFromDB(val);
			if(response != null && !response.isEmpty()){
				updateExcel(Double.parseDouble(val),Double.parseDouble(response));
			}
		}
		
	}
	
	public String getPatientIdFromDB(String val){
		Query q = em.createNativeQuery("");
		String retVal = (String)q.getSingleResult();
		return retVal;
		
	}
}
