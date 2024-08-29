package org.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.FileInputStream;

public class Main {
    public static void main(String[] args) {
        String[] amounts = {"1000", "2000"};
        String excelFilePath = "src/main/resources/MPT.xlsx";
        JSONArray dataPacksArray = new JSONArray();
        JSONArray categoryArray = new JSONArray();
        JSONArray categoryArray2 = new JSONArray();
        JSONArray categoryArray3 = new JSONArray();
        JSONArray categoryArray4 = new JSONArray();
        JSONArray categoryArray5 = new JSONArray();
        JSONArray categoryArray6 = new JSONArray();
        JSONObject dataPacks = new JSONObject();
        JSONObject category = new JSONObject();
        JSONObject category2 = new JSONObject();
        JSONObject category3 = new JSONObject();
        JSONObject category4 = new JSONObject();
        JSONObject category5 = new JSONObject();
        JSONObject category6 = new JSONObject();
        JSONObject mainObject = new JSONObject();
        JSONObject otherMessageObject = new JSONObject();
        try(FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis)){
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                if(sheet.getSheetName().equalsIgnoreCase("Data")){
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category.put("order", 2);
                        category.put("categoryName", "Data");
                        category.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category.put("category", categoryArray);
                        categoryArray.put(categoryObject);
                    }
                } else if (sheet.getSheetName().equalsIgnoreCase("Voice")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category2.put("order", 3);
                        category2.put("categoryName", "Voice");
                        category2.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category2.put("category", categoryArray2);
                        categoryArray2.put(categoryObject);
                    }
                }else if (sheet.getSheetName().equalsIgnoreCase("Gaming")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category3.put("order", 4);
                        category3.put("categoryName", "Voice");
                        category3.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category3.put("category", categoryArray3);
                        categoryArray3.put(categoryObject);
                    }
                }else if (sheet.getSheetName().equalsIgnoreCase("Entertainment")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category4.put("order", 5);
                        category4.put("categoryName", "Entertainment");
                        category4.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category4.put("category", categoryArray4);
                        categoryArray4.put(categoryObject);
                    }
                }
                else if (sheet.getSheetName().equalsIgnoreCase("Entertainment")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category4.put("order", 5);
                        category4.put("categoryName", "Entertainment");
                        category4.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category4.put("category", categoryArray4);
                        categoryArray4.put(categoryObject);
                    }
                }else if (sheet.getSheetName().equalsIgnoreCase("Data+Call Combos")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category5.put("order", 6);
                        category5.put("categoryName", "Data+Call Combos");
                        category5.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category5.put("category", categoryArray5);
                        categoryArray5.put(categoryObject);
                    }
                }else if (sheet.getSheetName().equalsIgnoreCase("SMS")) {
                    for(int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++){
                        Row row = sheet.getRow(rowIndex);
                        JSONObject packageName = new JSONObject();
                        JSONObject categoryObject = new JSONObject();
                        for(int colIndex = 1; colIndex < row.getLastCellNum(); colIndex++){
                            Cell cell = row.getCell(colIndex);
                            if(cell != null){
                                switch (cell.getColumnIndex()){
                                    case 2:
                                        packageName.put("en", cell.getStringCellValue());
                                        packageName.put("my", cell.getStringCellValue());
                                        packageName.put("zw", cell.getStringCellValue());
                                        categoryObject.put("packageName", packageName);
                                        break;
                                    case 3:
                                        categoryObject.put("packageCode", cell.getStringCellValue());
                                        break;
                                    case 4:
                                        categoryObject.put("amount", cell.getNumericCellValue());
                                        categoryObject.put("validity", "");
                                        break;
                                }
                            }
                        }
                        category6.put("order", 7);
                        category6.put("categoryName", "SMS");
                        category6.put("categoryIcon", "https://i.ibb.co/89LFNZ9/data-3x.png");
                        category6.put("category", categoryArray6);
                        categoryArray6.put(categoryObject);
                    }
                }
            }
            dataPacksArray.put(category);
            dataPacksArray.put(category2);
            dataPacksArray.put(category3);
            dataPacksArray.put(category6);
            dataPacksArray.put(category5);
            dataPacksArray.put(category6);
            dataPacks.put("dataPacks", dataPacksArray);
            mainObject.put("regex", "^(?:09|9)(7[01])\\\\d{4}$|^(?:09|9)(?:2[0-4]|5[0-6]|8[13-7]|8[19])\\\\d{5}$|^(?:09)(?:8[18])\\\\d{5}$|^(?:09|9)(?:4[1379]|73|91)\\\\d{6}$|^(?:09|9)(?:2[56]|4[0245]|8[789])\\\\d{7}$");
            mainObject.put("dataPacks", dataPacksArray);
            mainObject.put("amount", amounts);
            mainObject.put("othersEnabled", Boolean.valueOf("false"));
            mainObject.put("name", "MPT");
            mainObject.put("icon", "https://files.wavemoney.io:8199/operators/MyTel.jpg");
            otherMessageObject.put("en", "The amount must be a multiple of 1000 not larger than 30000.");
            otherMessageObject.put("my", "ပမာဏသညျ ၁၀၀၀ နှငျ့စား၍ ပွတျပွီး အမြားဆုံး ၃၀၀၀၀ ဖွစျရမညျ။");
            otherMessageObject.put("zw", "ပမာဏသည် ၁၀၀၀ နှင့်စား၍ ပြတ်ပြီး အများဆုံး ၃၀၀၀၀ ဖြစ်ရမည်။");
            mainObject.put("othersErrorMessage", otherMessageObject);
            mainObject.put("othersRegex", "");
            mainObject.put("offerType", "standard");
            System.out.println(mainObject.toString());
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}