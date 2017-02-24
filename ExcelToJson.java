
import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;

public class ExcelToJson {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		String ExcelPath = "C:\\folder\\filename.xlsx"; // Mention the path of excel
		File FileName = new File(ExcelPath);
		JSONObject excelContents= excelToJson(FileName,"Heading"); // Pass the first column heading name as second parameter   
		System.out.println(excelContents.toString());		
	}
	
	public static JSONObject excelToJson(File FileName,String heading) throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException{
		Workbook Wb = WorkbookFactory.create(FileName);
		int NoOfSheets = Wb.getNumberOfSheets();
		JSONObject Json = new JSONObject();
		for(int i=0;i<NoOfSheets;i++){
			Sheet sheet = Wb.getSheetAt(i);
			boolean headingAvail = false;
			int headingColnIndex = 0,headingRowIndex=0, lastRowIndex=0,lastCellIndex=0;
			DataFormatter df = new DataFormatter();
			loop:for(Row rows : sheet){
				for(Cell cell : rows){
					if(heading.equalsIgnoreCase(df.formatCellValue(cell))){
						headingAvail = true;
						headingColnIndex = cell.getColumnIndex();
						headingRowIndex = cell.getRowIndex();
						lastRowIndex = rows.getLastCellNum();
						lastCellIndex = rows.getLastCellNum();
						break loop;
					}
				}
			}
			
			if(headingAvail){
				JSONArray JSheet = new JSONArray();
				System.out.println("true");
				for(int j= headingRowIndex+1;j<lastRowIndex;j++){
					JSONObject Jrow = new JSONObject();
					for(int k=headingColnIndex ; k<lastCellIndex;k++){
						Row Heading = sheet.getRow(headingRowIndex);
						Row row = sheet.getRow(j);
						Jrow.put(""+Heading.getCell(k), df.formatCellValue(row.getCell(k)));
					}
					JSheet.put(Jrow);
				}
				Json.put("Sheet "+i, JSheet);
			}else{
				System.out.println(" Heading is not available in the sheet"+(i+1));
			}
			
			
		}
	
		return Json;
	}

}

