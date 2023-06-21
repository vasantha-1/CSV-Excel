package demo;



import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aspose.cells.CellArea;
import com.aspose.cells.DataSorter;
import com.aspose.cells.SortOrder;
import com.aspose.cells.Workbook;

public class CopyndPaste {
	public static void main(String[] args) throws IOException {
		String csvfileAddress = "D:\\CSV2Excel.txt"; 
        String xlsxfileAddress ="D:\\ExcelTask\\xl.xlsx";
        
        XSSFWorkbook wbk = new XSSFWorkbook();
        XSSFSheet sheet = wbk.createSheet("sheet1");
        String currentLine=null;
        int RowNO=-1;
        BufferedReader br = new BufferedReader(new FileReader(csvfileAddress));
        while ((currentLine = br.readLine()) != null) {
        	currentLine =currentLine.replaceAll("\"", "");
            String s[] = currentLine.split(",");
            RowNO++;
            XSSFRow currentRow=sheet.createRow(RowNO);
            for(int i=0;i<s.length;i++){
                currentRow.createCell(i).setCellValue(s[i]);
            }
        }      
       sheet.setAutoFilter(CellRangeAddress.valueOf("A1:Z60640"));
       
       
        FileOutputStream fos =  new FileOutputStream(xlsxfileAddress);
        wbk.write(fos);
        fos.close();
        System.out.println("Done");
      
	}
	

}
