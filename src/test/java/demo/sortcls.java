package demo;

import com.aspose.cells.CellArea;
import com.aspose.cells.DataSorter;
import com.aspose.cells.SortOrder;
import com.aspose.cells.Workbook;


public class sortcls {

	public static void main(String[] args) throws Exception {	
		Workbook  a = new Workbook("D:\\ExcelTask\\xl.xlsx");
		      DataSorter sort = a.getDataSorter();
		     
		      sort.setOrder1(SortOrder.ASCENDING);
		      sort.setKey1(6);	    
		      CellArea CA = CellArea.createCellArea("A2", "U60639");
		     sort.sort(a.getWorksheets().get(0).getCells(), CA);
		      a.save("xl.xlsx");
		      System.out.println("done");
		         
	}	      
	

}
