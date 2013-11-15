#Class to convert Excel to Json
/**
 * @author Rohan Kamat
 */

public class ExcelToJson{

public JSONArray constructJsonArrayFromExcel(File file) throws Exception{
		JSONArray result=new JSONArray();
		JSONArray json= processExcelFile(file);
		System.out.println(json);
		if(!json.isEmpty()){
			JSONObject ColumnNames=(JSONObject) json.get(0);
			for(int i=1;i<json.size();i++){
				JSONObject exelobject=new JSONObject();
				JSONObject object=(JSONObject) json.get(i);
				for(int j=0;j<ColumnNames.size();j++){
					exelobject.put(ColumnNames.get(j), object.get(j));
				}
				result.add(exelobject);
			}
		}
		return result;
	}

private JSONArray processExcelFile(File file) throws Exception{
		   JSONArray rows = new JSONArray();
     try{
         FileInputStream myInput = new FileInputStream(file);
         XSSFWorkbook myWorkBook = new XSSFWorkbook(myInput);
         XSSFSheet mySheet = myWorkBook.getSheetAt(0);
         Iterator<Row> rowIter = mySheet.rowIterator();
         while(rowIter.hasNext()){
             XSSFRow myRow = (XSSFRow) rowIter.next();
             Iterator<Cell> cellIter = myRow.cellIterator();
             JSONObject jRow = new JSONObject();
             while(cellIter.hasNext()){
                 XSSFCell myCell = (XSSFCell) cellIter.next();
                 switch (myCell.getCellType()) {
                 case XSSFCell.CELL_TYPE_NUMERIC :
                     jRow.put(myCell.getColumnIndex(), myCell.getNumericCellValue());
                     break;
                 case XSSFCell.CELL_TYPE_STRING:   
                     jRow.put(myCell.getColumnIndex(), myCell.getStringCellValue());
                     break;
                 default:   
                     jRow.put(myCell.getColumnIndex(), myCell.getRawValue());
                     break;   
                 }
             }
             rows.add(jRow);
         }
     }
     catch (Exception ex){
    System.out.println(ex);
     }
     return rows;
 }
