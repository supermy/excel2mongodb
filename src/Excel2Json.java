import java.io.File;
import java.io.IOException;
import java.net.UnknownHostException;

import jxl.Sheet;
import jxl.Workbook;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringEscapeUtils;

import com.mongodb.BasicDBObject;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.Mongo;

/**
 * 
 * Excel to Json
 * 
 * @author jamesmo
 * @version 创建时间：2012-11-9 上午10:46:51
 *  
 */
public class Excel2Json {
	public Sheet sheet;
	/**
	 * 将Excel内容存放到二维数组中。
	 * 
	 * @param is
	 * @return
	 */
	public String[][] getExcelContent(String filename,int sheetnum) {
		try {
			Workbook workBook = Workbook.getWorkbook(new File(filename));

			//Workbook workBook = Workbook.getWorkbook(is);
			sheet = workBook.getSheet(sheetnum);
			
			int sheetColumns = sheet.getColumns();
			int sheetRows = sheet.getRows();
			// excel内容
			String[][] excelContent = new String[sheetRows][sheetColumns];
			for (int i = 0; i < sheetRows; i++) {
				for (int j = 0; j < sheetColumns; j++) {
					// 将excel值放入二维数组excelContent中
					excelContent[i][j] = sheet.getCell(j, i).getContents();
				}
			}
			return excelContent;
		} catch (Exception e) {
			// TODO: handle exception
		}
		return null;

	}

	private static final String QUOTATION_MARKS = "\""; // 双引号
	private static final String COMMA = ","; // 逗号

	public StringBuffer convertToJson(String[][] srcArray) throws UnknownHostException {

		StringBuffer stringBuffer = new StringBuffer();
		
		boolean first = true;
		stringBuffer.append("[");
		
		//Mongo m = new Mongo( "localhost" , 27017 );  
		Mongo mongo = new Mongo();
		DB db = mongo.getDB("poetry");
		
		DBCollection dbCollection = db.getCollection(sheet.getName());
		dbCollection.drop();

		for (int i = 0; i < srcArray.length; i++) {
			String[] arrayItem = srcArray[i];
			if (!first) {
				stringBuffer.append(COMMA);
			}
			stringBuffer.append("{");
			
			DBObject doc =new BasicDBObject();

			
			boolean first2 = true;
			for (int j = 0; j < arrayItem.length; j++) {

				if (!first2) {
					stringBuffer.append(COMMA);
				}
				
				
//				System.out.println(">>>>>>>>>>>>>>>>>>>:+"+srcArray[0][j]);
				String content=StringEscapeUtils.escapeHtml(arrayItem[j]);  
				
				stringBuffer.append(QUOTATION_MARKS + srcArray[0][j]
						+ QUOTATION_MARKS + ":\"" + content
						+ QUOTATION_MARKS);
				
				doc.put(srcArray[0][j],arrayItem[j]);

				//String content = arrayItem[j];
				//content.replaceAll("/"","'");
				
				first2 = false;
			}
			stringBuffer.append("}");
			
			dbCollection.insert(doc);
			
			first = false;
			
//			System.out.println("==================================================:"+line);
			//DBObject doc = (DBObject) JSON.parse(line.toString());
//			line = new StringBuffer();

		}

		stringBuffer.append("]");
		return stringBuffer;

	}

	public static void main(String args[]) throws IOException {
		Excel2Json  ej=new Excel2Json();
		String[][] srcArray=ej.getExcelContent("I:/env_nosql/excel2json/poetry.xls",0);
		StringBuffer convertToJson = ej.convertToJson(srcArray);
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>:"+convertToJson);
		File file1 = new File("I:/env_nosql/excel2json/poetry-"+ej.sheet.getName()+".json");
		FileUtils.writeStringToFile(file1, convertToJson.toString());
		
		srcArray=ej.getExcelContent("I:/env_nosql/excel2json/poetry.xls",1);
		convertToJson = ej.convertToJson(srcArray);
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>:"+convertToJson);
		File file2 = new File("I:/env_nosql/excel2json/poetry-"+ej.sheet.getName()+".json");
		FileUtils.writeStringToFile(file2, convertToJson.toString());

		srcArray=ej.getExcelContent("I:/env_nosql/excel2json/poetry.xls",2);
		convertToJson = ej.convertToJson(srcArray);
		System.out.println(">>>>>>>>>>>>>>>>>>>>>>:"+convertToJson);
		File file3 = new File("I:/env_nosql/excel2json/poetry-"+ej.sheet.getName()+".json");
		FileUtils.writeStringToFile(file3, convertToJson.toString());
		
//		Mongo mongo = new Mongo();
//		DB db = mongo.getDB("poetry");
//		DBCollection dbCollection = db.getCollection("authors");
//		DBObject doc = new BasicDBObject();
//		dbCollection.insert(doc);

		
		
//		try {
//			Workbook book = Workbook.getWorkbook(new File("I:/env_nosql/excel2json/poetry.xls"));
//			// 获得第一个工作表对象
//			Sheet sheet = book.getSheet(0);
//			// 得到第一列第一行的单元格
//			Cell cell1 = sheet.getCell(0, 0);
//			String result = cell1.getContents();
//			System.out.println(result);
//			book.close();
//		} catch (Exception e) {
//			System.out.println(e);
//		}
	}

}
