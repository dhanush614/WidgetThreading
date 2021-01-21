import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.Date;
import java.util.Iterator;

import java.io.IOException;
import java.io.InputStream;

import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.DocumentSet;
import com.filenet.api.constants.ClassNames;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.core.Connection;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.util.UserContext;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;

public class MainClass {

	static UserContext old = null;
	static CaseMgmtContext oldCmc = null;
	static String TOS = "tos";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			ConnectionClass connectionClass = new ConnectionClass();
			Connection conn = connectionClass.getConnection();
			Domain domain = Factory.Domain.fetchInstance(conn, null, null);
			System.out.println("Domain: " + domain.get_Name());
			System.out.println("Connection to Content Platform Engine successful");
			ObjectStore targetOS = (ObjectStore) domain.fetchObject(ClassNames.OBJECT_STORE, TOS, null);
			System.out.println("Object Store =" + targetOS.get_DisplayName());
			SimpleVWSessionCache vwSessCache = new SimpleVWSessionCache();
			CaseMgmtContext cmc = new CaseMgmtContext(vwSessCache, new SimpleP8ConnectionCache());
			oldCmc = CaseMgmtContext.set(cmc);

			ExecutorService threadExecutor = Executors.newFixedThreadPool(1);
			List<Future<HashMap<Integer, HashMap<String, Object>>>> list = new ArrayList<Future<HashMap<Integer, HashMap<String, Object>>>>();

			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			String folderPath = "/Bulk Case Creation";
			Folder myFolder = Factory.Folder.fetchInstance(targetOS, folderPath, null);
			DocumentSet myLoanDocs = myFolder.get_ContainedDocuments();
			Iterator itr = myLoanDocs.iterator();
			while (itr.hasNext()) {
				Document doc = (Document) itr.next();
				doc.fetchProperties(pf);
				SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
				String docCheckInDate = formatter.format(doc.get_DateCheckedIn());
				String todayDate = formatter.format(new Date());
				if (docCheckInDate.equals(todayDate)) {
					HashMap<Integer, HashMap<String, Object>> excelRows = readExcelRows(targetOS, doc);
					Future<HashMap<Integer, HashMap<String, Object>>> threadList = threadExecutor
							.submit(new ThreadClass(excelRows, targetOS, doc.get_Name()));
					list.add(threadList);
					/*
					 * for (Future<HashMap<Integer, HashMap<String, Object>>>
					 * object : list) { try { HashMap<Integer, HashMap<String,
					 * Object>> map = object.get(); System.out.println(map); }
					 * catch (InterruptedException e) { // TODO Auto-generated
					 * catch block e.printStackTrace(); } catch
					 * (ExecutionException e) { // TODO Auto-generated catch
					 * block e.printStackTrace(); } }
					 */
					threadExecutor.shutdown();
				} else {
					System.out.println("No Templates Available, Please upload template and try again..!!");
				}
			}

		} catch (Exception e) {
			System.out.println(e);
		} finally {
			if (oldCmc != null) {
				CaseMgmtContext.set(oldCmc);
			}

			if (old != null) {
				UserContext.set(old);
			}
		}

	}

	public static HashMap<Integer, HashMap<String, Object>> readExcelRows(ObjectStore targetOS, Document doc)
			throws IOException {
		ContentElementList docContentList = doc.get_ContentElements();
		HashMap<Integer, HashMap<String, Object>> caseProperties = new HashMap<Integer, HashMap<String, Object>>();
		Iterator iter = docContentList.iterator();
		while (iter.hasNext()) {
			ContentTransfer ct = (ContentTransfer) iter.next();
			InputStream stream = ct.accessContentStream();
			int rowLastCell = 0;
			HashMap<Integer, String> headers = new HashMap<Integer, String>();
			HashMap<String, String> propDescMap = new HashMap<String, String>();
			XSSFWorkbook workbook = new XSSFWorkbook(stream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFSheet sheet1 = workbook.getSheetAt(1);
			Iterator<Row> rowIterator = sheet.iterator();
			Iterator<Row> rowIterator1 = sheet1.iterator();
			while (rowIterator1.hasNext()) {
				Row row = rowIterator1.next();
				if (row.getRowNum() > 0) {
					String key = null, value = null;
					key = row.getCell(0).getStringCellValue();
					value = row.getCell(1).getStringCellValue();
					if (key != null && value != null) {
						propDescMap.put(key, value);
					}
				}
			}
			String headerValue;
			int rowNum = 0;
			if (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				int colNum = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					headerValue = cell.getStringCellValue();
					if (headerValue.contains("*")) {
						if (headerValue.contains("datetime")) {
							headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
							headerValue += "dateField";
						} else {
							headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
						}
					}
					if (headerValue.contains("datetime")) {
						headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
						headerValue += "dateField";
					} else {
						headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
					}
					headers.put(colNum++, headerValue);
				}
				rowLastCell = row.getLastCellNum();
				Cell cell1 = row.createCell(rowLastCell, Cell.CELL_TYPE_STRING);
				if (row.getRowNum() == 0) {
					cell1.setCellValue("Status");
				}
			}
			int rowStart = sheet.getFirstRowNum() + 1;
			int rowEnd = sheet.getLastRowNum();
			for (int rowNumber = rowStart; rowNumber < rowEnd; rowNumber++) {
				Row row = sheet.getRow(rowNumber);
				if (row == null) {
					break;
				} else {
					HashMap<String, Object> rowValue = new HashMap<String, Object>();
					// Row row = rowIterator.next();
					int colNum = 0;
					for (int i = 0; i < row.getLastCellNum(); i++) {
						Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
						try {
							if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
								colNum++;
							} else {
								if (headers.get(colNum).contains("dateField")) {
									String symName = headers.get(colNum).replace("dateField", "");
									if (HSSFDateUtil.isCellDateFormatted(cell)) {
										Date date = cell.getDateCellValue();
										rowValue.put(propDescMap.get(symName), date);
										colNum++;
									}
								} else {
									rowValue.put(propDescMap.get(headers.get(colNum)), getCharValue(cell));
									colNum++;
								}
							}
						} catch (Exception e) {
							System.out.println(e);
							e.printStackTrace();
						}

					}
					caseProperties.put(++rowNum, rowValue);
				}
			}
		}
		return caseProperties;
	}

	private static Object getCharValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();

		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}

}
