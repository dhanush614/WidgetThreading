import java.util.HashMap;
import java.util.Iterator;
import java.util.Map.Entry;
import java.util.concurrent.Callable;

import com.filenet.api.constants.ClassNames;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.core.Connection;
import com.filenet.api.core.Domain;
import com.filenet.api.core.Factory;
import com.filenet.api.core.ObjectStore;
import com.ibm.casemgmt.api.Case;
import com.ibm.casemgmt.api.CaseType;
import com.ibm.casemgmt.api.constants.ModificationIntent;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;

public class ThreadClass implements Callable<HashMap<Integer, HashMap<String, Object>>> {
	HashMap<String, Object> caseProperties;
	int rowNumber;
	String casetypeName;

	public ThreadClass(int rowNumber, HashMap<String, Object> hashMap, String casetypeName) {
		super();
		this.rowNumber = rowNumber;
		this.caseProperties = hashMap;
		this.casetypeName = casetypeName;
	}

	@Override
	public HashMap<Integer, HashMap<String, Object>> call() throws Exception {
		// TODO Auto-generated method stub
		CaseMgmtContext oldCmc = null;
		ConnectionClass connectionClass = new ConnectionClass();
		Connection conn = connectionClass.getConnection();
		Domain domain = Factory.Domain.fetchInstance(conn, null, null);
		HashMap<Integer, HashMap<String, Object>> responseMap = new HashMap<Integer, HashMap<String, Object>>();
		//System.out.println("Domain: " + domain.get_Name());
		//System.out.println("Connection to Content Platform Engine successful");
		ObjectStore targetOS = (ObjectStore) domain.fetchObject(ClassNames.OBJECT_STORE, "tos", null);
		//System.out.println("Object Store =" + targetOS.get_DisplayName());
		SimpleVWSessionCache vwSessCache = new SimpleVWSessionCache();
		CaseMgmtContext cmc = new CaseMgmtContext(vwSessCache, new SimpleP8ConnectionCache());
		oldCmc = CaseMgmtContext.set(cmc);
		String caseId = "";
		HashMap<String, Object> rowValue = new HashMap<String, Object>();
		try {
			Case pendingCase = null;
			//System.out.println("RowNumber :   " + rowNumber);
			ObjectStoreReference targetOsRef = new ObjectStoreReference(targetOS);
			CaseType caseType = CaseType.fetchInstance(targetOsRef, casetypeName);
			pendingCase = Case.createPendingInstance(caseType);
			Iterator<Entry<String, Object>> propertyValues = caseProperties.entrySet().iterator();
			while (propertyValues.hasNext()) {
				Entry<String, Object> propertyValuesPair = propertyValues.next();
				rowValue.put(propertyValuesPair.getKey(), propertyValuesPair.getValue());
				pendingCase.getProperties().putObjectValue(propertyValuesPair.getKey(), propertyValuesPair.getValue());
				propertyValues.remove();
			}
			pendingCase.save(RefreshMode.REFRESH, null, ModificationIntent.MODIFY);
			caseId = pendingCase.getId().toString();
			//System.out.println("Case_ID: " + caseId);
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
		}
		if (!caseId.isEmpty()) {
			rowValue.put("Status", "Success");
			responseMap.put(rowNumber, rowValue);
		} else {
			rowValue.put("Status", "Failure");
			responseMap.put(rowNumber, rowValue);
		}
		return responseMap;
	}
}
