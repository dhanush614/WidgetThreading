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
	HashMap<Integer, HashMap<String, Object>> caseProperties;
	ObjectStore objectStore;
	String casetypeName;

	public ThreadClass(HashMap<Integer, HashMap<String, Object>> caseProperties, ObjectStore objectStore,
			String casetypeName) {
		super();
		this.caseProperties = caseProperties;
		this.objectStore = objectStore;
		this.casetypeName = casetypeName;
	}

	@Override
	public HashMap<Integer, HashMap<String, Object>> call() throws Exception {
		// TODO Auto-generated method stub
		CaseMgmtContext oldCmc = null;
		ConnectionClass connectionClass = new ConnectionClass();
		Connection conn = connectionClass.getConnection();
		Domain domain = Factory.Domain.fetchInstance(conn, null, null);
		System.out.println("Domain: " + domain.get_Name());
		System.out.println("Connection to Content Platform Engine successful");
		ObjectStore targetOS = (ObjectStore) domain.fetchObject(ClassNames.OBJECT_STORE, "tos", null);
		System.out.println("Object Store =" + targetOS.get_DisplayName());
		SimpleVWSessionCache vwSessCache = new SimpleVWSessionCache();
		CaseMgmtContext cmc = new CaseMgmtContext(vwSessCache, new SimpleP8ConnectionCache());
		oldCmc = CaseMgmtContext.set(cmc);
		Iterator<Entry<Integer, HashMap<String, Object>>> caseProperty = caseProperties.entrySet().iterator();
		while (caseProperty.hasNext()) {
			try {
				Case pendingCase = null;
				String caseId = "";
				Entry<Integer, HashMap<String, Object>> propertyPair = caseProperty.next();
				System.out.println("RowNumber :   " + propertyPair.getKey());
				ObjectStoreReference targetOsRef = new ObjectStoreReference(targetOS);
				CaseType caseType = CaseType.fetchInstance(targetOsRef, casetypeName);
				pendingCase = Case.createPendingInstance(caseType);
				Iterator<Entry<String, Object>> propertyValues = (propertyPair.getValue()).entrySet().iterator();
				while (propertyValues.hasNext()) {
					Entry<String, Object> propertyValuesPair = propertyValues.next();
					pendingCase.getProperties().putObjectValue(propertyValuesPair.getKey(),
							propertyValuesPair.getValue());
					propertyValues.remove();
				}
				pendingCase.save(RefreshMode.REFRESH, null, ModificationIntent.MODIFY);
				caseId = pendingCase.getId().toString();
				System.out.println("Case_ID: " + caseId);
			} catch (Exception e) {
				System.out.println(e);
				e.printStackTrace();
			}
		}
		return null;
	}
}
