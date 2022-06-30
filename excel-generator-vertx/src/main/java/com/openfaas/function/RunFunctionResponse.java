
// import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.*;

public class RunFunctionResponse {

	public Map<String, Object> dataMap = new LinkedHashMap<>();
	public List<Object> dataList = new ArrayList<>();
	public Set<String> noCredentialsDB = new HashSet<>();
	public Set<String> insufficientPermission = new HashSet<>();
	public Set<String> invalidCredentials = new HashSet<>();
	public Set<String> generalException = new HashSet<>();
	public Set<String> resourceTags = new HashSet<>();
	public String self;
	// public NextRequest next = new NextRequest();
	// public NextRequest previous = new NextRequest();
	public String cloud;
	public String invocationType;
	public String module;
	public String page;
	public String insight;
	public String insightText;
	public boolean referS3UrlBydefault;
	public Integer basicLimit;
	public Integer filterLimit;
	public String insightState;
	public Boolean autoRenew;
	public List<Object> otherException = new ArrayList<>();
	public List<String> restrictedFilters = new ArrayList<>();
	public String message;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String customerId;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String email;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String requestId;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String requestTime;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String ownerCustomerId;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public Boolean invalidate;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String format;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String triggered_by;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public String marks;

	// @JsonInclude(JsonInclude.Include.NON_NULL)
	// public List<Object> compliance;

	// public void addNoCredentialsDB(String accountName, String accountNumber) {
	// 	noCredentialsDB.add(accountName + "( " + accountNumber + " )");
	// }

	// public void addInsufficientPermission(String accountName, String accountNumber) {
	// 	insufficientPermission.add(accountName + "( " + accountNumber + " )");
	// }

	// public void addInvalidCredentials(String accountName, String accountNumber) {
	// 	invalidCredentials.add(accountName + "( " + accountNumber + " )");
	// }

	// public void addGeneralException(String accountName, String accountNumber) {
	// 	generalException.add(accountName + "( " + accountNumber + " )");
	// }

	// public String appendNameAndNumber(String accountName, String accountNumber) {
	// 	return accountName + "( " + accountNumber + " )";

	// }

	// public void addOtherException(String accountName, String accountNumber, String message) {
	// 	Map<String, String> errorMap = new HashMap<String, String>();
	// 	errorMap.put(accountName + "( " + accountNumber + " )", message);
	// 	otherException.add(errorMap);
	// }

}
