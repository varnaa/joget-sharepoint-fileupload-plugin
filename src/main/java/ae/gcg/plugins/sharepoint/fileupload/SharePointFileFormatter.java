package ae.gcg.plugins.sharepoint.fileupload;

import java.util.ArrayList;
import java.util.List;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.DataList;
import org.joget.apps.datalist.model.DataListColumn;
import org.joget.apps.datalist.model.DataListColumnFormatDefault;
import org.joget.commons.util.SecurityUtil;
import org.joget.commons.util.StringUtil;
import org.joget.workflow.util.WorkflowUtil;
import org.json.JSONObject;

public class SharePointFileFormatter extends DataListColumnFormatDefault {

    private final static String MESSAGE_PATH = "messages/SharePointFileDownloadFormatter";

    @Override
    public String format(DataList dataList, DataListColumn column, Object row, Object value) {
        StringBuilder result = new StringBuilder();
        if (value != null) {
            String[] values = value.toString().split(";");
            List<String> results = new ArrayList<String>();

            AppDefinition appDef = AppUtil.getCurrentAppDefinition();
            String appId = "";
            String appVersion = "";
            if (appDef != null) {
                appId = appDef.getId();
                appVersion = appDef.getVersion().toString();
            }

            String enableDownload = getPropertyString("enableDownload");
            String applicationId = getPropertyString("applicationId");
            String clientId = getPropertyString("clientId");
            String clientSecret = getPropertyString("clientSecret");
            clientSecret = SecurityUtil.decrypt(clientSecret);
            String refreshToken = getPropertyString("refreshToken");
            refreshToken = SecurityUtil.decrypt(refreshToken);
            String tenantName = getPropertyString("tenantName");
            String tenantId = getPropertyString("tenantId");
            String siteName = getPropertyString("siteName");
            String folderName = getPropertyString("folderName");

            JSONObject jsonParams = new JSONObject();

            jsonParams.put("applicationId", applicationId);
            jsonParams.put("clientId", clientId);
            jsonParams.put("clientSecret", clientSecret);
            jsonParams.put("refreshToken", refreshToken);
            jsonParams.put("tenantName", tenantName);
            jsonParams.put("tenantId", tenantId);
            jsonParams.put("siteName", siteName);
            jsonParams.put("folderName", folderName);

            for (String v : values) {
                if (v != null && !v.isEmpty() && v.indexOf('|') != -1) {
                    String[] verticalBarSplit = v.split("\\|"); // filename.pdf|documentID -> ["filenameName.pdf", "documentId"]
                    if (verticalBarSplit.length > 0) {
                        String filename = verticalBarSplit[0];
                        String documentId = verticalBarSplit[1];
                        jsonParams.put("fileName", filename);
                        String params = StringUtil.escapeString(SecurityUtil.encrypt(jsonParams.toString()), StringUtil.TYPE_URL, null);

                        if ("true".equalsIgnoreCase(enableDownload)) {
                            String filePath = WorkflowUtil.getHttpServletRequest().getContextPath() + "/web/json/app/" + appId + "/" + appVersion + "/plugin/ae.gcg.plugins.sharepoint.fileupload.SharePointFileUpload/service?dID=" + documentId + "&action=download&params=" + params;
                            String downloadUrl = "<a href=\"" + filePath + "\" target=\"_blank\">" + filename + "</a>";
                            result.append(downloadUrl);
                        } else {
                            result.append(filename);
                        }
                        result.append(";");
                    }

                } else {
                    result.append(v);
                }

            }
            if (result.length() > 0) {
                result.deleteCharAt(result.length() - 1);
            }
        }
        return result.toString();
    }

    @Override
    public String getName() {
        return "Sharepoint File Formatter";
    }

    @Override
    public String getVersion() {
        return "8.0";
    }

    @Override
    public String getDescription() {
        return "Format filename and download file from Sharepoint inside the datalist";
    }

    @Override
    public String getLabel() {
        return "Sharepoint  File Formatter";
    }

    @Override
    public String getClassName() {
        return this.getClass().getName();
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/SharePointFileDownloadFormatter.json", null, true, MESSAGE_PATH);
    }

}
