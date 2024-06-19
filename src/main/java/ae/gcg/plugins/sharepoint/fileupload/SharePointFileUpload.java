package ae.gcg.plugins.sharepoint.fileupload;

import ae.gcg.plugins.sharepoint.fileupload.util.SharepointAPIHelper;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.model.*;
import org.joget.apps.form.service.FormUtil;
import org.joget.apps.userview.model.PwaOfflineResources;
import org.joget.commons.util.*;
import org.joget.plugin.base.PluginWebSupport;
import org.joget.workflow.util.WorkflowUtil;
import org.json.JSONObject;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;

public class SharePointFileUpload extends Element implements FormBuilderPaletteElement, FileDownloadSecurity, PluginWebSupport, PwaOfflineResources {
    private final static String MESSAGE_PATH = "messages/SharePointFileUpload";

    @Override
    public String getName() {
        return AppPluginUtil.getMessage("ae.gcg.plugins.sharepoint.fileupload.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getVersion() {
        return "8.0.0";
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("ae.gcg.plugins.sharepoint.fileupload.pluginDescription", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getFormBuilderCategory() {
        return "Marketplace";
    }

    @Override
    public int getFormBuilderPosition() {
        return 900;
    }

    @Override
    public String getFormBuilderIcon() {
        return "<i class=\"fas fa-upload\"></i>";
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("ae.gcg.plugins.sharepoint.fileupload.pluginLabel", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/SharePointFileUpload.json", null, true, MESSAGE_PATH);
    }

    @Override
    public boolean isDownloadAllowed(Map map) {
        return true;
    }

    @Override
    public String getFormBuilderTemplate() {
        return null;
    }

    public static Map<String, String> parseFileName(String input) {
        Map<String, String> resultMap = new HashMap<>();

        // Split the input based on "|"
        String[] parts = input.split("\\|");

        if (parts.length == 2) {
            // Extract the filename (part before "|")
            String filename = parts[0].trim();
            String documentId = parts[1].trim();

            resultMap.put("filename", filename);
            resultMap.put("documentId", documentId);
        } else {
            System.err.println("Invalid input format.");
        }

        return resultMap;
    }

    @Override
    public Set<String> getOfflineStaticResources() {
        Set<String> urls = new HashSet<>();
        String contextPath = AppUtil.getRequestContextPath();
        urls.add(contextPath + "/js/dropzone/dropzone.css");
        urls.add(contextPath + "/js/dropzone/dropzone.js");
        urls.add(contextPath + "/plugin/org.joget.apps.form.lib.FileUpload/js/jquery.fileupload.js");

        return urls;
    }

    @Override
    public Boolean selfValidate(FormData formData) {
        String id = FormUtil.getElementParameterName(this);
        Boolean valid = true;
        String error = "";
        try {
            String[] values = FormUtil.getElementPropertyValues(this, formData);

            for (String value : values) {
                File file = FileManager.getFileByPath(value);
                if (file != null) {
                    if (getPropertyString("maxSize") != null && !getPropertyString("maxSize").isEmpty()) {
                        long maxSize = Long.parseLong(getPropertyString("maxSize")) * 1024;

                        if (file.length() > maxSize) {
                            valid = false;
                            error += getPropertyString("maxSizeMsg") + " ";

                        }
                    }
                    if (getPropertyString("fileType") != null && !getPropertyString("fileType").isEmpty()) {
                        String[] fileType = getPropertyString("fileType").split(";");
                        String filename = file.getName().toUpperCase();
                        Boolean found = false;
                        for (String type : fileType) {
                            if (filename.endsWith(type.toUpperCase())) {
                                found = true;
                            }
                        }
                        if (!found) {
                            valid = false;
                            error += getPropertyString("fileTypeMsg");
                            FileManager.deleteFile(file);
                        }
                    }
                }
            }

            if (!valid) {
                formData.addFormError(id, error);
            }
        } catch (Exception e) {
        }

        return valid;
    }


    @Override
    public FormData formatDataForValidation(FormData formData) {
        String filePathPostfix = "_path";
        String id = FormUtil.getElementParameterName(this);
        if (id != null) {
            String[] tempFilenames = formData.getRequestParameterValues(id);
            String[] tempExisting = formData.getRequestParameterValues(id + filePathPostfix);

            List<String> filenames = new ArrayList<>();
            if (tempFilenames != null && tempFilenames.length > 0) {
                filenames.addAll(Arrays.asList(tempFilenames));
            }

            if (tempExisting != null && tempExisting.length > 0) {
                filenames.addAll(Arrays.asList(tempExisting));
            }

            if (filenames.isEmpty()) {
                formData.addRequestParameterValues(id, new String[]{""});
            } else if (!"true".equals(getPropertyString("multiple"))) {
                formData.addRequestParameterValues(id, new String[]{filenames.get(0)});
            } else {
                formData.addRequestParameterValues(id, filenames.toArray(new String[]{}));
            }
        }
        return formData;
    }

    public String getServiceUrl() {
        String url = WorkflowUtil.getHttpServletRequest().getContextPath() + "/web/json/plugin/ae.gcg.plugins.sharepoint.fileupload.SharePointFileUpload/service";
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        //create nonce
        String paramName = FormUtil.getElementParameterName(this);
        String fileType = getPropertyString("fileType");
        String nonce = SecurityUtil.generateNonce(new String[]{"FileUpload", appDef.getAppId(), appDef.getVersion().toString(), paramName, fileType}, 1);
        try {
            url = url + "?_nonce=" + URLEncoder.encode(nonce, "UTF-8") + "&_paramName=" + URLEncoder.encode(paramName, "UTF-8") + "&_appId=" + URLEncoder.encode(appDef.getAppId(), "UTF-8") + "&_appVersion=" + URLEncoder.encode(appDef.getVersion().toString(), "UTF-8") + "&_ft=" + URLEncoder.encode(fileType, "UTF-8");
        } catch (Exception e) {
        }
        return url;
    }

    @Override
    public FormRowSet formatData(FormData formData) {
        FormRowSet rowSet = null;

        String id = getPropertyString(FormUtil.PROPERTY_ID);

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

        Set<String> remove = null;
        if ("true".equals(getPropertyString("removeFile"))) {
            remove = new HashSet<String>();
            Form form = FormUtil.findRootForm(this);
            String originalValues = formData.getLoadBinderDataProperty(form, id);
            if (originalValues != null) {
                remove.addAll(Arrays.asList(originalValues.split(";")));
            }
        }

        // get value
        if (id != null) {
            String[] values = FormUtil.getElementPropertyValues(this, formData);
            if (values != null && values.length > 0) {
                // set value into Properties and FormRowSet object
                FormRow result = new FormRow();
                List<String> resultedValue = new ArrayList<String>();
                List<String> filePaths = new ArrayList<String>();

                for (String value : values) {
                    // check if the file is in temp file
                    File file = FileManager.getFileByPath(value);

                    if (file != null) {
                        // upload file to SharePoint
                        String documentId = "";
                        try {
                            documentId = new SharepointAPIHelper().uploadFileToSharePoint(applicationId, tenantName, clientId, clientSecret, refreshToken, tenantId, siteName, folderName, file.getName(), file);
                        } catch (IOException e) {
                            // Convert stack trace to a single string
                            StringWriter sw = new StringWriter();
                            PrintWriter pw = new PrintWriter(sw);
                            e.printStackTrace(pw);
                            String stackTrace = sw.toString();

                            // Log the stack trace using your custom LogUtil
                            LogUtil.info("An Exception occurred while creating document: " + e.getMessage() , "\nStackTrace: " + stackTrace);
                        }

                        filePaths.add(value + "|" + documentId);
                        resultedValue.add(file.getName() + "|" + documentId);
                    } else {
                        if (remove != null && !value.isEmpty()) {
                            remove.removeIf(item -> {
                                if (item.contains(value)) {
                                    resultedValue.add(item);
                                    return true;
                                }
                                return false;
                            });
                        }
                    }
                }

                if (!filePaths.isEmpty()) {
                    result.putTempFilePath(id, filePaths.toArray(new String[]{}));
                }

                if (remove != null && !remove.isEmpty() && !remove.contains("")) {
                    result.putDeleteFilePath(id, remove.toArray(new String[]{}));
                    for (String r : remove) {
                        Map<String, String> fileMap = parseFileName(r);
                        String documentId = fileMap.get("documentId");

                        if (documentId != null && !documentId.isEmpty()) {
                            // delete file(s) from mayan
                            //new DoxisApiHelper().deleteDocument(serverUrl, username, password, customerName, documentId, repositoryName);
                        }
                    }
                }

                // formulate values
                String delimitedValue = FormUtil.generateElementPropertyValues(resultedValue.toArray(new String[]{}));
                String paramName = FormUtil.getElementParameterName(this);
                formData.addRequestParameterValues(paramName, resultedValue.toArray(new String[]{}));

                if (delimitedValue == null) {
                    delimitedValue = "";
                }

                // set value into Properties and FormRowSet object
                result.setProperty(id, delimitedValue);
                rowSet = new FormRowSet();
                rowSet.add(result);

                String filePathPostfix = "_path";
                formData.addRequestParameterValues(id + filePathPostfix, new String[]{});
            }
        }

        return rowSet;
    }

    @Override
    public String renderTemplate(FormData formData, Map dataModel) {
        String template = "SharePointFileUpload.ftl";

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

        // set value
        String[] values = FormUtil.getElementPropertyValues(this, formData);

        Map<String, String> tempFilePaths = new LinkedHashMap<>();
        Map<String, String> filePaths = new LinkedHashMap<>();

        String filePathPostfix = "_path";
        String id = FormUtil.getElementParameterName(this);

        //check is there a stored value
        String storedValue = formData.getStoreBinderDataProperty(this);
        if (storedValue != null) {
            values = storedValue.split(";");
        } else {
            //if there is no stored value, get the temp files
            String[] tempExisting = formData.getRequestParameterValues(id + filePathPostfix);

            if (tempExisting != null && tempExisting.length > 0) {
                values = tempExisting;
            }
        }

        Form form = FormUtil.findRootForm(this);
        if (form != null) {
            form.getPropertyString(FormUtil.PROPERTY_ID);
        }
        String appId = "";
        String appVersion = "";

        AppDefinition appDef = AppUtil.getCurrentAppDefinition();

        if (appDef != null) {
            appId = appDef.getId();
            appVersion = appDef.getVersion().toString();
        }

        for (String value : values) {
            // check if the file is in temp file

            Map<String, String> fileMap = parseFileName(value);
            value = fileMap.get("filename");
            String documentId = fileMap.get("documentId");

            File file = FileManager.getFileByPath(value);

            if (file != null) {
                tempFilePaths.put(value, file.getName());
            } else if (value != null && !value.isEmpty()) {
                // determine actual path for the file uploads
                String fileName = value;
                String encodedFileName = fileName;
                if (fileName != null) {
                    try {
                        encodedFileName = URLEncoder.encode(fileName, "UTF8").replaceAll("\\+", "%20");
                    } catch (UnsupportedEncodingException ex) {
                        // ignore
                    }
                }

                jsonParams.put("applicationId", applicationId);
                jsonParams.put("clientId", clientId);
                jsonParams.put("clientSecret", clientSecret);
                jsonParams.put("refreshToken", refreshToken);
                jsonParams.put("tenantName", tenantName);
                jsonParams.put("tenantId", tenantId);
                jsonParams.put("siteName", siteName);
                jsonParams.put("folderName", folderName);


                String params = StringUtil.escapeString(SecurityUtil.encrypt(jsonParams.toString()), StringUtil.TYPE_URL, null);
                String filePath = "/web/json/app/" + appId + "/" + appVersion + "/plugin/ae.gcg.plugins.sharepoint.fileupload.SharePointFileUpload/service?dID=" + documentId + "&action=download&params=" + params;
                filePaths.put(filePath, value);
            }
        }

        if (!tempFilePaths.isEmpty()) {
            dataModel.put("tempFilePaths", tempFilePaths);
        }
        if (!filePaths.isEmpty()) {
            dataModel.put("filePaths", filePaths);
        }

        String html = FormUtil.generateElementHtml(this, formData, template, dataModel);
        return html;
    }

    @Override
    public void webService(javax.servlet.http.HttpServletRequest request, javax.servlet.http.HttpServletResponse response) throws IOException, ServletException {
        String nonce = request.getParameter("_nonce");
        String paramName = request.getParameter("_paramName");
        String appId = request.getParameter("_appId");
        String appVersion = request.getParameter("_appVersion");
        String filePath = request.getParameter("_path");
        String fileType = request.getParameter("_ft");

        String action = request.getParameter("action");
        String documentId = request.getParameter("dID");

        if ("download".equals(action) && (documentId != null && !documentId.isEmpty())) {
           // Download yet to implement.
        }


        if (SecurityUtil.verifyNonce(nonce, new String[]{"FileUpload", appId, appVersion, paramName, fileType})) {
            if ("POST".equalsIgnoreCase(request.getMethod())) {

                try {
                    JSONObject obj = new JSONObject();
                    try {
                        // handle multipart files
                        String validatedParamName = SecurityUtil.validateStringInput(paramName);
                        MultipartFile file = FileStore.getFile(validatedParamName);
                        if (file != null && file.getOriginalFilename() != null && !file.getOriginalFilename().isEmpty()) {
                            String ext = file.getOriginalFilename().substring(file.getOriginalFilename().lastIndexOf(".")).toLowerCase();
                            if (fileType != null && (fileType.isEmpty() || fileType.contains(ext + ";") || fileType.endsWith(ext))) {
                                String path = FileManager.storeFile(file);
                                obj.put("path", path);
                                obj.put("filename", file.getOriginalFilename());
                                obj.put("newFilename", path.substring(path.lastIndexOf(File.separator) + 1));
                            } else {
                                obj.put("error", ResourceBundleUtil.getMessage("form.fileupload.fileType.msg.invalidFileType"));
                            }
                        }

                        Collection<String> errorList = FileStore.getFileErrorList();
                        if (errorList != null && !errorList.isEmpty() && errorList.contains(paramName)) {
                            obj.put("error", ResourceBundleUtil.getMessage("general.error.fileSizeTooLarge", new Object[]{FileStore.getFileSizeLimit()}));
                        }
                    } catch (Exception e) {
                        obj.put("error", e.getLocalizedMessage());
                    } finally {
                        FileStore.clear();
                    }
                    obj.write(response.getWriter());
                } catch (Exception ex) {
                    LogUtil.error(getClassName(), ex, ex.getMessage());
                }
            } else if (filePath != null && !filePath.isEmpty()) {
                String normalizedFilePath = SecurityUtil.normalizedFileName(filePath);
                File file = FileManager.getFileByPath(normalizedFilePath);
                if (file != null) {
                    ServletOutputStream stream = response.getOutputStream();
                    DataInputStream in = new DataInputStream(new FileInputStream(file));
                    byte[] bbuf = new byte[65536];

                    try {
                        String contentType = request.getSession().getServletContext().getMimeType(file.getName());
                        if (contentType != null) {
                            response.setContentType(contentType);
                        }

                        // send output
                        int length = 0;
                        while ((in != null) && ((length = in.read(bbuf)) != -1)) {
                            stream.write(bbuf, 0, length);
                        }
                    } catch (Exception e) {

                    } finally {
                        in.close();
                        stream.flush();
                        stream.close();
                    }
                } else {
                    response.sendError(HttpServletResponse.SC_NOT_FOUND);
                    return;
                }
            }
        } else {
            response.sendError(HttpServletResponse.SC_FORBIDDEN, ResourceBundleUtil.getMessage("general.error.error403"));
        }
    }


}
