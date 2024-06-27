package ae.gcg.plugins.sharepoint.fileupload.util;

import okhttp3.*;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.JDOMException;
import org.jdom2.input.SAXBuilder;
import org.joget.commons.util.LogUtil;
import org.json.JSONObject;

import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.nio.file.Files;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SharepointAPIHelper {
    private final MediaType JSON = MediaType.get("application/json; charset=utf-8");

    public String uploadFileToSharePoint(String applicationId, String tenantName, String clientId, String clientSecret, String refreshToken, String tenantId, String siteName, String folderName, String fileName, File file, String MOMId) throws IOException {
        OkHttpClient client = new OkHttpClient();

        // Prepare the request body with the binary file content
        byte[] fileContent = Files.readAllBytes(file.toPath());
        RequestBody body = RequestBody.create(fileContent, MediaType.parse("application/octet-stream"));

        String accessToken = getAccessToken(applicationId, tenantName, tenantId, clientId, clientSecret, refreshToken);
        LogUtil.info("<- Generated Access Token -->", accessToken);

        String formDigest = getFormDigestValue(tenantName, siteName, accessToken);
        LogUtil.info("<- Generated Form Digest Value: -->", formDigest);


        // Build the request with necessary headers
        Request request = new Request.Builder()
                .url(buildSharePointFileAddUrl(tenantName, siteName, folderName, fileName))
                .post(body)
                .addHeader("Accept", "application/json;odata=nometadata")
                .addHeader("Content-Type", "application/octet-stream")
                .addHeader("X-RequestDigest", formDigest)
                .addHeader("Authorization", "Bearer " + accessToken)
                .build();

        // Execute the request and handle the response
        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                throw new IOException("Unexpected code " + response);
            }

            // Parse the JSON response to extract the UniqueId
            String jsonResponse = response.body().string();
            JSONObject jsonObject = new JSONObject(jsonResponse);
            String uniqueId = jsonObject.getString("UniqueId"); // Get File Unique ID


            // Get ListItem Row Number & List ID
            String getFileRowAPIUrl = getFileListItemAllFieldsURL(tenantName, siteName, uniqueId); // Get FieldItems API URL
            String rowResponseXml = getApiXmlResponse(getFileRowAPIUrl, accessToken); // API CALL 2
            String listItemId = getUpdateColumnAPIURLFromResponse(rowResponseXml); // Get Update Column API From Response


            // Update MOM-ID column for the File
            updateListItemColumn(tenantName, siteName, listItemId, MOMId, accessToken); // API CALL 3: Execute Update Column API


            return uniqueId;

        } catch (JDOMException e) {
            throw new RuntimeException(e);
        }
    }


    // Method to get XML response from SharePoint API
    public String getApiXmlResponse(String apiUrl, String accessToken) throws IOException {
        OkHttpClient client = new OkHttpClient();
        // Build the request with necessary headers
        Request request = new Request.Builder()
                .url(apiUrl)
                .addHeader("Accept", "application/atom+xml")
                .addHeader("Authorization", "Bearer " + accessToken)
                .build();

        // Execute the request
        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                throw new IOException("Unexpected code " + response);
            }
            return response.body().string(); // Return the XML response as a string
        }
    }


    // Method to update the MOM-ID column
    private void updateListItemColumn(String tenantName, String siteName, String listItemId, String momId, String accessToken) throws IOException {
        String url = "https://" + tenantName + ".sharepoint.com/sites/" + siteName + "/_api/" + listItemId;

        OkHttpClient client = new OkHttpClient();
        // Prepare the JSON body for the request
        MediaType JSON = MediaType.parse("application/json;odata=verbose; charset=utf-8");
        JSONObject jsonObject = new JSONObject();
        jsonObject.put("__metadata", new JSONObject().put("type", "SP.Data.Shared_x0020_DocumentsItem"));
        jsonObject.put("MOM_x002d_ID", momId);
        LogUtil.info("Setting MOM-ID: ", jsonObject.toString());

        RequestBody body = RequestBody.create(jsonObject.toString(), JSON);

        // Build the request for updating the list item
        Request request = new Request.Builder()
                .url(url)
                .addHeader("Content-Type", "application/json;odata=verbose")
                .addHeader("IF-MATCH", "*")
                .addHeader("X-HTTP-Method", "MERGE")
                .addHeader("Authorization", "Bearer " + accessToken)
                .post(body)
                .build();

        // Execute the request
        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                LogUtil.info("Document Uploaded Successfully, but failed to set MOM-ID", response.body().string());
                //throw new IOException("Failed to update list item: " + response);
            }
        }
    }

    public String getUpdateColumnAPIURLFromResponse(String xmlContent) throws JDOMException, IOException {
        SAXBuilder saxBuilder = new SAXBuilder();
        Document document = saxBuilder.build(new StringReader(xmlContent));
        Element rootElement = document.getRootElement();
        List<Element> links = rootElement.getChildren("link", rootElement.getNamespace());

        Pattern pattern = Pattern.compile("Web/Lists\\(guid'.+?'\\)/Items\\(\\d+\\)");

        for (Element link : links) {
            String href = link.getAttributeValue("href");
            if (href != null) {
                Matcher matcher = pattern.matcher(href);
                if (matcher.find()) {
                    LogUtil.info("Column Update API: ", matcher.group());
                    return matcher.group(); // Returns only the specific part of the href
                }
            }
        }
        return null; // Return null if no matching href is found
    }

    private String getFormDigestURL(String tenantName, String siteName) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName + "/_api/contextinfo";
    }

    private String getFileListItemAllFieldsURL(String tenantName, String siteName, String fileId) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName + "/_api/Web/GetFileById('" + fileId + "')/ListItemAllFields";
    }


    private String getAccessToken(String applicationId, String tenantName, String tenantId, String clientId, String clientSecret, String refreshToken) throws IOException {
        LogUtil.info("", "<- Start Get Access Token -->");
        OkHttpClient client = new OkHttpClient();

        String url = getAccessTokenURL(tenantId);

        MultipartBody requestBody = new MultipartBody.Builder()
                .setType(MultipartBody.FORM)
                .addFormDataPart("client_id", getClientIdFormattedString(clientId, tenantId))
                .addFormDataPart("client_secret", clientSecret)
                .addFormDataPart("resource", getResourceFormattedString(applicationId, tenantName, tenantId))
                .addFormDataPart("grant_type", "refresh_token")
                .addFormDataPart("refresh_token", refreshToken)
                .build();

        Request request = new Request.Builder()
                .url(url)
                .post(requestBody)
                .build();

        try (Response response = client.newCall(request).execute()) {
            // Reading the response body
            String responseBody = response.body() != null ? response.body().string() : "null";
            LogUtil.info("", "Response Status Code: " + response.code());
            LogUtil.info("", "Response Body: " + responseBody);

            if (!response.isSuccessful()) {
                LogUtil.info(response.body().toString(), "");
                throw new IOException("Unexpected code " + response.code() + " with body " + responseBody);
            }

            JSONObject jsonObject = new JSONObject(responseBody);
            return jsonObject.getString("access_token");
        }
    }

    public String getFormDigestValue(String tenantName, String siteName, String accessToken) throws IOException {
        OkHttpClient client = new OkHttpClient();

        String url = getFormDigestURL(tenantName, siteName);
        RequestBody body = RequestBody.create("", JSON); // Empty POST body
        Request request = new Request.Builder()
                .url(url)
                .addHeader("Authorization", "Bearer " + accessToken)
                .addHeader("Accept", "application/json;odata=nometadata")
                .post(body)
                .build();

        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                throw new IOException("Unexpected code " + response);
            }

            String responseData = response.body().string();
            LogUtil.info("Response Data From Form Digest: ", responseData);
            JSONObject jsonObject = new JSONObject(responseData);
            return jsonObject.getString("FormDigestValue");
        }
    }

    private String getAccessTokenURL(String tenantId) {
        return "https://accounts.accesscontrol.windows.net/" + tenantId + "/tokens/OAuth/2";
    }

    private String getClientIdFormattedString(String clientId, String tenantId) {
        return clientId + "@" + tenantId;
    }

    private String getResourceFormattedString(String applicationId, String tenantName, String tenantId) {
        return applicationId + "/" + tenantName + ".sharepoint.com@" + tenantId;
    }

    private String buildSharePointFileAddUrl(String tenantName, String siteName, String folderName, String fileName) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName +
                "/_api/web/GetFolderByServerRelativeUrl('/sites/" + siteName + "/Shared Documents/" + folderName +
                "')/Files/add(url='" + fileName + "',overwrite=" + true + ")";
    }

    public Response downloadFileFromSharePoint(String applicationId, String tenantName, String clientId, String clientSecret, String refreshToken, String tenantId, String siteName, String documentID) throws IOException {
        OkHttpClient client = new OkHttpClient();
        String accessToken = getAccessToken(applicationId, tenantName, tenantId, clientId, clientSecret, refreshToken);
        LogUtil.info("<- Generated Access Token -->", accessToken);


        // Build the download URL
        String url = buildFileAccessURL(tenantName, siteName, documentID);
        LogUtil.info("Download URL: ", url);


        // Build the request with necessary headers
        Request request = new Request.Builder()
                .url(url)
                .addHeader("Accept", "application/octet-stream")
                .addHeader("Authorization", "Bearer " + accessToken)
                .build();


        // Execute the request and handle the response
        Response response = client.newCall(request).execute();
            if (!response.isSuccessful()) {
                LogUtil.info("Download Response: ", response.body().toString());
                throw new IOException("Unexpected code " + response);
            }

            return response;

    }


    public String buildFileAccessURL(String tenantName, String siteName, String fileId) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName + "/_api/Web/GetFileById('" + fileId + "')/$value";
    }


}
