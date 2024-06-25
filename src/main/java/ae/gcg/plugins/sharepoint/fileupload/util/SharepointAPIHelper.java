package ae.gcg.plugins.sharepoint.fileupload.util;

import com.google.gson.JsonObject;
import jdk.jpackage.internal.Log;
import okhttp3.*;
import org.joget.commons.util.LogUtil;
import org.json.JSONObject;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class SharepointAPIHelper {
    private final MediaType JSON = MediaType.get("application/json; charset=utf-8");

    public String uploadFileToSharePoint(String applicationId, String tenantName, String clientId, String clientSecret, String refreshToken, String tenantId, String siteName, String folderName, String fileName, File file) throws IOException {
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
                .url(buildSharePointFileAddUrl(tenantName, siteName, folderName, fileName, true))
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
            String uniqueId = jsonObject.getString("UniqueId");
            return uniqueId;
        }
    }


    private String getFormDigestURL(String tenantName, String siteName) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName + "/_api/contextinfo";
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

    private String buildSharePointFileAddUrl(String tenantName, String siteName, String folderName, String fileName, boolean overwrite) {
        return "https://" + tenantName + ".sharepoint.com/sites/" + siteName +
                "/_api/web/GetFolderByServerRelativeUrl('/sites/" + siteName + "/Shared Documents/" + folderName +
                "')/Files/add(url='" + fileName + "',overwrite=" + overwrite + ")";
    }

    public Response downloadFileFromSharePoint(String applicationId, String tenantName, String clientId, String clientSecret, String refreshToken, String tenantId, String siteName, String folderName, String documentID) throws IOException {
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
