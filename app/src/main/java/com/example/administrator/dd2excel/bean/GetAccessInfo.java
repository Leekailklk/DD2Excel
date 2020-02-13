package com.example.administrator.dd2excel.bean;

public class GetAccessInfo {
    private String errCode;
    private String errMsg;
    private String access_token;
    private String expires_in;

    public GetAccessInfo(String errCode, String errMsg, String access_token, String expires_in) {
        this.errCode = errCode;
        this.errMsg = errMsg;
        this.access_token = access_token;
        this.expires_in = expires_in;
    }

    public String getErrCode() {
        return errCode;
    }

    public void setErrCode(String errCode) {
        this.errCode = errCode;
    }

    public String getErrMsg() {
        return errMsg;
    }

    public void setErrMsg(String errMsg) {
        this.errMsg = errMsg;
    }

    public String getAccess_token() {
        return access_token;
    }

    public void setAccess_token(String access_token) {
        this.access_token = access_token;
    }

    public String getExpires_in() {
        return expires_in;
    }

    public void setExpires_in(String expires_in) {
        this.expires_in = expires_in;
    }

    @Override
    public String toString() {
        return "GetAccessInfo{" +
                "errCode='" + errCode + '\'' +
                ", errMsg='" + errMsg + '\'' +
                ", access_token='" + access_token + '\'' +
                ", expires_in='" + expires_in + '\'' +
                '}';
    }
}
