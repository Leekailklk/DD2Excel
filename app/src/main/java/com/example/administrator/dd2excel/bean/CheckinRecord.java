package com.example.administrator.dd2excel.bean;

import android.support.annotation.NonNull;
import com.example.administrator.dd2excel.util.Util;

import java.io.Serializable;
import java.text.Collator;
import java.util.Locale;

public class CheckinRecord implements Serializable,Comparable<CheckinRecord> {
    private String name;
    private String timestamp;
    private String detailPlace;
    private String[] imageList;
    private String latitude;

    public CheckinRecord(boolean use_it) {
        this.use_it = use_it;
    }

    private String longitude;
    private String place;
    private String remark;
    private String userId;
    private boolean use_it=true;
    @Override
    public int compareTo(@NonNull CheckinRecord checkinRecord) {
        Collator collator=Collator.getInstance(Locale.CHINA);
        return collator.compare(this.name,checkinRecord.name);
    }

    @Override
    public String toString() {
        return  "    姓名    ：" + name + '\n' +
                "签到时间：" + timestamp + '\n' +
                "详细地址：" + detailPlace;
    }

    public boolean isUse_it() {
        return use_it;
    }

    public void setUse_it(boolean use_it) {
        this.use_it = use_it;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setTimestamp(String timestamp) {
        this.timestamp = timestamp;
    }

    public void setDetailPlace(String detailPlace) {
        this.detailPlace = detailPlace;
    }

    public void setImageList(String[] imageList) {
        this.imageList = imageList;
    }

    public void setLatitude(String latitude) {
        this.latitude = latitude;
    }

    public void setLongitude(String longitude) {
        this.longitude = longitude;
    }

    public void setPlace(String place) {
        this.place = place;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }

    public void setUserId(String userId) {
        this.userId = userId;
    }

    public String getName() {
        return name;
    }

    public String getTimestamp() {
        return timestamp;
    }

    public String getDetailPlace() {
        return detailPlace;
    }

    public String[] getImageList() {
        return imageList;
    }

    public String getLatitude() {
        return latitude;
    }

    public String getLongitude() {
        return longitude;
    }

    public String getPlace() {
        return place;
    }

    public String getRemark() {
        return remark;
    }

    public String getUserId() {
        return userId;
    }

    public CheckinRecord(String name, String timestamp, String detailPlace) {
        this.name = name;
        this.timestamp = timestamp;
        this.detailPlace = detailPlace;
    }

    public CheckinRecord(String name, String timestamp, String detailPlace, String[] imageList, String latitude, String longitude, String place, String remark, String userId) {
        this.name = name;
        this.timestamp = timestamp;
        this.detailPlace = detailPlace;
        this.imageList = imageList;
        this.latitude = latitude;
        this.longitude = longitude;
        this.place = place;
        this.remark = remark;
        this.userId = userId;
    }
}
