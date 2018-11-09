package com.demo.springboot.controller;

/**
 * Created by csh on 2018/8/23.
 */
public class Substation {
    private String subNo;
    private String markerNo;
    private String longitude;
    private String latitude;

    public String getSubNo() {
        return subNo;
    }

    public void setSubNo(String subNo) {
        this.subNo = subNo;
    }

    public String getMarkerNo() {
        return markerNo;
    }

    public void setMarkerNo(String markerNo) {
        this.markerNo = markerNo;
    }

    public String getLongitude() {
        return longitude;
    }

    public void setLongitude(String longitude) {
        this.longitude = longitude;
    }

    public String getLatitude() {
        return latitude;
    }

    public void setLatitude(String latitude) {
        this.latitude = latitude;
    }
}
