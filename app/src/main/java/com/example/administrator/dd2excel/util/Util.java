package com.example.administrator.dd2excel.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

public class Util {
    public static void main(String[] args) throws ParseException {
        System.out.println(dateToStamp("2020-2-1 00:00:00"));
        System.out.println(stampToDate("1580486400000"));
        System.out.println(getCurrentTime("YYYY-MM-dd"));
        System.out.println(getCity("河北省保定市定兴县北南蔡乡北堤线"));
    }
    public static String getCity(String detail){
        int pos=detail.indexOf("市");
        return detail.substring(0,pos+1);

    }
    public static String getCurrentTime(String timeFormat){
        //SimpleDateFormat formatter = new SimpleDateFormat("YYYY-MM-dd HH:mm:ss"); //设置时间格式
        SimpleDateFormat formatter = new SimpleDateFormat(timeFormat); //设置时间格式
        formatter.setTimeZone(TimeZone.getTimeZone("GMT+08")); //设置时区
        Date curDate = new Date(System.currentTimeMillis()); //获取当前时间
        String createDate = formatter.format(curDate);   //格式转换
        return createDate;
    }
    /*
     * 将时间转换为时间戳
     */
    public static String dateToStamp(String time) throws ParseException{
        String stap;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date date = simpleDateFormat.parse(time);
        long ts = date.getTime();//获取时间的时间戳
        stap = String.valueOf(ts);
        return stap;
    }/*
    public static Long dateToStamp(String time) throws ParseException{
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date date = simpleDateFormat.parse(time);
        return date.getTime();//获取时间的时间戳
            }*/
    /*
     * 将时间戳转换为时间
     */
    public static String stampToDate(String s){
        String res;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        long lt = new Long(s);
        Date date = new Date(lt);
        res = simpleDateFormat.format(date);
        return res;
    }/*
     public static String stampToDate(Long stap){
        String time;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        long lt = stap;
        Date date = new Date(lt);
        time = simpleDateFormat.format(date);
        return time;
    }*/
}