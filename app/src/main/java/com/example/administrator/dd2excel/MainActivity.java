package com.example.administrator.dd2excel;

import android.Manifest;
import android.annotation.TargetApi;
import android.app.Activity;
import android.content.ContentUris;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.support.v4.content.FileProvider;
import android.util.Log;
import android.view.View;
import android.widget.*;
import com.example.administrator.dd2excel.bean.CheckinRecord;
import com.example.administrator.dd2excel.util.ExcelUtils;
import com.example.administrator.dd2excel.util.Util;
import jxl.read.biff.BiffException;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static com.example.administrator.dd2excel.util.ExcelUtils.readExcel;

public class MainActivity extends Activity implements View.OnClickListener {
    private String filePath;
    private Button btn;
    private Button sendRequest;
    private Button save2file;
    private Button openfile;
    private TextView responseText;
    private EditText sourceEditText;
    private ListView listView;
    private CheckinAdapter checkinAdapter;
    private File file;
    private File newFile;
    private String fileName;
    private ArrayList<ArrayList<String>> recordList;
    private static String[] title = { "姓名","签到时间","签到地点" };

    private List<CheckinRecord> checkinList=new ArrayList<>();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        sendRequest = (Button) findViewById(R.id.send_request);
        save2file=(Button) findViewById(R.id.save2file);
        responseText = (TextView) findViewById(R.id.response_text);
        listView = (ListView) findViewById(R.id.list_view);
        btn = (Button) findViewById(R.id.btn);
        openfile=(Button)findViewById(R.id.openfile);
        openfile.setOnClickListener(this);
        btn.setOnClickListener(this);
        sendRequest.setOnClickListener(this);
        save2file.setOnClickListener(this);
        listView.setOnItemClickListener(new AdapterView.OnItemClickListener() {

            @Override
            public void onItemClick(AdapterView<?> parent, View view,
                                    int position, long id) {
                // 取得ViewHolder对象
                CheckinAdapter.ViewHolder viewHolder = (CheckinAdapter.ViewHolder) view.getTag();
// 改变CheckBox的状态
                viewHolder.useIt.toggle();
// 将CheckBox的选中状况记录下来
                checkinList.get(position).setUse_it(viewHolder.useIt.isChecked());
                responseText.setText(checkinList.get(position).toString());
                // 刷新
                checkinAdapter.notifyDataSetChanged();

            }

        });
        //初始化参数
        initParam();
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.btn:
                setParam();
                if (isGrantExternalRW(MainActivity.this)) {
                    Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                    intent.setType("application/vnd.ms-excel application/x-excel");   //打开文件类型   Excel文档
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    startActivityForResult(intent, 1);
                } else {
                    Toast.makeText(MainActivity.this, "请检查是否开启读写权限", Toast.LENGTH_LONG).show();
                }
                break;
            case R.id.send_request:
                loadFile();
                break;
            case R.id.save2file:
                exportExcel();
                break;
            case R.id.openfile:
                openFile();
                break;
            default:
                break;
        }
    }

    private void setParam() {
        file=null;
        checkinList=new ArrayList<>();
    }

    private void initParam() {
        checkinList=new ArrayList<>();
    }
    /**
     * 导入源excel记录

     */
    public void loadFile(){
        List<String[]> list = new ArrayList<String[]>();
        //filefileName = getSDPath() + "/aaa/"+sourceEditText.getText().toString();
        //Log.d("loadFile",file);

        try {
            list = readExcel(file);
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        for (int i=2;i< list.size();i++) {

            CheckinRecord rd = new CheckinRecord(true);
            rd.setName(list.get(i)[0]);
            rd.setTimestamp(list.get(i)[5]+" "+list.get(i)[6]);
            rd.setDetailPlace(list.get(i)[10]);
            Log.d("loadFile",rd.toString());
            checkinList.add(rd);
        }
        Collections.sort(checkinList);
        checkinAdapter = new CheckinAdapter(MainActivity.this,
                R.layout.checkin_record, checkinList);
        listView.setAdapter(checkinAdapter);
        responseText.setText("导入完成，共找到"+checkinList.size()+"条记录。");
        Toast.makeText(MainActivity.this, "导入完成，共找到"+checkinList.size()+"条记录。",
                Toast.LENGTH_SHORT).show();
    }
    /**
     * 导出excel

     */
    public void exportExcel() {
        newFile = new File(getSDPath() + "/aaa");
        makeDir(newFile);
        String timestamp=Util.getCurrentTime("YYYY-MM-dd");
        ExcelUtils.initExcel(newFile.toString() + "/"+timestamp+"大客户服务科.xls", title);
        fileName = getSDPath() + "/aaa/"+timestamp+"大客户服务科.xls";
        ExcelUtils.writeObjListToExcel(getRecordData(), fileName, this);
        Toast.makeText(MainActivity.this, "文件保存成功！",
                Toast.LENGTH_SHORT).show();
        responseText.setText("文件成功保存在："+"aaa/"+timestamp+"大客户服务科.xls");
    }
    private void  openFile(){

    Intent intent = new Intent();
        intent.setFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
    //设置intent的Action属性  
   intent.setAction(Intent.ACTION_VIEW);
   //获取文件file的MIME类型  
   String type = "application/vnd.ms-excel";
   Uri fileURI = FileProvider.getUriForFile(MainActivity.this, MainActivity.this.getApplicationContext().getPackageName() + ".provider", new File(fileName));

        //设置intent的data和Type属性。  
   intent.setDataAndType(fileURI,type);
   //跳转  
   startActivity(intent);

    }
    /**
     * 将数据集合 转化成ArrayList<ArrayList<String>>
     * @return
     */
    private  ArrayList<ArrayList<String>> getRecordData() {
        recordList = new ArrayList<>();
        for (int i = 0; i <checkinList.size(); i++) {
            CheckinRecord checkinRecord = checkinList.get(i);
            ArrayList<String> beanList = new ArrayList<String>();
            if(checkinRecord.isUse_it()) {
                beanList.add(checkinRecord.getName());
                beanList.add(checkinRecord.getTimestamp());
                beanList.add(Util.getCity(checkinRecord.getDetailPlace()));
                recordList.add(beanList);
            }
        }
        return recordList;
    }

    private  String getSDPath() {
        File sdDir = null;
        boolean sdCardExist = Environment.getExternalStorageState().equals(
                android.os.Environment.MEDIA_MOUNTED);
        if (sdCardExist) {
            sdDir = Environment.getExternalStorageDirectory();
        }
        String dir = sdDir.toString();
        return dir;
    }

    public  void makeDir(File dir) {
        if (!dir.getParentFile().exists()) {
            makeDir(dir.getParentFile());
        }
        dir.mkdir();
    }

    /**
     * 创建文件夹,暂时未使用
     */
    public void createFolder() {
        //获取SD卡的路径
        //String path = MyApplication.getContext().getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS).getPath();
        String Tag = "filetest";
        //getFilesDir()获取你app的内部存储空间
        File Folder = new File(Environment.
                getExternalStorageDirectory(), "aaa");

        if (!Folder.exists())//判断文件夹是否存在，不存在则创建文件夹，已经存在则跳过
        {
            Folder.mkdir();//创建文件夹
            //两种方式判断文件夹是否创建成功
            //Folder.isDirectory()返回True表示文件路径是对的，即文件创建成功，false则相反
            boolean isFilemaked1 = Folder.isDirectory();
            //Folder.mkdirs()返回true即文件创建成功，false则相反
            boolean isFilemaked2 = Folder.mkdirs();

            if (isFilemaked1 || isFilemaked2) {
                Log.i(Tag, "创建文件夹成功"+Folder.getAbsolutePath());
            } else {
                Log.i(Tag, "创建文件夹失败");
            }

        } else {
            Log.i(Tag, "文件夹已存在");
        }

    }
    //暂时未使用
    public void newFile(String _path, String _fileName) {
        File file = new File(_path + "/" + _fileName);
        try {
            if (!file.exists()) {
                file.createNewFile();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    /**
     * 解决安卓6.0以上版本不能读取外部存储权限的问题
     *
     * @param activity
     * @return
     */
    public static boolean isGrantExternalRW(Activity activity) {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M && activity.checkSelfPermission(
                Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            activity.requestPermissions(new String[]{
                    Manifest.permission.READ_EXTERNAL_STORAGE,
                    Manifest.permission.WRITE_EXTERNAL_STORAGE
            }, 1);
            return false;
        }
        return true;
    }
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (data == null) {
            return;
        }
        Uri uri = data.getData();//得到uri，后面就是将uri转化成file的过程。

        if(!uri.getPath().equals(filePath)){//判断是否第二次选择文件
            file=null;
        }

        //获取到选中文件的路径
        filePath = uri.getPath();

        //判断是否是外部打开
        if(filePath.contains("external")){
            isExternal(uri);
        }
        //获取的是否是真实路径
        if(file==null){
            isWhetherTruePath(uri);
        }
        //如果前面都获取不到文件，则自己拼接路径
        if(file==null){
            splicingPath(uri);
        }
        Log.i("hxl", "路径转化成的file========="+file);

    }
    /**
     * 拿到文件外部路径，通过外部路径遍历出真实路径
     * @param uri
     */
    private void isExternal(Uri uri){
        Log.i("hxl", "获取文件的路径filePath========="+filePath);
        Log.i("hxl", "===调用外部遍历出路径方法===");
        String[] proj = { MediaStore.Images.Media.DATA };
        Cursor actualimagecursor = this.managedQuery(uri,proj,null,null,null);
        int actual_image_column_index = actualimagecursor.getColumnIndexOrThrow(MediaStore.Images.Media.DATA);
        actualimagecursor.moveToFirst();
        String img_path = actualimagecursor.getString(actual_image_column_index);
        file = new File(img_path);
//        Log.i("hxl", "file========="+file);
        filePath=file.getAbsolutePath();
        if(!filePath.endsWith(".xls")){
            Toast.makeText(MainActivity.this, "您选中的文件不是xls格式文档", Toast.LENGTH_LONG).show();
            filePath=null;
            return;
        }

    }
    /**
     * 判断打开文件的是那种类型
     * @param uri
     */
    private void isWhetherTruePath(Uri uri){
        try {
            Log.i("hxl", "获取文件的路径filePath========="+filePath);
            if (filePath != null) {
                if (filePath.endsWith(".xls")) {
                    if ("file".equalsIgnoreCase(uri.getScheme())) {//使用第三方应用打开
                        filePath = getPath(this, uri);
                        Log.i("hxl", "===调用第三方应用打开===");
                        fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
                        file = new File(filePath);
                    }
                    if (Build.VERSION.SDK_INT > Build.VERSION_CODES.KITKAT) {//4.4以后
                        Log.i("hxl", "===调用4.4以后系统方法===");
                        filePath = getRealPathFromURI(uri);
                        fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
                        file = new File(filePath);
                    } else {//4.4以下系统调用方法
                        filePath = getRealPathFromURI(uri);
                        Log.i("hxl", "===调用4.4以下系统方法===");
                        fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
                        file = new File(filePath);
                    }
                } else {
                    Toast.makeText(MainActivity.this, "您选中的文件格式不是xls格式文档", Toast.LENGTH_LONG).show();
                }
//                Log.i("hxl", "file========="+file);
            }else{

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }




    /**
     * 专为Android4.4设计的从Uri获取文件绝对路径，以前的方法已不好使
     */
    @TargetApi(Build.VERSION_CODES.KITKAT)
    public String getPath(final Context context, final Uri uri) {

        final boolean isKitKat = Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT;

        // DocumentProvider
        if (isKitKat && DocumentsContract.isDocumentUri(context, uri)) {
            // ExternalStorageProvider
            if (isExternalStorageDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];

                if ("primary".equalsIgnoreCase(type)) {
                    return Environment.getExternalStorageDirectory() + "/" + split[1];
                }
            }
            // DownloadsProvider
            else if (isDownloadsDocument(uri)) {

                final String id = DocumentsContract.getDocumentId(uri);
                final Uri contentUri = ContentUris.withAppendedId(
                        Uri.parse("content://downloads/public_downloads"), Long.valueOf(id));

                return getDataColumn(context, contentUri, null, null);
            }

            // MediaProvider
            else if (isMediaDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];

                Uri contentUri = null;
                if ("image".equals(type)) {
                    contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                } else if ("video".equals(type)) {
                    contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
                } else if ("audio".equals(type)) {
                    contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
                }

                final String selection = "_id=?";
                final String[] selectionArgs = new String[]{split[1]};

                return getDataColumn(context, contentUri, selection, selectionArgs);
            }
        }
        // MediaStore (and general)
        else if ("content".equalsIgnoreCase(uri.getScheme())) {
            return getDataColumn(context, uri, null, null);
        }
        // File
        else if ("file".equalsIgnoreCase(uri.getScheme())) {
            return uri.getPath();
        }
        return null;
    }


    public String getDataColumn(Context context, Uri uri, String selection, String[] selectionArgs) {
        Cursor cursor = null;
        final String column = "_data";
        final String[] projection = {column};

        try {
            cursor = context.getContentResolver().query(uri, projection, selection, selectionArgs,
                    null);
            if (cursor != null && cursor.moveToFirst()) {
                final int column_index = cursor.getColumnIndexOrThrow(column);
                return cursor.getString(column_index);
            }
        } finally {
            if (cursor != null)
                cursor.close();
        }
        return null;
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is ExternalStorageProvider.
     */
    public boolean isExternalStorageDocument(Uri uri) {
        return "com.android.externalstorage.documents".equals(uri.getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is DownloadsProvider.
     */
    public boolean isDownloadsDocument(Uri uri) {
        return "com.android.providers.downloads.documents".equals(uri.getAuthority());
    }

    /**
     * @param uri The Uri to check.
     * @return Whether the Uri authority is MediaProvider.
     */
    public boolean isMediaDocument(Uri uri) {
        return "com.android.providers.media.documents".equals(uri.getAuthority());
    }



    //获取文件的真实路径
    public String getRealPathFromURI(Uri contentUri) {
        String res = null;
        String[] proj = { MediaStore.Images.Media.DATA };
        Cursor cursor = getContentResolver().query(contentUri, proj, null, null, null);
        if(cursor.moveToFirst()){;
            int column_index = cursor.getColumnIndexOrThrow(MediaStore.Images.Media.DATA);
            res = cursor.getString(column_index);
        }
        cursor.close();
        return res;
    }


    /**
     * 如果前面两种都获取不到文件
     * 则使用此种方法拼接路径
     * 此方法在Andorid7.0系统中可用
     */
    private void splicingPath(Uri uri){
        Log.i("hxl", "获取文件的路径filePath========="+filePath);
        if(filePath.endsWith(".xls")){
            Log.i("hxl", "===调用拼接路径方法===");
            String string =uri.toString();
            String a[]=new String[2];
            //判断文件是否在sd卡中
            if (string.indexOf(String.valueOf(Environment.getExternalStorageDirectory()))!=-1){
                //对Uri进行切割
                a = string.split(String.valueOf(Environment.getExternalStorageDirectory()));
                //获取到file
                file = new File(Environment.getExternalStorageDirectory(),a[1]);
            }else if(string.indexOf(String.valueOf(Environment.getDataDirectory()))!=-1) { //判断文件是否在手机内存中
                //对Uri进行切割
                a = string.split(String.valueOf(Environment.getDataDirectory()));
                //获取到file
                file = new File(Environment.getDataDirectory(), a[1]);
            }
//            fileName = filePath.substring(filePath.lastIndexOf("/") + 1);
//            Log.i("hxl", "file========="+file);
        }else{
            Toast.makeText(MainActivity.this, "您选中的文件不是xls格式文档", Toast.LENGTH_LONG).show();
        }
    }





}
