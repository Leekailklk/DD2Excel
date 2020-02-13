package com.example.administrator.dd2excel;

import android.content.Context;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ArrayAdapter;
import android.widget.CheckBox;
import android.widget.TextView;
import com.example.administrator.dd2excel.bean.CheckinRecord;
import com.example.administrator.dd2excel.util.Util;

import java.util.List;

public class CheckinAdapter extends ArrayAdapter<CheckinRecord> {
    private int resourceId;
    public CheckinAdapter(Context context, int textViewResourceId,
                          List<CheckinRecord> objects) {
        super(context, textViewResourceId, objects);
        resourceId = textViewResourceId;
    }
    @Override
    public View getView(int position, View convertView, ViewGroup parent) {
        CheckinRecord record = getItem(position); // 获取当前项的实例
        View view;
        ViewHolder viewHolder;

        if (convertView == null) {
            view = LayoutInflater.from(getContext()).inflate(resourceId, null);
            viewHolder = new ViewHolder();
            viewHolder.useIt=(CheckBox) view.findViewById(R.id.checkin_use);
            viewHolder.name= (TextView) view.findViewById(R.id.checkin_name);
            viewHolder.time=(TextView) view.findViewById(R.id.checkin_time);
            viewHolder.place=(TextView) view.findViewById(R.id.checkin_place);
            viewHolder.detail=(TextView) view.findViewById(R.id.checkin_detail);
            view.setTag(viewHolder);// 将ViewHolder存储在View中
        } else {
            view = convertView;
            viewHolder = (ViewHolder) view.getTag(); // 重新获取ViewHolder
        }

        viewHolder.useIt.setChecked(record.isUse_it());
        viewHolder.name.setText(record.getName());
        //viewHolder.time.setText(Util.stampToDate(record.getTimestamp()));
        viewHolder.time.setText(record.getTimestamp());
        viewHolder.place.setText(Util.getCity(record.getDetailPlace()));
        //viewHolder.detail.setText(record.getDetailPlace());
        return view;
    }
    class ViewHolder {
        CheckBox useIt;
        TextView name;
        TextView time;
        TextView place;
        TextView detail;
    }

}
