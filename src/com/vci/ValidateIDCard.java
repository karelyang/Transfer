package com.vci;

import java.io.IOException;
import java.util.Hashtable;
import java.util.regex.Pattern;

public class ValidateIDCard {
 
    /**
     * 身份证号验证(18位新身份证号)
     * @param validateStr 身份证号码
     * @throws IOException 
     */
    public static boolean isValidIDCardNum(String name, String validateStr) throws IOException{
        if(validateStr.length() != 18){
        	System.out.println(name + " 身份证号码长度有错误，期望是18，实际是" + validateStr.length());
            return false;
        }
        String regex = "[0-9]{17}[0-9|xX]$";
        Pattern pattern = Pattern.compile(regex, Pattern.CASE_INSENSITIVE);
        if(!pattern.matcher(validateStr).matches()){//前17位数字，第18位数字或字母X
            return false;
        }
        Hashtable<String, String> areacodeHashtable = new Hashtable<String, String>();
        areacodeHashtable.put("11", "北京");     
        areacodeHashtable.put("12", "天津");     
        areacodeHashtable.put("13", "河北");     
        areacodeHashtable.put("14", "山西");     
        areacodeHashtable.put("15", "内蒙古");   
        areacodeHashtable.put("21", "辽宁");     
        areacodeHashtable.put("22", "吉林");     
        areacodeHashtable.put("23", "黑龙江");   
        areacodeHashtable.put("31", "上海");     
        areacodeHashtable.put("32", "江苏");     
        areacodeHashtable.put("33", "浙江");     
        areacodeHashtable.put("34", "安徽");     
        areacodeHashtable.put("35", "福建");     
        areacodeHashtable.put("36", "江西");     
        areacodeHashtable.put("37", "山东");     
        areacodeHashtable.put("41", "河南");     
        areacodeHashtable.put("42", "湖北");     
        areacodeHashtable.put("43", "湖南");     
        areacodeHashtable.put("44", "广东");     
        areacodeHashtable.put("45", "广西");     
        areacodeHashtable.put("46", "海南");     
        areacodeHashtable.put("50", "重庆");     
        areacodeHashtable.put("51", "四川");     
        areacodeHashtable.put("52", "贵州");     
        areacodeHashtable.put("53", "云南");     
        areacodeHashtable.put("54", "西藏");     
        areacodeHashtable.put("61", "陕西");     
        areacodeHashtable.put("62", "甘肃");     
        areacodeHashtable.put("63", "青海");     
        areacodeHashtable.put("64", "宁夏");     
        areacodeHashtable.put("65", "新疆");     
        areacodeHashtable.put("71", "台湾");     
        areacodeHashtable.put("81", "香港");     
        areacodeHashtable.put("82", "澳门");     
        areacodeHashtable.put("91", "国外");
        if(areacodeHashtable.get(validateStr.substring(0,2)) == null){
        	System.out.println(name + " 身份证所在位置不正确");
            return false;
        }
        String birthdate = validateStr.substring(6, 14);//获取生日
        int birthyear = Integer.parseInt(birthdate.substring(0, 4));//出生年份
        int curyear = java.util.Calendar.getInstance().get(java.util.Calendar.YEAR);//当前时间年份
        if(birthyear>=curyear || birthyear <1900){//不在1900年与当前时间之间
        	System.out.println(name + " 身份证的出生年份不正确");
            return false;
        }
        String dateRegex = "^((\\d{2}(([02468][048])|([13579][26]))[\\/\\/\\s]?((((0 ?[13578])|(1[02]))[\\/\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\/\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\/\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|([13579][01345789]))[\\/\\/\\s]?((((0?[13578])|(1[02]))[\\/\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\/\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\/\\/\\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))";
        if(!Pattern.matches(dateRegex, birthdate)){//出生年月日不合法
        	System.out.println(name + " 身份证的出生日期不正确");
            return false;
        }
        /**
         * 计算校验码（第十八位数）：
         * 十七位数字本体码加权求和公式 S = Sum(Ai * Wi), i = 0...16 ，先对前17位数字的权求和；
         * Ai:表示第i位置上的身份证号码数字值 Wi:表示第i位置上的加权因子 Wi: 7 9 10 5 8 4 2 1 6 3 7 9 10 5 8 4 2；
         * 计算模 Y = mod(S, 11)
         * 通过模Y得到对应的校验码: 0 1 2 3 4 5 6 7 8 9 10 校验码: 1 0 X 9 8 7 6 5 4 3 2
         */
/*        final String[] LASTCODE = { "1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2" };// 18位身份证中最后一位校验码
        final int[] WEIGHT = { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };// 18位身份证中，前17位数字各个数字的生成校验码时的权值
        String tempLastCode = "";//临时记录身份证号码最后一位
        int sum = 0;//前17位号码与对应权重值相乘总和
        for(int i=0; i<17; i++){
            sum += ((int)(validateStr.charAt(i)-'0')) * WEIGHT[i];
        }
        tempLastCode = LASTCODE[sum%11];//实际最后一位号码
        if(validateStr.substring(17).equals(tempLastCode)){//最后一位符合
            return true;
        }else{
        	System.out.println(name + "的身份证的最后一位不对，应该是" + tempLastCode);
            return false;
        }
*/
        return true;
    }
}