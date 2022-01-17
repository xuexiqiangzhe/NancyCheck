package org.imc.tools;

import com.sun.istack.internal.NotNull;

import java.security.InvalidParameterException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author dengchengchao
 * @Time 2018/5/22
 * @Description 处理数字的工具类
 */
public class DigitUtils {
    /**
     * 中文數字转阿拉伯数组、是数字则直接返回数字
     * @author 雪见烟寒
     * @param chineseNumber
     * @return
     */
    @SuppressWarnings("unused")
    public static int chineseNumber2Int(String chineseNumber) throws Exception {
        if(NumberTool.isInteger(chineseNumber)){
            return Integer.parseInt(chineseNumber);
        }
        int result = 0;
        int temp = 1;//存放一个单位的数字如：十万
        int count = 0;//判断是否该把上一个单位先加入result
        char[] cnArr = new char[]{'一','二','三','四','五','六','七','八','九','两'};
        char[] chArr = new char[]{'十','百','千','万','亿'};


        for (int i = 0; i < chineseNumber.length(); i++) {
            char c = chineseNumber.charAt(i);
            boolean flag = false;
            for (int j = 0; j < cnArr.length; j++) {
                if (c == cnArr[j]) {
                    flag = true;
                }
            }
            for (int j = 0; j < chArr.length; j++) {
                if (c == chArr[j]) {
                    flag = true;
                }
            }
            if(c=='零'){
                flag = true;
            }
            if(!flag){
                System.out.println("不支持该中文章节号转英文:"+chineseNumber);
                throw new Exception();
            }
        }


        for (int i = 0; i < chineseNumber.length(); i++) {
            boolean b = true;//判断是否是chArr
            char c = chineseNumber.charAt(i);

            // 数字{'一','二','三','四','五','六','七','八','九','两'};
            for (int j = 0; j < cnArr.length; j++) {
                if (c == cnArr[j]) {
                    if(0 != count){//添加下一个单位之前，先把上一个单位值添加到结果中
                        result += temp;
                        temp = 1;
                        count = 0;
                    }
                    // 下标+1，就是对应的值
                    if(j<9){
                        temp = j + 1;
                    }else if(j==9){
                        temp = 2;
                    }
                    b = false;
                    break;
                }
            }
            if(b){//单位{'十','百','千','万','亿'}
                for (int j = 0; j < chArr.length; j++) {
                    if (c == chArr[j]) {
                        switch (j) {
                            case 0:
                                temp *= 10;
                                break;
                            case 1:
                                temp *= 100;
                                break;
                            case 2:
                                temp *= 1000;
                                break;
                            case 3:
                                temp *= 10000;
                                break;
                            case 4:
                                temp *= 100000000;
                                break;
                            default:
                                break;
                        }
                        count++;
                    }
                }
            }
            if (i == chineseNumber.length() - 1) {//遍历到最后一个字符
                result += temp;
            }
        }
        return result;
    }

    public static void main(String[] args) throws Exception {
        String[] inputs = {"两千万","两亿两万"};
        for(String input:inputs){
            System.out.println(chineseNumber2Int(input));
        }
    }
}

