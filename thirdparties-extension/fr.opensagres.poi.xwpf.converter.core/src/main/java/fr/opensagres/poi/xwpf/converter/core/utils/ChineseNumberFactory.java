package fr.opensagres.poi.xwpf.converter.core.utils;


public class ChineseNumberFactory {

    private static final char[] CN_SIMPLIFIED = {'〇', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'};
    private static final char[] CN_SIMPLIFIED_SERIES = {'十', '百', '千', '万'};

    // chineseCountingThousand
    private static final char[] CN_THOUSAND = {'零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', };
    private static final char[] CN_THOUSAND_SERIES = {'拾', '佰', '仟', '萬'};

    // ideographTraditional
    private static final String[] TRADITIONAL = {"甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"};
    // ideographZodiac
    private static final String[] ZODIAC = {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"};

    /**
     * 中文简体转换
     * @return
     */
    public static String getChineseLegalSimplified(int number)
    {
        if (number <= 0 || number > 99999)
        {
            return String.valueOf(CN_SIMPLIFIED[0]);
        }
        if (number <= 9)
        {
            return String.valueOf(CN_SIMPLIFIED[number]);
        }
        StringBuilder sb = new StringBuilder();
        String numStr = String.valueOf(number);
        int len = numStr.length();
        boolean isAddZero = false;
        for (int i = 0; i < len; i++)
        {
            int t = numStr.charAt(i) - 0x30;
            if (t > 0)
            {
                sb.append(CN_SIMPLIFIED[t]);
                if (len - i - 2 >= 0)
                {
                    sb.append(CN_SIMPLIFIED_SERIES[len - i - 2]);
                }
                isAddZero = true;
            }
            else if (isAddZero && i != len - 1)
            {
                sb.append(CN_SIMPLIFIED[0]);
                isAddZero = false;
            }
        }
        // delete last "零"
        if (sb.charAt(sb.length() - 1) == CN_SIMPLIFIED[0])
        {
            sb.deleteCharAt(sb.length() - 1);
        }

        return sb.toString();
    }

    /**
     * 中文繁体转换
     * @param number
     * @return
     */
    public static String getChineseCountingThousand(int number)
    {
        if (number <= 0 || number > 99999)
        {
            return String.valueOf(CN_THOUSAND[0]);
        }
        if (number <= 9)
        {
            return String.valueOf(CN_THOUSAND[number]);
        }
        StringBuilder sb = new StringBuilder();
        String numStr = String.valueOf(number);
        int len = numStr.length();
        boolean isAddZero = false;
        for (int i = 0; i < len; i++)
        {
            int t = numStr.charAt(i) - 0x30;
            if (t > 0)
            {
                sb.append(CN_THOUSAND[t]);
                if (len - i - 2 >= 0)
                {
                    sb.append(CN_THOUSAND_SERIES[len - i - 2]);
                }
                isAddZero = true;
            }
            else if (isAddZero && i != len - 1)
            {
                sb.append(CN_THOUSAND[0]);
                isAddZero = false;
            }
        }
        // delete last "〇"
        if (sb.charAt(sb.length() - 1) == CN_THOUSAND[0])
        {
            sb.deleteCharAt(sb.length() - 1);
        }
        // delete first "一"
        if (number > 10 && number < 20)
        {
            if (sb.charAt(0) == CN_THOUSAND[1])
            {
                sb.deleteCharAt(0);
            }
        }

        return sb.toString();
    }

    public static String getIdeographTraditional(int num){
        return num <= 10 ? TRADITIONAL[num - 1] : String.valueOf(num);
    }

    public static String getIdeographZodiac(int num){
        return num <= 12 ? ZODIAC[num - 1] : String.valueOf(num);
    }

    public static String getCardinalText(int num)
    {
        String numberStr = String.valueOf(num);
        String lStr = numberStr;// 没有小数点的情况
        String lStrRev = reverseString(lStr);// 对左边的字串取反字串
        String[] a = new String[5];// 定义5个字串变量用来存放解析出的三位一组的字串
        switch (lStrRev.length() % 3)
        {
            case 1:
                lStrRev = lStrRev + "00";
                break;
            case 2:
                lStrRev = lStrRev + "0";
                break;
            default:
                break;
        }
        String StrInt = "";
        for (int i = 0; i <= lStrRev.length() / 3 - 1; i++)// 计算有多少个三位
        {
            a[i] = reverseString(lStrRev.substring(3 * i, 3 * i + 3));// 截取第1个三位
            if (!a[i].equals("000"))// 用来避免这种情况“1000000=ONE MILLION THOUSAND ONLY”
            {
                if (i != 0)
                {
                    StrInt = w3(a[i]) + " " + dw(String.valueOf(i)) + " "
                            + StrInt;// 用来加上“THOUSAND
                    // OR
                    // MILLION
                    // OR
                    // BILLION”
                }
                else
                {
                    StrInt = w3(a[i]);// 防止i=0时“lm=w3(a(i))+" "+dw(i)+" "+lm”多加两个尾空格
                }
            }
            else
            {
                StrInt = w3(a[i]) + StrInt;
            }
        }
        return toUpperCaseFirstOne(StrInt);
    }

    // 将字符串反置
    private static String reverseString(String str)
    {
        int lenInt = str.length();
        String[] z = new String[str.length()];
        for (int i = 0; i < lenInt; i++)
        {
            z[i] = str.substring(i, i + 1);
        }
        str = "";
        for (int i = lenInt - 1; i >= 0; i--)
        {
            str = str + z[i];
        }
        return str;
    }

    private static String zr4(String y)
    {
        String[] z = new String[10];
        z[0] = "";
        z[1] = "one";
        z[2] = "two";
        z[3] = "three";
        z[4] = "four";
        z[5] = "five";
        z[6] = "six";
        z[7] = "seven";
        z[8] = "eight";
        z[9] = "nine";
        return z[Integer.parseInt(y.substring(0, 1))];
    }

    private static String zr3(String y)
    {
        String[] z = new String[10];
        z[0] = "";
        z[1] = "one";
        z[2] = "two";
        z[3] = "three";
        z[4] = "four";
        z[5] = "five";
        z[6] = "six";
        z[7] = "seven";
        z[8] = "eight";
        z[9] = "nine";
        return z[Integer.parseInt(y.substring(2, 3))];
    }

    private static String zr2(String y)
    {
        String[] z = new String[20];
        z[10] = "ten";
        z[11] = "eleven";
        z[12] = "twelve";
        z[13] = "thirteen";
        z[14] = "fourteen";
        z[15] = "fifteen";
        z[16] = "sixteen";
        z[17] = "seventeen";
        z[18] = "eighteen";
        z[19] = "nineteen";
        return z[Integer.parseInt(y.substring(1, 3))];
    }

    private static String zr1(String y)
    {
        String[] z = new String[10];
        z[1] = "ten";
        z[2] = "twenty";
        z[3] = "thirty";
        z[4] = "forty";
        z[5] = "fifty";
        z[6] = "sixty";
        z[7] = "seventy";
        z[8] = "eighty";
        z[9] = "ninety";
        return z[Integer.parseInt(y.substring(1, 2))];
    }

    private static String dw(String y)
    {
        String[] z = new String[5];
        z[0] = "";
        z[1] = "thousand";
        z[2] = "million";
        z[3] = "billion";
        return z[Integer.parseInt(y)];
    }

    // 用来制作2位数字转英文
    private static String w2(String y)
    {
        String tempstr;
        if (y.substring(1, 2).equals("0"))// 判断是否小于十
        {
            tempstr = zr3(y);
        }
        else if (y.substring(1, 2).equals("1"))// 判断是否在十到二十之间
        {
            tempstr = zr2(y);
        }
        else
        {
            if (y.substring(2, 3).equals("0"))// 判断是否为大于二十小于一百的能被十整除的数（为了去掉尾空格）
            {
                tempstr = zr1(y);
            }
            else
            {
                tempstr = zr1(y) + "-" + zr3(y);
            }
        }
        return tempstr;
    }

    private static String w3(String y)
    {
        String tempstr;
        if (y.substring(0, 1).equals("0"))// 判断是否小于一百
        {
            tempstr = w2(y);
        }
        else
        {
            if (y.substring(1, 3).equals("00"))// 判断是否能被一百整除
            {
                tempstr = zr4(y) + " " + "hundred";
            }
            else
            {
                tempstr = zr4(y) + " " + "hundred" + " " + w2(y);// 不能整除的要后面加“AND”
            }
        }
        return tempstr;
    }

    public static String toUpperCaseFirstOne(String s)
    {
        if (s.equals(""))
        {
            return String.valueOf(0);
        }
        else
        {
            return (new StringBuilder())
                    .append(Character.toUpperCase(s.charAt(0)))
                    .append(s.substring(1)).toString();
        }
    }

}
