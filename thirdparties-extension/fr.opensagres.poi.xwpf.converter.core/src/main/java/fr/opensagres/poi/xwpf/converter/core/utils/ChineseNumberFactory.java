package fr.opensagres.poi.xwpf.converter.core.utils;


public class ChineseNumberFactory {

    private static final char[] CN_SIMPLIFIED = {'零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', };
    private static final char[] CN_SIMPLIFIED_SERIES = {'拾', '佰', '仟', '萬'};

    // chineseCountingThousand
    private static final char[] CN_THOUSAND = {'〇', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'};
    private static final char[] CN_THOUSAND_SERIES = {'十', '百', '千', '万'};


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


}
