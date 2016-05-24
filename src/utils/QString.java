package utils;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class QString {

	/**
	 * @category 验证是不是Email
	 */
	public static boolean isEmail(String email) {
		String str = "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$";
		return regexCheckValidity(email, str);
	}

	/**
	 * @category 根据正则检查数据有效性
	 */
	public static boolean regexCheckValidity(String data, String regex) {
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(data);
		return m.matches();
	}

	/**
	 * 获取字符串的长度 一个汉字算两个字符
	 *
	 * @param s
	 * @return
	 */
	public static int getWordCount(String s) {

		s = s.replaceAll("[^\\x00-\\xff]", "**");
		return s.length();
	}

	/**
	 * @param type
	 *            正则模式 0中文，1英文，2数字
	 * @return String
	 * @category 过滤字符串符合正则模式的部分
	 */
	public static String regexCheck(String data, int type) {
		switch (type) {
		case 0:
			// 这里手动处理一下正则不能过滤出来的字符
			// 中文后加x会过滤到的中文有x，不知道是什么原因
			String[] sps = { ",", "，", ".", "。", "“", "”", "\"", "-", "+", "{", "}", "『", "』", "?", "？", "×", "*", "x",
					"y", "!", "@", "#", "$", "^", "z", "v", "w", "~", "%", "&", "(", ")", "|", "'", "/", "～", "！", "￥",
					"…", "（", "）", "—", "：", ".", "》", "《", "；", "‘", "【", "】", "、" };
			for (String s : sps) {
				data = data.replace(s, "");
			}
			return regexCheck(data, "[u4e00-u9fa5]");
		case 1:
			return regexCheck(data, "[^A-Za-z]");
		case 2:
			return regexCheck(data, "[^0-9]");
		}
		return data;
	}

	/**
	 * @param regex
	 *            正则表达式
	 * @return String
	 * @category 过滤字符串符合正则部分的字符
	 * @author jinghq
	 */
	public static String regexCheck(String data, String regex) {
		try {
			Pattern p = Pattern.compile(regex);
			Matcher m = p.matcher(data);
			return m.replaceAll("").trim();
		} catch (Exception e) {
			return "";
		}
	}

	/**
	 * 判断是否是手机号码
	 */
	public static boolean isMoblieNumber_old(String number) {
		if (number == null || number.length() == 0) {
			return false;
		}
		String reg = "^(\\+{0,1}|0{0,2})(86){0,1}0{0,1}(13[0-9]|15[0-9]|18[0-9]|14[0-9])[0-9]{8}$";
		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(number);
		return m == null ? false : m.find();
	}

	/**
	 * 判断是否是手机号码,只要是1开头的就可以
	 *
	 * @param number
	 * @return
	 * @since 2014-11-19
	 */
	public static boolean isMoblieNumber(String number) {
		if (number == null || number.length() == 0) {
			return false;
		}

		String reg = "^(\\+{0,1}|0{0,2})(86){0,1}0{0,1}(1[0-9][0-9])[0-9]{8}$";
		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(number);
		return m == null ? false : m.find();
	}

	/**
	 * 判断是否是网址
	 */
	public static String getUrl(String address) {
		if (address == null || address.length() == 0) {
			return "";
		}
		String reg = "(?i)\\b((?:[a-z][\\w-]+:(?:/{1,3}|[a-z0-9%])|www\\d{0,3}[.]|[a-z0-9.\\-]+[.][a-z]{2,4}/)(?:[^\\s()<>]+|\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\))+(?:\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\)|[^\\s`!()\\[\\]{};:\'\".,<>?«»“”‘’]))";
		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(address);

		boolean isFind = m == null ? false : m.find();
		if (isFind) {
			return m.group();
		}
		return "";
	}

	public static boolean isUrl(String url) {
		if (url == null || url.length() == 0) {
			return false;
		}
		String reg = "(?i)\\b((?:[a-z][\\w-]+:(?:/{1,3}|[a-z0-9%])|www\\d{0,3}[.]|[a-z0-9.\\-]+[.][a-z]{2,4}/)(?:[^\\s()<>]+|\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\))+(?:\\(([^\\s()<>]+|(\\([^\\s()<>]+\\)))*\\)|[^\\s`!()\\[\\]{};:\'\".,<>?«»“”‘’]))";
		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(url);

		return m == null ? false : m.find();
	}

	/**
	 * 判断是否只含有数字和字母
	 */
	public static boolean isOnlyNumAndLetter(String str) {
		if ((str == null) || (str.length() == 0)) {
			return false;
		}

		String reg = "^[0-9A-Za-z]+$";

		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(str);

		return m == null ? false : m.find();
	}

	/**
	 * 是否只含有数字
	 **/
	public static boolean isOnlyNum(String str) {
		if ((str == null) || (str.length() == 0)) {
			return false;
		}

		String reg = "^[0-9]*$";

		Pattern p = Pattern.compile(reg);
		Matcher m = p.matcher(str);

		return m == null ? false : m.find();
	}

	/**
	 * 是否包含中文的判断
	 *
	 * @param str
	 * @return
	 */
	public static boolean isContainsChinese(String str) {
		String reg = "[\u4e00-\u9fa5]";
		Pattern p = Pattern.compile(reg);
		Matcher matcher = p.matcher(str);
		return matcher.find() ? true : false;
	}

	/**
	 * @param text
	 *            要截取的字符串
	 * @return String
	 * @author jinghq
	 * @category 截取最后一个字
	 */
	public static String getLastWord(String text) {
		if (text == null) {
			return "";
		}
		text = getNameLastWord(text);
		int len = text.length();
		if (len > 1) {
			return text.substring(len - 2, len);
		} else if (len > 0) {
			return text.substring(len - 1, len);
		} else {
			return "";
		}
	}

	/**
	 * @category 按照汉字取最后一个字, 没有汉字的返回“”
	 */
	public static String getNameLastWord(String text) {
		String c = regexCheck(text, 0);
		return c.length() == 0 ? "" : c;
	}

	private QString() {
		throw new AssertionError();
	}

	/**
	 * is null or its length is 0 or it is made by space
	 * <p/>
	 * 
	 * <pre>
	 * isBlank(null) = true;
	 * isBlank(&quot;&quot;) = true;
	 * isBlank(&quot;  &quot;) = true;
	 * isBlank(&quot;a&quot;) = false;
	 * isBlank(&quot;a &quot;) = false;
	 * isBlank(&quot; a&quot;) = false;
	 * isBlank(&quot;a b&quot;) = false;
	 * </pre>
	 *
	 * @param str
	 * @return if string is null or its size is 0 or it is made by space, return
	 *         true, else return false.
	 */
	public static boolean isBlank(String str) {
		return (str == null || str.trim().length() == 0);
	}

	/**
	 * is null or its length is 0
	 * <p/>
	 * 
	 * <pre>
	 * isEmpty(null) = true;
	 * isEmpty(&quot;&quot;) = true;
	 * isEmpty(&quot;  &quot;) = false;
	 * </pre>
	 *
	 * @param str
	 * @return if string is null or its size is 0, return true, else return
	 *         false.
	 */
	public static boolean isEmpty(CharSequence str) {
		return (str == null || str.length() == 0);
	}

	/**
	 * compare two string
	 *
	 * @param actual
	 * @param expected
	 * @return
	 * @see QObject#isEquals(Object, Object)
	 */
	public static boolean isEquals(String actual, String expected) {
		return QObject.isEquals(actual, expected);
	}

	/**
	 * get length of CharSequence
	 * <p/>
	 * 
	 * <pre>
	 * length(null) = 0;
	 * length(\"\") = 0;
	 * length(\"abc\") = 3;
	 * </pre>
	 *
	 * @param str
	 * @return if str is null or empty, return 0, else return
	 *         {@link CharSequence#length()}.
	 */
	public static int length(CharSequence str) {
		return str == null ? 0 : str.length();
	}

	/**
	 * null Object to empty string
	 * <p/>
	 * 
	 * <pre>
	 * nullStrToEmpty(null) = &quot;&quot;;
	 * nullStrToEmpty(&quot;&quot;) = &quot;&quot;;
	 * nullStrToEmpty(&quot;aa&quot;) = &quot;aa&quot;;
	 * </pre>
	 *
	 * @param str
	 * @return
	 */
	public static String nullStrToEmpty(Object str) {
		return (str == null ? "" : (str instanceof String ? (String) str : str.toString()));
	}

	/**
	 * capitalize first letter
	 * <p/>
	 * 
	 * <pre>
	 * capitalizeFirstLetter(null)     =   null;
	 * capitalizeFirstLetter("")       =   "";
	 * capitalizeFirstLetter("2ab")    =   "2ab"
	 * capitalizeFirstLetter("a")      =   "A"
	 * capitalizeFirstLetter("ab")     =   "Ab"
	 * capitalizeFirstLetter("Abc")    =   "Abc"
	 * </pre>
	 *
	 * @param str
	 * @return
	 */
	public static String capitalizeFirstLetter(String str) {
		if (isEmpty(str)) {
			return str;
		}

		char c = str.charAt(0);
		return (!Character.isLetter(c) || Character.isUpperCase(c)) ? str
				: new StringBuilder(str.length()).append(Character.toUpperCase(c)).append(str.substring(1)).toString();
	}

	/**
	 * encoded in utf-8
	 * <p/>
	 * 
	 * <pre>
	 * utf8Encode(null)        =   null
	 * utf8Encode("")          =   "";
	 * utf8Encode("aa")        =   "aa";
	 * utf8Encode("啊啊啊啊")   = "%E5%95%8A%E5%95%8A%E5%95%8A%E5%95%8A";
	 * </pre>
	 *
	 * @param str
	 * @return
	 * @throws UnsupportedEncodingException
	 *             if an error occurs
	 */
	public static String utf8Encode(String str) {
		if (!isEmpty(str) && str.getBytes().length != str.length()) {
			try {
				return URLEncoder.encode(str, "UTF-8");
			} catch (UnsupportedEncodingException e) {
				throw new RuntimeException("UnsupportedEncodingException occurred. ", e);
			}
		}
		return str;
	}

	/**
	 * encoded in utf-8, if exception, return defultReturn
	 *
	 * @param str
	 * @param defultReturn
	 * @return
	 */
	public static String utf8Encode(String str, String defultReturn) {
		if (!isEmpty(str) && str.getBytes().length != str.length()) {
			try {
				return URLEncoder.encode(str, "UTF-8");
			} catch (UnsupportedEncodingException e) {
				return defultReturn;
			}
		}
		return str;
	}

	/**
	 * get innerHtml from href
	 * <p/>
	 * 
	 * <pre>
	 * getHrefInnerHtml(null)                                  = ""
	 * getHrefInnerHtml("")                                    = ""
	 * getHrefInnerHtml("mp3")                                 = "mp3";
	 * getHrefInnerHtml("&lt;a innerHtml&lt;/a&gt;")                    = "&lt;a innerHtml&lt;/a&gt;";
	 * getHrefInnerHtml("&lt;a&gt;innerHtml&lt;/a&gt;")                    = "innerHtml";
	 * getHrefInnerHtml("&lt;a&lt;a&gt;innerHtml&lt;/a&gt;")                    = "innerHtml";
	 * getHrefInnerHtml("&lt;a href="baidu.com"&gt;innerHtml&lt;/a&gt;")               = "innerHtml";
	 * getHrefInnerHtml("&lt;a href="baidu.com" title="baidu"&gt;innerHtml&lt;/a&gt;") = "innerHtml";
	 * getHrefInnerHtml("   &lt;a&gt;innerHtml&lt;/a&gt;  ")                           = "innerHtml";
	 * getHrefInnerHtml("&lt;a&gt;innerHtml&lt;/a&gt;&lt;/a&gt;")                      = "innerHtml";
	 * getHrefInnerHtml("jack&lt;a&gt;innerHtml&lt;/a&gt;&lt;/a&gt;")                  = "innerHtml";
	 * getHrefInnerHtml("&lt;a&gt;innerHtml1&lt;/a&gt;&lt;a&gt;innerHtml2&lt;/a&gt;")        = "innerHtml2";
	 * </pre>
	 *
	 * @param href
	 * @return
	 * 		<ul>
	 *         <li>if href is null, return ""</li>
	 *         <li>if not match regx, return source</li>
	 *         <li>return the last string that match regx</li>
	 *         </ul>
	 */
	public static String getHrefInnerHtml(String href) {
		if (isEmpty(href)) {
			return "";
		}

		String hrefReg = ".*<[\\s]*a[\\s]*.*>(.+?)<[\\s]*/a[\\s]*>.*";
		Pattern hrefPattern = Pattern.compile(hrefReg, Pattern.CASE_INSENSITIVE);
		Matcher hrefMatcher = hrefPattern.matcher(href);
		if (hrefMatcher.matches()) {
			return hrefMatcher.group(1);
		}
		return href;
	}

	/**
	 * process special char in html
	 * <p/>
	 * 
	 * <pre>
	 * htmlEscapeCharsToString(null) = null;
	 * htmlEscapeCharsToString("") = "";
	 * htmlEscapeCharsToString("mp3") = "mp3";
	 * htmlEscapeCharsToString("mp3&lt;") = "mp3<";
	 * htmlEscapeCharsToString("mp3&gt;") = "mp3\>";
	 * htmlEscapeCharsToString("mp3&amp;mp4") = "mp3&mp4";
	 * htmlEscapeCharsToString("mp3&quot;mp4") = "mp3\"mp4";
	 * htmlEscapeCharsToString("mp3&lt;&gt;&amp;&quot;mp4") = "mp3\<\>&\"mp4";
	 * </pre>
	 *
	 * @param source
	 * @return
	 */
	public static String htmlEscapeCharsToString(String source) {
		return QString.isEmpty(source) ? source
				: source.replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&amp;", "&").replaceAll("&quot;",
						"\"");
	}

	/**
	 * transform half width char to full width char
	 * <p/>
	 * 
	 * <pre>
	 * fullWidthToHalfWidth(null) = null;
	 * fullWidthToHalfWidth("") = "";
	 * fullWidthToHalfWidth(new String(new char[] {12288})) = " ";
	 * fullWidthToHalfWidth("！＂＃＄％＆) = "!\"#$%&";
	 * </pre>
	 *
	 * @param s
	 * @return
	 */
	public static String fullWidthToHalfWidth(String s) {
		if (isEmpty(s)) {
			return s;
		}

		char[] source = s.toCharArray();
		for (int i = 0; i < source.length; i++) {
			if (source[i] == 12288) {
				source[i] = ' ';
				// } else if (source[i] == 12290) {
				// source[i] = '.';
			} else if (source[i] >= 65281 && source[i] <= 65374) {
				source[i] = (char) (source[i] - 65248);
			} else {
				source[i] = source[i];
			}
		}
		return new String(source);
	}

	/**
	 * transform full width char to half width char
	 * <p/>
	 * 
	 * <pre>
	 * halfWidthToFullWidth(null) = null;
	 * halfWidthToFullWidth("") = "";
	 * halfWidthToFullWidth(" ") = new String(new char[] {12288});
	 * halfWidthToFullWidth("!\"#$%&) = "！＂＃＄％＆";
	 * </pre>
	 *
	 * @param s
	 * @return
	 */
	public static String halfWidthToFullWidth(String s) {
		if (isEmpty(s)) {
			return s;
		}

		char[] source = s.toCharArray();
		for (int i = 0; i < source.length; i++) {
			if (source[i] == ' ') {
				source[i] = (char) 12288;
				// } else if (source[i] == '.') {
				// source[i] = (char)12290;
			} else if (source[i] >= 33 && source[i] <= 126) {
				source[i] = (char) (source[i] + 65248);
			} else {
				source[i] = source[i];
			}
		}
		return new String(source);
	}
}
