using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace MailForALM
{

	/// <summary>
	/// Класс взят отсюда: http://www.codeproject.com/Articles/11902/Convert-HTML-to-Plain-Text
	/// </summary>
	public static class TextConverter
	{
		public static string ConvertHtmlToText(string source)
		{
			string AmpStr;
			MatchCollection AmpCodes;

			// Add all possible special character into a dictionary
			Dictionary<string, int> SpecChars = new Dictionary<string, int>();
			SpecChars.Add("&Aacute;", 193); SpecChars.Add("&aacute;", 225); SpecChars.Add("&Acirc;", 194);
			SpecChars.Add("&acirc;", 226); SpecChars.Add("&acute;", 180); SpecChars.Add("&AElig;", 198);
			SpecChars.Add("&aelig;", 230); SpecChars.Add("&Agrave;", 192); SpecChars.Add("&agrave;", 224);
			SpecChars.Add("&alefsym;", 8501); SpecChars.Add("&Alpha;", 913); SpecChars.Add("&alpha;", 945);
			SpecChars.Add("&amp;", 38); SpecChars.Add("&and;", 8743); SpecChars.Add("&ang;", 8736);
			SpecChars.Add("&Aring;", 197); SpecChars.Add("&aring;", 229); SpecChars.Add("&asymp;", 8776);
			SpecChars.Add("&Atilde;", 195); SpecChars.Add("&atilde;", 227); SpecChars.Add("&Auml;", 196);
			SpecChars.Add("&auml;", 228); SpecChars.Add("&bdquo;", 8222); SpecChars.Add("&Beta;", 914);
			SpecChars.Add("&beta;", 946); SpecChars.Add("&brvbar;", 166); SpecChars.Add("&bull;", 8226);
			SpecChars.Add("&cap;", 8745); SpecChars.Add("&Ccedil;", 199); SpecChars.Add("&ccedil;", 231);
			SpecChars.Add("&cedil;", 184); SpecChars.Add("&cent;", 162); SpecChars.Add("&Chi;", 935);
			SpecChars.Add("&chi;", 967); SpecChars.Add("&circ;", 710); SpecChars.Add("&clubs;", 9827);
			SpecChars.Add("&cong;", 8773); SpecChars.Add("&copy;", 169); SpecChars.Add("&crarr;", 8629);
			SpecChars.Add("&cup;", 8746); SpecChars.Add("&curren;", 164); SpecChars.Add("&dagger;", 8224);
			SpecChars.Add("&Dagger;", 8225); SpecChars.Add("&darr;", 8595); SpecChars.Add("&dArr;", 8659);
			SpecChars.Add("&deg;", 176); SpecChars.Add("&Delta;", 916); SpecChars.Add("&delta;", 948);
			SpecChars.Add("&diams;", 9830); SpecChars.Add("&divide;", 247); SpecChars.Add("&Eacute;", 201);
			SpecChars.Add("&eacute;", 233); SpecChars.Add("&Ecirc;", 202); SpecChars.Add("&ecirc;", 234);
			SpecChars.Add("&Egrave;", 200); SpecChars.Add("&egrave;", 232); SpecChars.Add("&emdash;", 8212);
			SpecChars.Add("&empty;", 8709); SpecChars.Add("&emsp;", 8195); SpecChars.Add("&endash;", 8211);
			SpecChars.Add("&ensp;", 8194); SpecChars.Add("&Epsilon;", 917); SpecChars.Add("&epsilon;", 949);
			SpecChars.Add("&equiv;", 8801); SpecChars.Add("&Eta;", 919); SpecChars.Add("&eta;", 951);
			SpecChars.Add("&ETH;", 208); SpecChars.Add("&eth;", 240); SpecChars.Add("&Euml;", 203);
			SpecChars.Add("&euml;", 235); SpecChars.Add("&euro;", 8364); SpecChars.Add("&exist;", 8707);
			SpecChars.Add("&fnof;", 402); SpecChars.Add("&forall;", 8704); SpecChars.Add("&frac12;", 189);
			SpecChars.Add("&frac14;", 188); SpecChars.Add("&frac34;", 190); SpecChars.Add("&frasl;", 8260);
			SpecChars.Add("&Gamma;", 915); SpecChars.Add("&gamma;", 947); SpecChars.Add("&ge;", 8805);
			SpecChars.Add("&gt;", 62); SpecChars.Add("&harr;", 8596); SpecChars.Add("&hArr;", 8660);
			SpecChars.Add("&hearts;", 9829); SpecChars.Add("&hellip;", 8230); SpecChars.Add("&Iacute;", 205);
			SpecChars.Add("&iacute;", 237); SpecChars.Add("&Icirc;", 206); SpecChars.Add("&icirc;", 238);
			SpecChars.Add("&iexcl;", 161); SpecChars.Add("&Igrave;", 204); SpecChars.Add("&igrave;", 236);
			SpecChars.Add("&image;", 8465); SpecChars.Add("&infin;", 8734); SpecChars.Add("&int;", 8747);
			SpecChars.Add("&Iota;", 921); SpecChars.Add("&iota;", 953); SpecChars.Add("&iquest;", 191);
			SpecChars.Add("&isin;", 8712); SpecChars.Add("&Iuml;", 207); SpecChars.Add("&iuml;", 239);
			SpecChars.Add("&Kappa;", 922); SpecChars.Add("&kappa;", 954); SpecChars.Add("&Lambda;", 923);
			SpecChars.Add("&lambda;", 955); SpecChars.Add("&lang;", 9001); SpecChars.Add("&laquo;", 171);
			SpecChars.Add("&larr;", 8592); SpecChars.Add("&lArr;", 8656); SpecChars.Add("&lceil;", 8968);
			SpecChars.Add("&ldquo;", 8220); SpecChars.Add("&le;", 8804); SpecChars.Add("&lfloor;", 8970);
			SpecChars.Add("&lowast;", 8727); SpecChars.Add("&loz;", 9674); SpecChars.Add("&lrm;", 8206);
			SpecChars.Add("&lsaquo;", 8249); SpecChars.Add("&lsquo;", 8216); SpecChars.Add("&lt;", 60);
			SpecChars.Add("&macr;", 175); SpecChars.Add("&mdash;", 8212); SpecChars.Add("&micro;", 181);
			SpecChars.Add("&middot;", 183); SpecChars.Add("&minus;", 8722); SpecChars.Add("&Mu;", 924);
			SpecChars.Add("&mu;", 956); SpecChars.Add("&nabla;", 8711); SpecChars.Add("&nbsp;", 32); // Using space instead of Ascii 160
			SpecChars.Add("&ndash;", 8211); SpecChars.Add("&ne;", 8800); SpecChars.Add("&ni;", 8715);
			SpecChars.Add("&not;", 172); SpecChars.Add("&notin;", 8713); SpecChars.Add("&nsub;", 8836);
			SpecChars.Add("&Ntilde;", 209); SpecChars.Add("&ntilde;", 241); SpecChars.Add("&Nu;", 925);
			SpecChars.Add("&nu;", 957); SpecChars.Add("&Oacute;", 211); SpecChars.Add("&oacute;", 243);
			SpecChars.Add("&Ocirc;", 212); SpecChars.Add("&ocirc;", 244); SpecChars.Add("&OElig;", 338);
			SpecChars.Add("&oelig;", 339); SpecChars.Add("&Ograve;", 210); SpecChars.Add("&ograve;", 242);
			SpecChars.Add("&oline;", 8254); SpecChars.Add("&Omega;", 937); SpecChars.Add("&omega;", 969);
			SpecChars.Add("&Omicron;", 927); SpecChars.Add("&omicron;", 959); SpecChars.Add("&oplus;", 8853);
			SpecChars.Add("&or;", 8744); SpecChars.Add("&ordf;", 170); SpecChars.Add("&ordm;", 186);
			SpecChars.Add("&Oslash;", 216); SpecChars.Add("&oslash;", 248); SpecChars.Add("&Otilde;", 213);
			SpecChars.Add("&otilde;", 245); SpecChars.Add("&otimes;", 8855); SpecChars.Add("&Ouml;", 214);
			SpecChars.Add("&ouml;", 246); SpecChars.Add("&para;", 182); SpecChars.Add("&part;", 8706);
			SpecChars.Add("&permil;", 8240); SpecChars.Add("&perp;", 8869); SpecChars.Add("&Phi;", 934);
			SpecChars.Add("&phi;", 966); SpecChars.Add("&Pi;", 928); SpecChars.Add("&pi;", 960);
			SpecChars.Add("&piv;", 982); SpecChars.Add("&plusmn;", 177); SpecChars.Add("&pound;", 163);
			SpecChars.Add("&prime;", 8242); SpecChars.Add("&Prime;", 8243); SpecChars.Add("&prod;", 8719);
			SpecChars.Add("&prop;", 8733); SpecChars.Add("&Psi;", 936); SpecChars.Add("&psi;", 968);
			SpecChars.Add("&quot;", 34); SpecChars.Add("&radic;", 8730); SpecChars.Add("&rang;", 9002);
			SpecChars.Add("&raquo;", 187); SpecChars.Add("&rarr;", 8594); SpecChars.Add("&rArr;", 8658);
			SpecChars.Add("&rceil;", 8969); SpecChars.Add("&rdquo;", 8221); SpecChars.Add("&real;", 8476);
			SpecChars.Add("&reg;", 174); SpecChars.Add("&rfloor;", 8971); SpecChars.Add("&Rho;", 929);
			SpecChars.Add("&rho;", 961); SpecChars.Add("&rlm;", 8207); SpecChars.Add("&rsaquo;", 8250);
			SpecChars.Add("&rsquo;", 8217); SpecChars.Add("&sbquo;", 8218); SpecChars.Add("&Scaron;", 352);
			SpecChars.Add("&scaron;", 353); SpecChars.Add("&sdot;", 8901); SpecChars.Add("&sect;", 167);
			SpecChars.Add("&shy;", 173); SpecChars.Add("&Sigma;", 931); SpecChars.Add("&sigma;", 963);
			SpecChars.Add("&sigmaf;", 962); SpecChars.Add("&sim;", 8764); SpecChars.Add("&spades;", 9824);
			SpecChars.Add("&sub;", 8834); SpecChars.Add("&sube;", 8838); SpecChars.Add("&sum;", 8721);
			SpecChars.Add("&sup;", 8835); SpecChars.Add("&sup1;", 185); SpecChars.Add("&sup2;", 178);
			SpecChars.Add("&sup3;", 179); SpecChars.Add("&supe;", 8839); SpecChars.Add("&szlig;", 223);
			SpecChars.Add("&Tau;", 932); SpecChars.Add("&tau;", 964); SpecChars.Add("&there4;", 8756);
			SpecChars.Add("&Theta;", 920); SpecChars.Add("&theta;", 952); SpecChars.Add("&thetasym;", 977);
			SpecChars.Add("&thinsp;", 8201); SpecChars.Add("&THORN;", 222); SpecChars.Add("&thorn;", 254);
			SpecChars.Add("&tilde;", 732); SpecChars.Add("&times;", 215); SpecChars.Add("&trade;", 8482);
			SpecChars.Add("&Uacute;", 218); SpecChars.Add("&uacute;", 250); SpecChars.Add("&uarr;", 8593);
			SpecChars.Add("&uArr;", 8657); SpecChars.Add("&Ucirc;", 219); SpecChars.Add("&ucirc;", 251);
			SpecChars.Add("&Ugrave;", 217); SpecChars.Add("&ugrave;", 249); SpecChars.Add("&uml;", 168);
			SpecChars.Add("&upsih;", 978); SpecChars.Add("&Upsilon;", 933); SpecChars.Add("&upsilon;", 965);
			SpecChars.Add("&Uuml;", 220); SpecChars.Add("&uuml;", 252); SpecChars.Add("&weierp;", 8472);
			SpecChars.Add("&Xi;", 926); SpecChars.Add("&xi;", 958); SpecChars.Add("&Yacute;", 221);
			SpecChars.Add("&yacute;", 253); SpecChars.Add("&yen;", 165); SpecChars.Add("&yuml;", 255);
			SpecChars.Add("&Yuml;", 376); SpecChars.Add("&Zeta;", 918); SpecChars.Add("&zeta;", 950);
			SpecChars.Add("&zwj;", 8205); SpecChars.Add("&zwnj;", 8204);

			HtmlToTextReplaceElement[] ReplaceStrings = new HtmlToTextReplaceElement[]
            {
                // Remove HTML Development formatting
                //Replace any white space characters (line breaks, tabs, spaces) with space because browsers inserts space 
                new HtmlToTextReplaceElement(@"\s", " ", HtmlToTextReplaceType.RegEx),
                // Remove repeating speces becuase browsers ignore them 
                new HtmlToTextReplaceElement(@" {2,}", " ", HtmlToTextReplaceType.RegEx),
                /*
                  I'm using .* in my regex from here to match "all" characters. It works here ONLY because I've removed
                  all linebreaks beforehand. If you doesn't done that you MUST replace .* with [\s\đ]* to match all characters
                  in multiple lines
                */
                // Remove HTML comment
                new HtmlToTextReplaceElement(@"<! *--.*?-- *>", " ", HtmlToTextReplaceType.RegEx),
                // Remove the header
                new HtmlToTextReplaceElement(@"< *head( *>| [^>]*>).*< */ *head *>", string.Empty, HtmlToTextReplaceType.RegEx),
                // remove all scripts
                new HtmlToTextReplaceElement(@"< *script( *>| [^>]*>).*?< */ *script *>", string.Empty, HtmlToTextReplaceType.RegEx),
                // remove all styles (prepare first by clearing attributes)
                new HtmlToTextReplaceElement(@"< *style( *>| [^>]*>).*?< */ *style *>", string.Empty, HtmlToTextReplaceType.RegEx),
                // insert tabs in spaces of <td> tags
                new HtmlToTextReplaceElement(@"< *td[^>]*>","\t", HtmlToTextReplaceType.RegEx),
                // insert line breaks in places of <BR> and <LI> tags 
                new HtmlToTextReplaceElement(@"< *(br|li) */{0,1} *>", "\r", HtmlToTextReplaceType.RegEx),
                new HtmlToTextReplaceElement(@"< *(div|tr|p)( *>| [^>]*>)", "\r\r", HtmlToTextReplaceType.RegEx),
                // Remove remaining tags like <a>, links, images, etc - anything thats enclosed inside < > 
                new HtmlToTextReplaceElement(@"<[^>]*>", string.Empty, HtmlToTextReplaceType.RegEx),
                // Replace &nbsp; with whitespace. It is done here because the generated space will be used in
                // whitespace optimizations
                new HtmlToTextReplaceElement(@"&nbsp;", " ", HtmlToTextReplaceType.String),
                // Remove extra line breaks and tabs:
                // Romove any whitespace and tab at and of any line
                new HtmlToTextReplaceElement(@"[ \t]+\r", "\r", HtmlToTextReplaceType.RegEx),
                // Remove whitespace beetween tabs
                new HtmlToTextReplaceElement(@"\t +\t", "\t\t", HtmlToTextReplaceType.RegEx),
                // Remove whitespace begining of a line if followed by a tab
                new HtmlToTextReplaceElement(@"\r +\t", "\r\t", HtmlToTextReplaceType.RegEx),
                // Remove multible tabs following a linebreak with just one tab 
                new HtmlToTextReplaceElement(@"\r\t{2,}", "\r\t", HtmlToTextReplaceType.RegEx),
                // replace over 2 breaks with 2 and over 4 tabs with 4.  
                new HtmlToTextReplaceElement(@"\r{3,}", "\r\r", HtmlToTextReplaceType.RegEx),
                new HtmlToTextReplaceElement(@"\t{4,}", "\t\t\t\t", HtmlToTextReplaceType.RegEx)
            };

			string result = source ?? "";
			// run pattern matching
			for (int i = 0; i < ReplaceStrings.Length; i++)
			{
				switch (ReplaceStrings[i].Type)
				{
					case HtmlToTextReplaceType.String:
						if (result != null) 
							result = result.Replace(ReplaceStrings[i].Pattern, ReplaceStrings[i].Substitute);
						break;
					case HtmlToTextReplaceType.RegEx:
						if (result != null) 
							result = Regex.Replace(result, ReplaceStrings[i].Pattern, ReplaceStrings[i].Substitute, RegexOptions.IgnoreCase);
						break;
				}
			}
			// Replace decimal character codes
			AmpCodes = Regex.Matches(result, @"&#\d{1,5};", RegexOptions.IgnoreCase);
			for (int i = AmpCodes.Count - 1; i >= 0; i--)
			{
				AmpStr = AmpCodes[i].Value;
				result = result.Substring(0, AmpCodes[i].Index) +
					Convert.ToChar(Int32.Parse(AmpStr.Substring(2, AmpStr.Length - 3))) +
					result.Substring(AmpCodes[i].Index + AmpCodes[i].Length);
			}
			// Replace hexadecimal character codes
			AmpCodes = Regex.Matches(result, @"&#x[0-9a-f]{1,4};", RegexOptions.IgnoreCase);
			for (int i = AmpCodes.Count - 1; i >= 0; i--)
			{
				AmpStr = AmpCodes[i].Value;
				result = result.Substring(0, AmpCodes[i].Index) +
					Convert.ToChar(Int32.Parse(AmpStr.Substring(3, AmpStr.Length - 4), NumberStyles.AllowHexSpecifier)) +
					result.Substring(AmpCodes[i].Index + AmpCodes[i].Length);
			}
			// Replace named character codes
			AmpCodes = Regex.Matches(result, @"&\w+;", RegexOptions.IgnoreCase);
			for (int i = AmpCodes.Count - 1; i >= 0; i--)
			{
				if (SpecChars.ContainsKey(AmpCodes[i].Value))
				{
					result = result.Substring(0, AmpCodes[i].Index) +
						Convert.ToChar(SpecChars[AmpCodes[i].Value]) +
						result.Substring(AmpCodes[i].Index + AmpCodes[i].Length);
				}
			}
			// Remove all others
			result = Regex.Replace(result, @"&[^;]*;", string.Empty, RegexOptions.IgnoreCase);
			return result;
		}
	}
	public enum HtmlToTextReplaceType { String, RegEx }
	public class HtmlToTextReplaceElement
	{
		public HtmlToTextReplaceElement() { }
		public HtmlToTextReplaceElement(string Pattern, string Substitute, HtmlToTextReplaceType Type)
		{
			this.Pattern = Pattern;
			this.Substitute = Substitute;
			this.Type = Type;
		}
		public string Pattern;
		public string Substitute;
		public HtmlToTextReplaceType Type;
	}
}
