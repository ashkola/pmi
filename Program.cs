using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PMI
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			var t = new Main();
			var str = "<html>  <body>  <ol style=\"margin-top: 0mm; margin-bottom: 0mm; \">  <li style=\"font-family: Arial;  color: #010101;  \"><font face=\"Arial\"><span style=\"font-size:8pt; font-family:Arial; color:#010101; font-weight:Normal; font-style:Normal; font-decoration:Normal\">&nbsp;Complete   the First Name and Last Name fields, as follows:  </span></font></li>  </ol>  <ul start=\"2\" style=\"margin-top: 0mm; margin-bottom: 0mm; list-style-type: disc; \">  <li style=\"list-style-type: none;\">  <ul style=\"margin-top: 0mm; margin-bottom: 0mm; list-style-type: circle; \">  <li style=\"font-family: Symbol;  color: #010101;  \"><font face=\"Arial\"><span style=\"font-size:8pt; font-family:Arial; color:#010101; font-weight:Normal; font-style:Normal; font-decoration:Normal\">&nbsp;&nbsp;   First and Last Name can not be empty</span></font></li>  <li style=\"font-family: Symbol;  color: #010101;  \"><font face=\"Arial\"><span style=\"font-size:8pt; font-family:Arial; color:#010101; font-weight:Normal; font-style:Normal; font-decoration:Normal\">&nbsp;   The First and Last Name fields may contain letters, hyphen &quot;-&quot;,&nbsp; and space.</span></font></li>  </ul>  </li>  </ul>  <ol start=\"2\" style=\"margin-top: 0mm; margin-bottom: 0mm; \">  <li style=\"font-family: Arial;  color: #010101;  \"><font face=\"Arial\"><span style=\"font-size:8pt; font-family:Arial; color:#010101; font-weight:Normal; font-style:Normal; font-decoration:Normal\">   Click the Secure Purchase button. 3. Check the results on the next page.</span></font></li>  </ol>  </body>  </html>";
			MessageBox.Show(t.htmlToText(str) + "\n\n" + t.StripHTML(str));

//			Application.EnableVisualStyles();
//			Application.SetCompatibleTextRenderingDefault(false);
//			Application.Run(new Form1());
		}
	}
}
