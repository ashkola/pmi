using System;
using System.IO;

namespace PMI
{
	class Config
	{
		public static string _server;
		public static string _login;
		public static string _password;
		public static string _domain;
		public static string _project;
		public static bool _attachment;
		public static string _root;

		private const string FILENAME = "Settings.dat";
		public static void Save()
		{
			using (BinaryWriter writer = new BinaryWriter(File.Open(FILENAME,FileMode.Create)))
			{
				writer.Write(_server);
				writer.Write(_login);
				writer.Write("");
//				writer.Write(_password);
				writer.Write(_domain);
				writer.Write(_project);
				writer.Write(_root);
				writer.Write(_attachment);
			}
		}
		public static void Get()
		{
			if(File.Exists(FILENAME))
			{
				using (BinaryReader reader = new BinaryReader(File.Open(FILENAME,FileMode.Open)))
				{
					_server = reader.ReadString();
					_login = reader.ReadString();
					_password = reader.ReadString();
					_domain = reader.ReadString();
					_project = reader.ReadString();
					_root = reader.ReadString();
					_attachment = reader.ReadBoolean();
				}				
			}
		}
	}
}
