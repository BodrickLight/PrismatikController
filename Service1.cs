using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PrismatikController
{
	public partial class PrismatikController : ServiceBase
	{
		public PrismatikController ()
		{
			InitializeComponent ();
		}

		protected override void OnStart (string[] args)
		{
			Run ();
		}

		protected override void OnStop ()
		{
			_cts.Cancel ();
			while (_cts2.IsCancellationRequested)
			{
				Thread.Sleep (100);
				Task.Yield ();
			}

			_cts.Dispose ();
			_cts2.Dispose ();
		}

		private async void Run ()
		{
			var currentOpen = IsCameraOpen ();
			while (_cts.IsCancellationRequested == false)
			{
				var open = IsCameraOpen ();
				if (open != currentOpen)
				{
					var profile = open ? "Light ring" : "Ambilight";
					//$"Changing profile to {profile}.".Dump ();
					await SetAmbilightProfile (profile);
				}
				currentOpen = open;
				await Task.Delay (1000);
			}

			_cts2.Cancel ();
		}

		private IEnumerable<Times> GetTimes (string rootKey, RegistryKey hive)
		{
			using (var keys = hive.OpenSubKey (rootKey))
			{
				foreach (var key in keys.GetSubKeyNames ())
				{
					using (var subkey = keys.OpenSubKey (key))
					{
						var startTime = DateTime.FromFileTime (((long?) subkey.GetValue ("LastUsedTimeStart")) ?? 0);
						var endTime = DateTime.FromFileTime (((long?) subkey.GetValue ("LastUsedTimeStop")) ?? 0);
						yield return new Times (subkey.Name.Split ('\\').Last (), startTime, endTime);
					}
				}
			}
		}

		private bool IsCameraOpen ()
		{
			try
			{
				var array = new List<Times> ();
				array.AddRange (GetTimes (@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam", Registry.LocalMachine));
				array.AddRange (GetTimes (@"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged", Registry.LocalMachine));
				array.AddRange (GetTimes (@"S-1-5-21-196879825-2503152832-2194935188-1001\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam", Registry.Users));
				array.AddRange (GetTimes (@"S-1-5-21-196879825-2503152832-2194935188-1001\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged", Registry.Users));
				//array.Select(x => $"{x.Process}: {x.StartTime} - {x.EndTime}").Dump ();
				var open = array.FirstOrDefault (x => x.EndTime < x.StartTime);
				//open?.Process.Dump ();
				return open != null;
			}
			catch (Exception e)
			{
				EventLog.WriteEntry (e.Message);
			}
			return false;
		}

		private async Task SetAmbilightProfile (string profile)
		{
			try
			{
				var key = "{b5fc6a64-0050-480d-8c85-75031aec782b}";
				using (var client = new TcpClient ())
				{
					await client.ConnectAsync ("127.0.0.1", 3636);
					using (var stream = client.GetStream ())
					{
						await WriteAsync ($"apikey:{key}", stream);
						await WriteAsync ($"lock", stream);
						await WriteAsync ($"setprofile:{profile}", stream);
						await WriteAsync ($"unlock", stream);
						await WriteAsync ($"exit", stream);
					}
				}
			}
			catch (Exception e)
			{
				EventLog.WriteEntry (e.Message);
			}
		}

		async Task WriteAsync (string content, Stream stream)
		{
			var data = Encoding.ASCII.GetBytes (content + "\n");
			//("-> " + content).Dump ();
			await stream.WriteAsync (data, 0, data.Length);
			await Task.Delay (100);
			var buf = new byte[256];
			await stream.ReadAsync (buf, 0, 256);
			//("<- " + Encoding.ASCII.GetString (buf.TakeWhile (x => x != '\0').ToArray ())).Dump ();
		}

		private CancellationTokenSource _cts = new CancellationTokenSource ();
		private CancellationTokenSource _cts2 = new CancellationTokenSource ();

		private class Times
		{
			public Times (string process, DateTime startTime, DateTime endTime)
			{
				Process = process;
				StartTime = startTime;
				EndTime = endTime;
			}
			public string Process { get; }
			public DateTime StartTime { get; }
			public DateTime EndTime { get; }
		}
	}
}