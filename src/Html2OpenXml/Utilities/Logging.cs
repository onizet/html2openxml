/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Diagnostics;
using System.Globalization;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Logging class to trace debugging information during the conversion process.
    /// </summary>
    static class Logging
	{
		private const string TraceSourceName = "html2openxml";
		private static TraceSource traceSource;
		private static bool enabled;


		static Logging()
		{
			Initialize();
		}

		#region PrintError

		public static void PrintError(string method, String message)
		{
			if (!ValidateSettings(TraceEventType.Error))
				return;

			PrintLine(TraceEventType.Error, 0, "Exception in the " + method + " - " + message);
		}

		public static void PrintError(string method, Exception exception)
		{
			if (!ValidateSettings(TraceEventType.Error))
				return;

			PrintLine(TraceEventType.Error, 0, "Exception in the " + method + " - " + exception.Message);
			if (!String.IsNullOrEmpty(exception.StackTrace))
				PrintLine(TraceEventType.Error, 0, exception.StackTrace);
		}

		#endregion

		#region PrintVerbose

		public static void PrintVerbose(string msg)
		{
			if (!ValidateSettings(TraceEventType.Verbose))
				return;

			PrintLine(TraceEventType.Verbose, 0, msg);
		}

		#endregion

		// Private Implementation

		#region Initialize

		/// <summary>
		/// Initialize the logger from the app.config.
		/// </summary>
		private static void Initialize()
		{
#if NETSTANDARD1_3 || NETSTANDARD2_0
            traceSource = new TraceSource(TraceSourceName);
			enabled = traceSource.Switch.Level != SourceLevels.Off;
#else
            try
			{
				traceSource = new TraceSource(TraceSourceName);
				enabled = traceSource.Switch.Level != SourceLevels.Off;
			}
            catch (System.Configuration.ConfigurationException)
            {
                // app.config has an error
                enabled = false;
			}

            if (enabled)
			{
				AppDomain appDomain = AppDomain.CurrentDomain;
				appDomain.DomainUnload += OnDomainUnload;
				appDomain.ProcessExit += OnDomainUnload;
			}
#endif
        }

        #endregion

        #region PrintLine

		/// <summary>
		/// Core method to write in the log.
		/// </summary>
		private static void PrintLine(TraceEventType eventType, int id, string msg)
		{
			if (!ValidateSettings(eventType)) return;
			traceSource.TraceEvent(eventType, id, msg);
		}

        #endregion

        #region OnDomainUnload

		/// <summary>
		/// Event handler to close properly the trace source when the program is shut down.
		/// </summary>
		private static void OnDomainUnload(object sender, EventArgs e)
		{
			traceSource.Close();
			enabled = false;
		}

        #endregion

        #region ValidateSettings

		/// <summary>
		/// Ensure the type of event should be traced, regarding the configuration.
		/// </summary>
		private static bool ValidateSettings(TraceEventType traceLevel)
		{
			if (!enabled) return false;

			if (traceSource == null || !traceSource.Switch.ShouldTrace(traceLevel))
				return false;

			return true;
		}

        #endregion

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets whether the tracing is enabled or not.
		/// </summary>
		public static bool On
		{
			get { return enabled; }
		}
	}
}