using DEA;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteLog
{
    internal class WriteLogClass
    {
        public static void WriteToLog(int Level, string LogEntry)
        {
            var LogFileName = "DEA_Logfile_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
            var LogFile = Path.Combine(GraphHelper.CheckFolders("Log"), LogFileName);

            var LogControlSwitch = new LoggingLevelSwitch();

            LogEventLevel LogLevel;
            
            switch (Level)
            {
                case 1 :
                    LogLevel = LogEventLevel.Error;                    
                    break;

                case 2 :
                    LogLevel = LogEventLevel.Warning;
                    break;

                case 3 :
                    LogLevel = LogEventLevel.Information;
                    break;

                case 4 :
                     LogLevel = LogEventLevel.Debug;
                    break;

                case 5 :
                     LogLevel = LogEventLevel.Verbose;
                    break;

                default :
                    LogLevel = LogEventLevel.Fatal;
                    break;
            }

            LogControlSwitch.MinimumLevel = LogLevel;

            Log.Logger = new LoggerConfiguration()
                        .MinimumLevel.ControlledBy(LogControlSwitch)
                        .WriteTo.File(LogFile)
                        .WriteTo.Console()                        
                        .CreateLogger();

            WriteLog(Level, LogEntry);
            
            Log.CloseAndFlush();
        }

        private static void WriteLog(int LogLevel, string LogEntryString)
        {
            if (LogLevel == 1)
            {
                Log.Error(LogEntryString);
            }
            else if (LogLevel == 2)
            {
                Log.Warning(LogEntryString);
            }
            else if (LogLevel == 3)
            {
                Log.Information(LogEntryString);
            }
            else if (LogLevel == 4)
            {
                Log.Debug(LogEntryString);
            }
            else if (LogLevel == 5)
            {
                Log.Verbose(LogEntryString);
            }
            else
            {
                Log.Error(LogEntryString);
            }
        }
    }
}
