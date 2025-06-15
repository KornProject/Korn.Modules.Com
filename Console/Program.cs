using Korn.Com;

var processWatcher = new WMIProcessWatcher();
processWatcher.SetProcessCreatedHandler(process => Console.WriteLine(process.Name));

Thread.Sleep(5000);
processWatcher.Dispose();
Thread.Sleep(-1);