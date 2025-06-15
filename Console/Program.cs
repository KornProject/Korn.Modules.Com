using Korn.Com.Wmi;

var processWatcher = new ProcessWatcher();

processWatcher.ProcessStarted += process => Console.WriteLine($"started: {process.Name}");
processWatcher.ProcessStopped += process => Console.WriteLine($"stopped: {process.Name}");

Thread.Sleep(5000);
processWatcher.Dispose();
Thread.Sleep(-1);