import os
import time
import win32serviceutil
import win32service
import win32event
import servicemanager
from datetime import datetime
import win32com.client

OUTPUT_FILE = r"C:\Users\Avik Jain\Documents\amazeme\ms_service\scheduled_tasks_log.txt"

class ScheduledTasksService(win32serviceutil.ServiceFramework):
    _svc_name_ = "ScheduledTasksService"
    _svc_display_name_ = "Scheduled Tasks Monitor Service"
    _svc_description_ = "Logs Windows scheduled tasks summary and history."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ""))

        while self.running:
            try:
                self.write_task_summary()
            except Exception as e:
                with open(OUTPUT_FILE, "a") as f:
                    f.write(f"Error: {e}\n")
            time.sleep(60)  # update every minute

    def write_task_summary(self):
        scheduler = win32com.client.Dispatch('Schedule.Service')
        scheduler.Connect()
        rootFolder = scheduler.GetFolder('\\')
        tasks = rootFolder.GetTasks(0)

        lines = []
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines.append(f"Task Summary as of {now}")
        lines.append("="*50)

        for task in tasks:
            name = task.Name
            path = task.Path
            next_run = task.NextRunTime if task.NextRunTime else "N/A"

            # Task triggers (schedule)
            triggers = []
            for trig in task.Definition.Triggers:
                triggers.append(str(trig.StartBoundary))

            # Last run status
            last_run_time = task.LastRunTime if task.LastRunTime else "N/A"
            last_result = task.LastTaskResult if task.LastTaskResult else "N/A"

            lines.append(f"Task: {name}")
            lines.append(f"  Path: {path}")
            lines.append(f"  Next Run: {next_run}")
            lines.append(f"  Last Run: {last_run_time}")
            lines.append(f"  Last Result: {last_result}")
            lines.append(f"  Schedule: {', '.join(triggers) if triggers else 'N/A'}")
            lines.append("-"*50)

        with open(OUTPUT_FILE, "w") as f:
            f.write("\n".join(lines))


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(ScheduledTasksService)
