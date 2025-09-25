import win32serviceutil
import win32service
import win32event
import servicemanager
import time
from datetime import datetime
import os

class TimeUpdateService(win32serviceutil.ServiceFramework):
    _svc_name_ = "TimeUpdateService"   # Service name
    _svc_display_name_ = "Time Update Service"  # How it shows in Services
    _svc_description_ = "Updates current time every second into a text file."

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

        self.main()

    def main(self):
        filepath = os.path.join(os.path.dirname(__file__), "time_log.txt")

        while self.running:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                with open(filepath, "w") as f:
                    f.write(now)
            except Exception as e:
                pass  # ignore errors for simplicity
            time.sleep(1)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(TimeUpdateService)
