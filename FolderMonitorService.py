import time
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32serviceutil
import win32service
import win32event
import servicemanager

FOLDER_TO_WATCH = r"C:\Users\Avik Jain\Documents\amazeme\ms_service\monitor"
OUTPUT_FILE = r"C:\Users\Avik Jain\Documents\amazeme\ms_service\folder_log.txt"

class FolderHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()

    def on_any_event(self, event):
        # Update the text file whenever a file is added/removed
        try:
            files = os.listdir(FOLDER_TO_WATCH)
            with open(OUTPUT_FILE, "w") as f:
                f.write("\n".join(files))
        except Exception as e:
            with open(OUTPUT_FILE, "a") as f:
                f.write(f"Error: {e}\n")

class FolderMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "FolderMonitorService"
    _svc_display_name_ = "Folder Monitor Service"
    _svc_description_ = "Tracks files in a folder and logs changes."

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

        event_handler = FolderHandler()
        observer = Observer()
        observer.schedule(event_handler, FOLDER_TO_WATCH, recursive=False)
        observer.start()

        # Initial write of existing files
        event_handler.on_any_event(None)

        while self.running:
            time.sleep(1)

        observer.stop()
        observer.join()


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(FolderMonitorService)
