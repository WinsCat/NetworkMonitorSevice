import os
import sys
import time
import socket
import subprocess
import win32serviceutil
import win32service
import win32event
import servicemanager

class NetworkMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NetworkMonitorService"
    _svc_display_name_ = "Network Monitor Service"
    _svc_description_ = "Monitors network and changes gateway based on ping results."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

        # Configuration
        self.target = "8.8.8.8"
        self.default_gateway = "192.168.1.999"
        self.backup_gateway = "10.211.55.999"
        self.interval = 60  # seconds

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )
        self.main()

    def main(self):
        while self.running:
            if self.ping_target(self.target):
                self.set_gateway(self.default_gateway)
            else:
                self.set_gateway(self.backup_gateway)
            time.sleep(self.interval)

    def ping_target(self, target):
        try:
            socket.setdefaulttimeout(2)
            host = socket.gethostbyname(target)
            s = socket.create_connection((host, 80), 2)
            s.close()
            return True
        except Exception as e:
            return False

    def set_gateway(self, gateway):
        subprocess.call(f'netsh interface ip set address name="Ethernet" gateway={gateway}', shell=True)

if __name__ == '__main__':
    import servicemanager
    if len(sys.argv) == 1:
        try:
            evtsrc_dll = os.path.abspath(servicemanager.__file__)
            # 如果修改过名字，名字要统一
            servicemanager.PrepareToHostSingle(NetworkMonitorService)
            # 如果修改过名字，名字要统一
            servicemanager.Initialize('NetworkMonitorService', evtsrc_dll)
            servicemanager.StartServiceCtrlDispatcher()
        except win32service.error as details:
            import winerror
            if details == winerror.ERROR_FAILED_SERVICE_CONTROLLER_CONNECT:
                win32serviceutil.usage()
    else:
        # 如果修改过名字，名字要统一
        win32serviceutil.HandleCommandLine(NetworkMonitorService)
