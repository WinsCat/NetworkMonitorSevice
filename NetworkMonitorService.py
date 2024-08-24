import socket
import subprocess
import time

import servicemanager
import win32event
import win32service
import win32serviceutil


class NetworkMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NetworkMonitorService"
    _svc_display_name_ = "Network Monitor Service"
    _svc_description_ = "Monitors network and changes gateway based on ping results."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

        # Configuration
        self.target = "www.baidu.com"
        self.default_gateway = "192.168.110.1"
        self.backup_gateway = "192.168.110.3"
        self.interval = 20  # seconds
        self.adapter_name = self.get_adapter_name()

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
        print(self.adapter_name)
        while self.running:
            if self.ping_target(self.target):
                # self.set_gateway(self.default_gateway)
                print('主网络正常')
            else:
                print('启用备用网络')
                self.set_gateway(self.backup_gateway)
            time.sleep(self.interval)

    def ping_target(self, target):
        try:
            socket.setdefaulttimeout(2)
            host = socket.gethostbyname(target)
            print(host)
            s = socket.create_connection((host, 80), 2)
            s.close()
            return True
        except Exception as e:
            print(e)
            return False

    def set_gateway(self, gateway):
        if self.is_dhcp_enabled(self.adapter_name):
            self.release_renew_dhcp(self.adapter_name)
        else:
            # subprocess.call(f'netsh interface ip set address source="static" name="{self.adapter_name}" gateway={gateway}', shell=True)
            print('no DHCP')

    def get_adapter_name(self):
        # 自动检测活动的网络适配器名称
        result = subprocess.run('netsh interface show interface', shell=True, capture_output=True, text=True)
        for line in result.stdout.splitlines():
            if "已连接" in line:
                return line.split()[-1]  # 返回连接状态的适配器名称
        return None

    def is_dhcp_enabled(self, adapter_name):
        result = subprocess.run(f'netsh interface ip show config name="{adapter_name}"', shell=True,
                                capture_output=True, text=True)
        return "DHCP 已启用:                          是" in result.stdout

    def release_renew_dhcp(self, adapter_name):
        # 释放并更新DHCP
        subprocess.call(f'ipconfig /release "{adapter_name}"', shell=True)
        time.sleep(5)  # 等待一段时间以确保释放
        subprocess.call(f'ipconfig /renew "{adapter_name}"', shell=True)
        time.sleep(5)  # 等待一段时间以确保续订
        servicemanager.LogInfoMsg(f"DHCP settings refreshed for adapter {adapter_name}")


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(NetworkMonitorService)
