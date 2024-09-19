import os
import socket
import subprocess
import sys
import time

import servicemanager
import win32event
import win32service
import win32serviceutil


class NetworkMonitorService(win32serviceutil.ServiceFramework):
    _svc_name_ = "NetworkMonitorService"
    _svc_display_name_ = "Network Monitor Service"
    _svc_description_ = "融汇Turing小组出品，用于实现网关热备功能。"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

        # Configuration
        self.target = "www.baidu.com"
        self.default_gateway = "192.168.110.1"
        self.backup_gateway = "192.168.110.11"
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
        # win32evtlogutil.AddSourceToRegistry("RonghuiTuring", "NetworkMonitorService")
        while self.running:
            if self.ping_target(self.default_gateway):
                # 判断当前是否使用备用网络，如果是切换为主网络
                if self.get_backup_gateway_ip(self.adapter_name):
                    print('当前为：备用网关状态')
                    self.set_gateway(self.default_gateway)

                else:
                    self.ping_target(self.target)
                    # self.set_gateway(self.default_gateway)
                    # print(self.get_current_ip(self.adapter_name))
                    print('当前为：主网络状态')
            else:
                print('启用备用网络')
                self.set_gateway(self.backup_gateway)
            time.sleep(self.interval)
            # win32evtlogutil.AddSourceToRegistry("RonghuiTuring", "NetworkMonitorService")
            # win32evtlogutil.ReportEvent("RonghuiTuring",0,0,"NetworkMonitorService")

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

    def get_backup_gateway_ip(self, adapter_name):
        # 获取当前IP
        result = subprocess.run(f'netsh interface ip show config name="{adapter_name}"', shell=True,
                                capture_output=True, text=True)
        if " IP 地址:                           192.168.108" in result.stdout:
            print("获取备份网关信息", result.stdout)
            return 1  # 返回1

    def is_dhcp_enabled(self, adapter_name):
        result = subprocess.run(f'netsh interface ip show config name="{adapter_name}"', shell=True,
                                capture_output=True, text=True)
        print(result)
        return "DHCP 已启用:                          是" in result.stdout

    def release_renew_dhcp(self, adapter_name):
        # 释放并更新DHCP
        subprocess.call(f'ipconfig /release "{adapter_name}"', shell=True)
        time.sleep(5)  # 等待一段时间以确保释放
        subprocess.call(f'ipconfig /renew "{adapter_name}"', shell=True)
        time.sleep(5)  # 等待一段时间以确保续订
        servicemanager.LogInfoMsg(f"DHCP settings refreshed for adapter {adapter_name}")


if __name__ == '__main__':
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
