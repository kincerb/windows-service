import socket
import subprocess
import sys
import time
from pathlib import Path

# noinspection PyPackageRequirements
import servicemanager
import win32api
import win32event
import win32process
import win32service
import win32serviceutil

HERE = Path().cwd()


def service_control(args):
    """Service control called by the command line to interact with
    Win32 services libraries.
    """
    if 'action' not in args:
        # If there was no action we were invoked by the Windows SCM
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PythonService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        args = _reassemble_command_args(args)
        win32serviceutil.HandleCommandLine(PythonService, argv=args)


def _reassemble_command_args(args):
    """Turn the Windows-compatible args back into a command line suitable for win32serviceutil.HandleCommandLine."""
    args = {k: v for k, v in args.items() if v is not None}
    # noinspection PyProtectedMember
    new_command_args = [PythonService._svc_name_]
    if args['action'] in ('start', 'restart', 'stop'):
        if 'wait' in args:
            new_command_args.append('--wait={}'.format(args['wait']))
        else:
            new_command_args.append('--wait=60')

    if args['action'] in ('install', 'update'):
        if 'interactive' in args:
            new_command_args.append('--interactive')

        for option in ('perfmonini', 'perfmondll', 'username', 'password'):
            if option in args:
                new_command_args.append('--{}={}'.format(option, args[option]))

        if 'startup' in args:
            new_command_args.append('--startup={}'.format(args['startup']))
        else:
            new_command_args.append('--startup=auto')

    new_command_args.append(args['action'])
    return new_command_args


class PythonService(win32serviceutil.ServiceFramework):
    """Let's create a Windows Service in Windows"""
    _svc_name_ = 'NucoderService'
    _svc_display_name_ = 'nucoder test service'
    _svc_description_ = 'Who knows'
    _exe_path_ = str(Path(sys.exec_prefix).joinpath('Scripts', 'pythonservice.exe'))

    def __init__(self, args):
        """Constructor of service"""
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.pid = None

    def GetAcceptedControls(self):
        result = win32serviceutil.ServiceFramework.GetAcceptedControls(self)
        result |= win32service.SERVICE_ACCEPT_PRESHUTDOWN | win32service.SERVICE_ACCEPT_PARAMCHANGE
        return result

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        if self.pid is None:
            return
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )
        # proc = subprocess.Popen(str(HERE.joinpath('spin.py')), close_fds=True, creationflags=win32process.DETACHED_PROCESS)
        # self.pid = proc.pid
        # proc.communicate()
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        # proc.terminate()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STOPPED,
            (f"{self._svc_name_} stopped", "")
        )
        self.pid = None


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PythonService)
