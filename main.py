import asyncio
import sys
from pathlib import Path

# noinspection PyPackageRequirements
import servicemanager
import win32event
import win32service
import win32serviceutil


# noinspection PyPep8Naming
class PythonService(win32serviceutil.ServiceFramework):
    """Let's create a Windows Service in Windows"""
    _svc_name_ = 'NucoderService'
    _svc_display_name_ = 'nucoder test service'
    _svc_description_ = 'Who knows'
    _exe_path_ = str(Path(sys.exec_prefix).joinpath('Scripts', 'pythonservice.exe'))

    def __init__(self, args):
        """Constructor of service"""
        win32serviceutil.ServiceFramework.__init__(self, args)
        self._stop_event = win32event.CreateEvent(None, 0, 0, None)
        self._event_loop = asyncio.SelectorEventLoop()
        asyncio.set_event_loop(self._event_loop)
        self.running = False

    def GetAcceptedControls(self):
        result = win32serviceutil.ServiceFramework.GetAcceptedControls(self)
        result |= win32service.SERVICE_ACCEPT_PRESHUTDOWN | win32service.SERVICE_ACCEPT_PARAMCHANGE
        return result

    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, "")
        )
        self._start_loop()

    def _start_loop(self):
        self._event_loop.create_task(self.random_sleep())
        self.running = True
        self._event_loop.run_forever()

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self._stop_event)
        self.running = False
        self._event_loop.call_soon_threadsafe(self._event_loop.stop)

    async def random_sleep(self):
        while self.running:
            await asyncio.sleep(5)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PythonService)
