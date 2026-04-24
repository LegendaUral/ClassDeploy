from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))


def _session_id() -> int:
    if not sys.platform.startswith("win"):
        return 0
    try:
        import ctypes
        sid = ctypes.c_ulong(0)
        ctypes.windll.kernel32.ProcessIdToSessionId(
            ctypes.windll.kernel32.GetCurrentProcessId(),
            ctypes.byref(sid),
        )
        return int(sid.value)
    except Exception:
        return 0


def _acquire_single_instance() -> bool:
    if not sys.platform.startswith("win"):
        return True
    try:
        import ctypes
        name = f"Local\\ClassDeployAgentSingleton_{_session_id()}"
        handle = ctypes.windll.kernel32.CreateMutexW(None, False, name)
        err = ctypes.windll.kernel32.GetLastError()
        if err == 183:
            return False
        globals()["_mutex_handle"] = handle
        return True
    except Exception:
        return True


def run_plain():
    import asyncio
    from agent.agent import Agent
    if not _acquire_single_instance():
        print("ClassDeploy Agent is already running")
        return
    print("ClassDeploy Agent (console mode, Ctrl+C to exit)")
    try:
        asyncio.run(Agent().run_forever())
    except KeyboardInterrupt:
        print("Stopped")


# Optional Windows service support. The main install flow no longer uses it,
# because screen capture and input control must run in the interactive user session.
try:
    import servicemanager  # type: ignore
    import win32event  # type: ignore
    import win32service  # type: ignore
    import win32serviceutil  # type: ignore
    import asyncio
    import threading
    from agent.agent import Agent

    class ClassDeployAgentSvc(win32serviceutil.ServiceFramework):
        _svc_name_ = "ClassDeployAgent"
        _svc_display_name_ = "Class Deploy Agent"
        _svc_description_ = "Background helper service for ClassDeploy."

        def __init__(self, args):
            super().__init__(args)
            self._stop_evt = win32event.CreateEvent(None, 0, 0, None)
            self._loop = None

        def SvcStop(self):
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            win32event.SetEvent(self._stop_evt)
            if self._loop:
                self._loop.call_soon_threadsafe(self._loop.stop)

        def SvcDoRun(self):
            servicemanager.LogInfoMsg("ClassDeployAgent service starting")

            def _runner():
                self._loop = asyncio.new_event_loop()
                asyncio.set_event_loop(self._loop)
                try:
                    self._loop.run_until_complete(Agent().run_forever())
                except Exception as exc:
                    try:
                        servicemanager.LogErrorMsg(f"ClassDeployAgent error: {exc}")
                    except Exception:
                        pass

            threading.Thread(target=_runner, daemon=True).start()
            win32event.WaitForSingleObject(self._stop_evt, win32event.INFINITE)


    def run_service():
        if len(sys.argv) > 1:
            win32serviceutil.HandleCommandLine(ClassDeployAgentSvc)
            return
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(ClassDeployAgentSvc)
        servicemanager.StartServiceCtrlDispatcher()

except Exception:
    ClassDeployAgentSvc = None

    def run_service():
        run_plain()


def main():
    # Default mode is always plain interactive process.
    # Service mode is kept only for explicit service commands.
    service_cmds = {"install", "update", "remove", "start", "stop", "restart", "debug"}
    if len(sys.argv) > 1 and sys.argv[1].lower() in service_cmds:
        run_service()
    else:
        run_plain()


if __name__ == "__main__":
    main()
