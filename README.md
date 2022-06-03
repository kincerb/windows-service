# windows-service

This is sample repo for setting up a Windows service from a python script. We've run into issues in the upgrade from 
python 3.7 to 3.10, and have found out painfully that the new *asyncio* default loop policy does not work when being
used as a service.

Starting with python 3.8, the default loop policy has changed from *SelectorEventLoop* to *ProactorEventLoop*.

The issue is due to the *ProactorEventLoop* not supporting *add_reader/writer*. My guess is this is needed to
properly support detaching from a process, because you can use the *ProactorEventLoop* to run the service in *debug*,
which stays attached to the process.

## Links to related issues

- [cpython github issue page](https://github.com/python/cpython/issues/81554)
- [tornadoweb](https://github.com/tornadoweb/tornado/issues/2608)

## Getting started

After cloning this repo, install [pywin32](https://pypi.org/project/pywin32/) and follow some post-installation steps.

### Install pywin32

```powershell
.\venv\Scripts\activate.ps1
python -m pip install -r .\requirements.txt
```

As an *Administrator*, you must run the post-install command to place the proper DLLs into the system python.

```powershell
python .\venv\Scripts\pywin32_postinstall.py -install
```

### Modify to run service properly from virtual environment

Copy the `main.py` script into the `Scripts` directory.

```powershell
Copy-Item .\main.py -Destination .\venv\Scripts\
```

Copy the `pythonservice.exe` to the `Scripts` directory, next to the python script that will be run as the service.

```powershell
Copy-Item .\venv\Lib\site-packages\win32\pythonservice.exe -Destination .\venv\Scripts\
```

Lastly, set the `_exe_path_` class attribute to the `pythonservice.exe` placed inside `Scripts`. This is needed to
ensure the virtual environment is used for the service.

```python
_exe_path_ = str(Path(sys.exec_prefix).joinpath('Scripts', 'pythonservice.exe'))
```

## Managing service

```powershell
python .\venv\Scripts\main.py [ install | remove | start | debug ]
```