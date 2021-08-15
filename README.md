# MiHsPyProLog
**Configurable logging and limitation of program usage with Python under Windows**

This project contains a small Python script which can be triggered to run with Windows Task Scheduler. A configuration file is used to define a list of processes to be watched, and optionally daily usage time can be restricted. In case the usage limit is exceeded, warning messages may be displayed a number of times, and finally the expired process may be killed.

## Installation
1. Download this repository / latest release.
2. Unpack the files to a folder you (and any other user) has write access to.
3. Edit the `MiHsPyProLog.cfg` configuration file to fit your needs (see below).
4. [optional] If you want to limit program usage on a per-user basis, create and edit a `MiHsPyProLog.cfg` file in each user's application data folder. File should be 
`%LOCALAPPDATA%\MiHs\MiHsPyProLog\MiHsPyProLog.cfg`.
6. Start Task Scheduler and import the task template file `MiHsPyProLog_ScheduledTask.xml`.
7. Edit the entries under the *Action* tab: Replace `your-PYTHON3-path` and `your-file-location` by the respective folders on your system.
8. Reboot.

## Requirements
Requires [psutil package](https://pypi.org/project/psutil/#files) to be installed.

```python
> pip install psutil
```

## Configuration
The configuration file contains up to two sections. One mandatory '[Processes]' section and an optional '[Options]' section.

* Processes

  Each process to watch is listed here by its filename including the extension, e.g. `calc.exe`. If any daily time limitation applies, it must be added as a value given in minutes, like `calc.exe=1` for a usage limit of 1 minute per day. After the time limit another parameter may be added (separated by comma). It describes the action to take when the time limit is exceeded. The following values are recognized:
  - log         - log to file only (default)
  - warn        - repeatedly show a warning message box
  - warn_kill   - show warning several times and then kill expired process 
  - kill        - kill process immediately without warning

  The warning messages will be closed automatically after a certain time, if not done by user.
  
* Options
  This section can contain three general settings, valid for all processes.
  - `CheckIntervalSec` defines the time interval in seconds which is used for updating the internal data of running processes.
  - `IntervalsBetweenWarnings` is a multiplier which determines the time between consecutive warning messages.
  - `NumWarningRepetitions` is the number of warnings which will be displayed when `warn` or `warn_kill` are selected for a process. If this is zero (0), the warning messages will be displayed endlessly together with `warn`. For `warn_kill`, however, this will be interpreted as immediate kill!

## Log file
A log file `MiHsPyProLog.log` is created the folder where the Python file resides. It is written to when the task is active and running, requiring write access for any user to that folder. The file contains data when a process under surveillance is started and stopped by the user, and any usage expiry is logged. It also states when a process is killed by the script.
