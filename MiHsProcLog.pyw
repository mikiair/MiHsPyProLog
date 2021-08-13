#!/usr/bin/env python3

__author__ = "Michael Heise"
__copyright__ = "Copyright (C) 2021 by Michael Heise"
__license__ = "Apache License Version 2.0"
__version__ = "0.0.3"
__date__ = "08/13/2021"

"""Configurable logging and limiting of program usage under Windows
"""

#    Copyright 2021 Michael Heise (mikiair)
#
#    Licensed under the Apache License, Version 2.0 (the "License");
#    you may not use this file except in compliance with the License.
#    You may obtain a copy of the License at
#
#        http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS,
#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#    See the License for the specific language governing permissions and
#    limitations under the License.

# standard imports
from datetime import datetime, timedelta, date
import configparser
import argparse
import pathlib
import sys
import time
import json
import signal
import ctypes
import threading


# 3rd party imports
import psutil

# local imports


logformat = "{0} - {1}\n"

checkIntervalSec = None
intervalsBetweenWarnings = None
numWarningRepetitions = None


def writeLogMsg(message):
    if not log:
        return

    log.write(
        logformat.format(datetime.now().isoformat(sep=" ", timespec="seconds"), message)
    )
    log.flush()


validExpiredActions = "log|warn_once|warn_repeat|warn_kill|kill".split("|")


def readConfig():
    try:
        writeLogMsg(f"Reading configuration from '{args.configfile}'")

        if not args.configfile.exists():
            raise FileNotFoundError("File not found")

        config = configparser.ConfigParser(allow_no_value=True, delimiters=("="))

        config.read(args.configfile)

        global checkIntervalSec
        global intervalsBetweenWarnings
        global numWarningRepetitions
        
        if config["Options"]:
            checkIntervalSec = config["Options"].getint("CheckIntervalSec")
            intervalsBetweenWarnings = config["Options"].getint("IntervalsBetweenWarnings")
            numWarningRepetitions = config["Options"].getint("NumWarningRepetitions")
        else:
            writeLogMsg("No [Options] section in configuration file, using defaults.")
            checkIntervalSec = 60
            intervalsBetweenWarnings = 1
            numWarningRepetitions = 3
            
        if config["Processes"]:
            processes_to_log = {
                pn.lower(): opt.split(",") if opt else None
                for (pn, opt) in config["Processes"].items()
            }
            for (pn, opt) in processes_to_log.items():
                if opt:
                    try:
                        opt[0] = int(opt[0])
                    except:
                        raise ValueError(f"Invalid time limit '{opt[0]}' for process '{pn}'")

                    if len(opt) == 1:
                        opt.append(0)
                    elif len(opt) == 2:
                        try:
                            opt[1] = validExpiredActions.index(opt[1])
                        except:
                            raise ValueError(f"Invalid option '{opt[1]}' for process '{pn}'")
                    else:
                        raise ValueError(f"Invalid options '{opt}' for process '{pn}'")
        else:
            raise NameError("Invalid configuration")

        if not processes_to_log:
            raise Exception("No processes to log")

        # TODO validate configuration: int usage times? valid options?

    except Exception as e:
        writeLogMsg(f"Reading configuration failed! ({e})")
        sys.exit(-1)

    return processes_to_log


def readTodaysUsage():
    """Read usage file: date, if not today --> return empty dictionary
    else return dictionary with keys = process names, values = {"usetime", "expired", "laststart", "lastend", "active"}
    """
    stateFilePath = pathlib.Path(stateFileName)
    if stateFilePath.exists():
        writeLogMsg("Reading state file with usage statistics.")
        try:
            with open(stateFilePath, "r", encoding="utf8") as stateFile:
                stateFileDateStr, pu = json.load(stateFile)
                if not date.fromisoformat(stateFileDateStr) == date.today():
                    writeLogMsg("State file is not of today, reset usage data.")
                    pu = {}
                else:
                    for (pn, pud) in pu.items():
                        pud["expired"] = 0
        except Exception as e:
            writeLogMsg(f"Reading process usage from state file failed! ({e})")
            sys.exit(-1)
        finally:
            if stateFile:
                stateFile.close()
    else:
        writeLogMsg("No state file with process usage found.")
        pu = {}
    return pu


def getActiveMatches(processes_to_log):
    """Return all matching processes by id"""
    return {
        p.pid: p.info
        for p in psutil.process_iter(["name", "create_time"])
        if p.info["name"].lower() in processes_to_log
    }


def getMatchingActiveProcesses(processes_to_log):
    # writeLogMsg("getMatchingActiveProcesses")

    active_matches = getActiveMatches(processes_to_log)

    active = {}
    for (pid, info) in active_matches.items():
        pn = info["name"]
        ct = datetime.fromtimestamp(info["create_time"])

        if pn not in active:
            # new process name
            active[pn] = {"cdatetime": max(ct, service_start)}
        else:
            # identical process name
            if ct > service_start and ct < active[pn]["cdatetime"]:
                # ...but started earlier
                active[pn]["cdatetime"] = ct

    return active


def logProcessesStartedBefore(active_proc):
    # writeLogMsg("logProcessesStartedBefore")
    for (pn, pd) in active_proc.items():
        writeLogMsg(f"Process '{pn}' has been started before.")


def logChanges(last_proc, active_proc):
    # writeLogMsg("logChanges")
    for (pn, pd) in last_proc.items():
        if not pn in active_proc:
            writeLogMsg(f"Process '{pn}' has ended.")
            pd["active"] = False

    for (pn, pd) in active_proc.items():
        if not pn in last_proc:
            writeLogMsg(f"Process '{pn}' started.")


def updateProcessUsage(process_usage, last_proc, active_proc):
    # writeLogMsg("updateProcessUsage")
    if last_proc:
        for (pn, pd) in last_proc.items():
            if pn in process_usage and "active" in pd and not pd["active"]:
                process_usage[pn]["active"] = False
                process_usage[pn]["lastend"] = time_now

    for (pn, pd) in active_proc.items():
        if pn in process_usage:
            # process was running today, and has started again or is still running
            pu = process_usage[pn]
            if pu["active"]:
                # still running
                pu["usetime"] += inc_time / to_minutes
            else:
                # re-started
                pu["active"] = True
                pu["laststart"] = pd["cdatetime"]
                pu["lastend"] = None
        else:
            # new process, running the first time today
            process_usage[pn] = {
                "usetime": 0.0,
                "expired": 0,
                "laststart": pd["cdatetime"],
                "lastend": None,
                "active": True,
            }


# see https://stackoverflow.com/questions/36440917/how-to-make-a-messagebox-auto-close-in-several-seconds-by-python

def worker(title,close_until_seconds):
    time.sleep(close_until_seconds)
    wd = ctypes.windll.user32.FindWindowW(None, title)
    if wd > 0:
        ctypes.windll.user32.SendMessageW(wd, 0x0010, 0, 0)
    return


def AutoCloseMessageBoxW(owner, text, title, options, close_until_seconds):
    t = threading.Thread(target=worker,args=(title,close_until_seconds))
    t.start()
    ctypes.windll.user32.MessageBoxW(owner, text, title, options)
    
    
def killAllProcesses(process_name):
    writeLogMsg(f"--> Killing processes '{process_name}'...")
    
    for p in psutil.process_iter():
        try:
            if p.name().lower() == process_name:
                p.kill()
        except e:
            writeLogMsg("    Could not kill a process. Access denied?")
            
    writeLogMsg(f"... all processes named '{process_name}' killed.")

    
def evalProcessUsage(processes_to_log, process_usage):
    # writeLogMsg("evalProcessUsage")
    for (pn, pud) in process_usage.items():
        process_options = processes_to_log[pn]

        if process_options:
            process_time_limit = process_options[0]
            process_expired_mode = process_options[1]
        
            if pud["active"] and pud["usetime"] > process_time_limit:
                if pud["expired"] == 0:
                    writeLogMsg(
                        f"--> Today's usage limit of {processes_to_log[pn][0]} minutes for process '{pn}' has expired!"
                    )
                    
                if process_expired_mode > 0:
                    if (
                        (process_expired_mode == 1 and pud["expired"] == 0) or
                        (process_expired_mode == 2 and (numWarningRepetitions == 0 or pud["expired"] < numWarningRepetitions)) or
                        (process_expired_mode == 3 and pud["expired"] < numWarningRepetitions) 
                        ):
                        pem3 = process_expired_mode == 3
                        if pem3:
                            closeInMinutes = (numWarningRepetitions - pud["expired"]) * intervalsBetweenWarnings * checkIntervalSec / 60
                        AutoCloseMessageBoxW(
                            None,
                            f"Today's usage limit of {process_time_limit} minutes for process '{pn}' has expired!" +
                            " Please save your work...\n" +
                            (f"[Application will be closed in {closeInMinutes:3.1f} minutes!]\n" if pem3 else "") +
                            f"\nDie heutige Nutzungsdauer von {process_time_limit} Minuten für den Prozess '{pn}' ist abgelaufen!" +
                            " Bitte Dateien sichern..." +
                            (f"\n[Anwendung wird in {closeInMinutes:3.1f} Minuten geschlossen!]" if pem3 else ""),
                            f"{pn} - Usage warning / Nutzungswarnung ({pud['expired']})",
                            0x11030,
                            int((0.5 if process_expired_mode > 1 else 1) * intervalsBetweenWarnings * checkIntervalSec)
                        )
                    if ((process_expired_mode == 3 and pud["expired"] == numWarningRepetitions) or
                        (process_expired_mode == 4 and pud["expired"] == 0)
                        ):
                        AutoCloseMessageBoxW(
                            None,
                            f"Application '{pn}' will be closed NOW!\n" +
                            f"Die Anwendung '{pn}' wird JETZT geschlossen!]",
                            f"{pn} - Usage warning / Nutzungswarnung ({pud['expired']})",
                            0x11030,
                            int(0.75 * intervalsBetweenWarnings * checkIntervalSec)
                        )
                        killAllProcesses(pn)
                    pud["expired"] += 1
                else:
                    pud["expired"] = 1

def formatDateTimeForJSON(obj):
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()


def writeTodaysUsage(process_usage):
    if len(process_usage) == 0:
        return

    try:
        stateFile = open(pathlib.Path(stateFileName), "w", encoding="utf8")
        json.dump(
            (date.today(), process_usage), stateFile, default=formatDateTimeForJSON
        )
    except Exception as e:
        writeLogMsg(f"Writing process usage to state file failed! ({e})")
        sys.exit(-1)
    finally:
        if stateFile:
            stateFile.close()


def sigterm_handler(_signo, _stack_frame):
    """clean exit on SIGTERM signal (when systemd stops the process)"""
    sys.exit(0)


try:
    log = None
    process_usage = None

    to_minutes = timedelta(minutes=1)

    stateFileName = "./MiHsProcLog.state"

    # install handler
    signal.signal(signal.SIGTERM, sigterm_handler)

    # start logging
    log = open(pathlib.Path("./MiHsProcLog.log"), "a")
    writeLogMsg("MiHsProcLog started.")

    # define commandline arguments
    parser = argparse.ArgumentParser(
        description="Configurable logging of program usage under Windows"
    )
    parser.add_argument(
        "configfile",
        default="./MiHsProcLog.cfg",
        nargs="?",
        type=pathlib.Path,
        help="specify a path to the configuration file to use",
    )

    # collect commandline arguments
    args = parser.parse_args()

    processes_to_log = readConfig()

    service_start = datetime.now()
    last_now = service_start

    process_usage = readTodaysUsage()

    last_proc = None

    while True:
        time_now = datetime.now()
        inc_time = time_now - last_now

        active_proc = getMatchingActiveProcesses(processes_to_log)

        if last_proc:
            logChanges(last_proc, active_proc)
        else:
            logProcessesStartedBefore(active_proc)

        updateProcessUsage(process_usage, last_proc, active_proc)

        evalProcessUsage(processes_to_log, process_usage)

        writeTodaysUsage(process_usage)

        last_proc = active_proc
        last_now = time_now

        time.sleep(checkIntervalSec)

except Exception as e:
    writeLogMsg(f"Unhandled exception: {e}")
finally:
    writeLogMsg("Stopping MiHsProcLog...")
    if log:
        log.close()
