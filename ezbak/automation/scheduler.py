"""Scheduled task registration pieces: task XML, schtasks argv, error text.

Task arguments, XML layout, file name and encoding are fixed by C3 (ADR 0003 §6).
Values in <Author>, <UserId>, <Command> and <Arguments> are XML-escaped (& < >),
so the registered text is unchanged. The XML file name uses the same rule as the
config JSON (schedule_config.safe_name); schtasks /tn gets the original task name.
Time, environment and home folder are read at call time: the caller passes `now`.
create_task/delete_task run schtasks through winsys.proc with the current kwargs
(delete keeps shell=True and no encoding). Message text and order are the current ones.
"""
import datetime as _dt
import locale
import logging
import os
import sys
from datetime import timedelta
from xml.sax.saxutils import escape

from ezbak.winsys import proc

from . import schedule_config

_log = logging.getLogger(__name__)


def _debug(text):
    # Default debug reporter of create_task: the diagnostic log (P6-5d; was print)
    _log.debug(text.removeprefix("DEBUG: "))

_HEADER = '''<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>{date}</Date>
    <Author>{run_as_user}</Author>
    <Description>ezBAK Automated Backup Task</Description>
  </RegistrationInfo>
  <Triggers>
'''

_DAILY = '''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
'''

_WEEKLY = '''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Sunday />
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
'''

_MONTHLY = '''    <CalendarTrigger>
      <StartBoundary>{start_boundary}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>{monthly_day}</Day>
        </DaysOfMonth>
        <Months>
          <January />
          <February />
          <March />
          <April />
          <May />
          <June />
          <July />
          <August />
          <September />
          <October />
          <November />
          <December />
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
'''

_FOOTER = '''  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{run_as_user}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{executable}</Command>
      <Arguments>{arguments}</Arguments>
    </Exec>
  </Actions>
</Task>'''

_TRIGGERS = {"DAILY": _DAILY, "WEEKLY": _WEEKLY, "MONTHLY": _MONTHLY}

_SCHEDULE_TYPES = {"daily": "DAILY", "weekly": "WEEKLY", "monthly": "MONTHLY"}

_ERROR_SOLUTIONS = {
    1: "General error. Check if task name is valid and you have admin privileges.",
    2: "Access denied. Try running as administrator.",
    267: "Invalid task name or path. Avoid special characters in task name.",
    2147942487: "Access denied. Administrator privileges required.",
    2147943711: "Object already exists. Task with this name already exists.",
    -2147024809: "Invalid parameter. Check task settings.",
    -2147216609: "Task Scheduler service is not available.",
}


def normalize_time(time_str):
    """'H:MM' -> 'HH:MM'; ValueError with the current message when not a valid time."""
    try:
        time_obj = _dt.datetime.strptime(time_str.strip(), "%H:%M")
    except ValueError:
        raise ValueError(f"Invalid time format '{time_str}'. Please use HH:MM format (e.g., 14:30)")
    return time_obj.strftime("%H:%M")


def schedule_type(schedule):
    """'Daily'/'Weekly'/'Monthly' (any case, spaces trimmed) -> schtasks type."""
    sc_type = _SCHEDULE_TYPES.get(schedule.lower().strip())
    if sc_type is None:
        raise ValueError(f"Unsupported schedule type: {schedule}. Use Daily, Weekly, or Monthly")
    return sc_type


def task_arguments(entry_script, config_file):
    """<Arguments> text; the script path is left out when running frozen (exe)."""
    if getattr(sys, 'frozen', False):
        return f'--schedule-config "{config_file}"'
    return f'"{entry_script}" --schedule-config "{config_file}"'


def run_as_user(environ):
    """DOMAIN\\User unless the domain is empty or equals the user name."""
    current_user = environ.get('USERNAME', 'Administrator')
    domain = environ.get('USERDOMAIN', environ.get('COMPUTERNAME', ''))
    if domain and domain.upper() != current_user.upper():
        return f"{domain}\\{current_user}"
    return current_user


def build_task_xml(executable, arguments, formatted_time, sc_type, run_as_user, monthly_day=1, *, now):
    """Task Scheduler XML; first run is tomorrow (from `now`) at `formatted_time`."""
    start_date = (now + timedelta(days=1)).strftime('%Y-%m-%d')
    start_boundary = f"{start_date}T{formatted_time}:00"
    user = escape(run_as_user)
    xml = _HEADER.format(date=now.isoformat(), run_as_user=user)
    trigger = _TRIGGERS.get(sc_type)
    if trigger is not None:
        xml += trigger.format(start_boundary=start_boundary, monthly_day=monthly_day)
    xml += _FOOTER.format(run_as_user=user, executable=escape(executable), arguments=escape(arguments))
    return xml


CREATED_AT_FORMAT = "%Y-%m-%d %H:%M:%S"  # created_at as schedule_config.save writes it
_MAX_MONTHS = 13  # a monthly day exists at least once a year (31 in January)


def _occurrences(sc_type, day_of_month, start):
    """Run dates of a trigger kind from `start`'s date on (DAILY, WEEKLY on Sunday as
    in the task XML, MONTHLY on day_of_month; a month without that day is skipped)."""
    day = start.date()
    if sc_type == "DAILY":
        for offset in range(2):
            yield day + timedelta(days=offset)
    elif sc_type == "WEEKLY":
        sunday = day + timedelta(days=(6 - day.weekday()) % 7)
        yield sunday
        yield sunday + timedelta(days=7)
    else:
        year, month = day.year, day.month
        for _ in range(_MAX_MONTHS + 1):
            try:
                yield _dt.date(year, month, day_of_month)
            except ValueError:
                pass  # no such day in this month: Task Scheduler skips the month too
            year, month = (year + 1, 1) if month == 12 else (year, month + 1)


def next_run(config, now):
    """Next start time of a schedule config dict (U7-1), or None when it cannot be told.

    Same rules as the registered task: the first run is the day after created_at at
    time_str (build_task_xml), Weekly runs on Sunday, Monthly on monthly_day. Tasks
    changed or removed outside ezBAK are not known here (decision D2 A).
    """
    try:
        hour, minute = map(int, normalize_time(str(config.get('time_str', ''))).split(':'))
        sc_type = schedule_type(str(config.get('schedule', '')))
        day_of_month = int(config.get('monthly_day', 1))
    except (ValueError, TypeError):
        return None
    start = now
    try:
        created = _dt.datetime.strptime(str(config.get('created_at')), CREATED_AT_FORMAT)
        first = _dt.datetime.combine(created.date() + timedelta(days=1), _dt.time(hour, minute))
        start = max(now, first)
    except ValueError:
        pass  # no or odd created_at: count from now
    for day in _occurrences(sc_type, day_of_month, start):
        candidate = _dt.datetime.combine(day, _dt.time(hour, minute))
        if candidate >= start:
            return candidate
    return None


def next_scheduled(configs, now):
    """(task_name, start time) of the earliest next run among config dicts, or None."""
    runs = [(run, str(config.get('task_name', ''))) for config in configs
            if isinstance(config, dict) and (run := next_run(config, now)) is not None]
    if not runs:
        return None
    run, name = min(runs)
    return name, run


def schedules_dir():
    """~\\AppData\\Local\\ezBAK\\schedules (home folder read at call time)."""
    return os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ezBAK', 'schedules')


def task_xml_path(task_name):
    """`<schedules_dir>/<safe_name(task_name)>_task.xml` (same rule as the config JSON)."""
    return os.path.join(schedules_dir(), f"{schedule_config.safe_name(task_name)}_task.xml")


def write_task_xml(task_name, xml):
    """Write `<safe_name(task_name)>_task.xml` as UTF-16 (BOM, CRLF) and return its path."""
    os.makedirs(schedules_dir(), exist_ok=True)
    xml_file = task_xml_path(task_name)
    with open(xml_file, 'w', encoding='utf-16') as f:
        f.write(xml)
    return xml_file


def build_create_cmd(task_name, xml_file):
    """schtasks argv that registers (or overwrites) a task from an XML file."""
    return ["schtasks", "/create", "/tn", task_name, "/xml", xml_file, "/f"]


def explain_error(return_code, stderr, stdout):
    """Error text for a failed schtasks run: code, hints, then the raw output."""
    error_msg = f"schtasks failed with return code {return_code}"
    if stderr:
        error_msg += _stderr_hint(stderr.lower())
    if return_code in _ERROR_SOLUTIONS:
        error_msg += f"\n\n{_ERROR_SOLUTIONS[return_code]}"
    if return_code in [2, 2147942487] or "access" in (stderr or "").lower():
        error_msg += (
            "\n\nAdditional Info:"
            "\n- Make sure ezBAK is running as Administrator"
            "\n- The scheduled task will run with HIGHEST privilege level"
            "\n- Check Windows Task Scheduler service is running"
        )
    if stderr:
        error_msg += f"\n\nSystem Error: {stderr}"
    if stdout:
        error_msg += f"\n\nSystem Output: {stdout}"
    return error_msg


def _stderr_hint(stderr_lower):
    """First matching 'Solution:' line for known stderr patterns, else ''."""
    if "access" in stderr_lower and "denied" in stderr_lower:
        return "\n\nSolution: Run the application as Administrator"
    if "invalid" in stderr_lower and "time" in stderr_lower:
        return "\n\nSolution: Check time format (use HH:MM, e.g., 14:30)"
    if "task" in stderr_lower and "exist" in stderr_lower:
        return "\n\nSolution: Choose a different task name or delete existing task"
    if "privilege" in stderr_lower or "permission" in stderr_lower:
        return "\n\nSolution: Administrator privileges required for task creation"
    if "user" in stderr_lower and ("not found" in stderr_lower or "invalid" in stderr_lower):
        return "\n\nSolution: Check user account format (try DOMAIN\\Username or just Username)"
    if "service" in stderr_lower and "not" in stderr_lower:
        return "\n\nSolution: Task Scheduler service may not be running. Try: net start schedule"
    return ""


def build_delete_cmd(task_name):
    """schtasks argv that deletes a task without asking."""
    return ["schtasks", "/Delete", "/TN", task_name, "/F"]


def delete_task(task_name):
    """Run schtasks /Delete; RuntimeError(stderr or stdout) when it fails.

    After a successful delete the task XML is removed too (missing file ignored,
    other OSError only logged); the config JSON is the caller's (schedule_config.delete).
    """
    res = proc.run(build_delete_cmd(task_name), shell=True, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(res.stderr or res.stdout)
    _remove_task_xml(task_name)


def _remove_task_xml(task_name):
    """Remove the task XML; FileNotFoundError ignored, other OSError logged as a warning."""
    try:
        os.remove(task_xml_path(task_name))
        _log.info("delete_task xml_removed=True task=%s", task_name)
    except FileNotFoundError:
        pass
    except OSError as e:
        _log.warning("delete_task xml_remove_failed task=%s error=%s: %s",
                     task_name, type(e).__name__, e)


def create_task(spec, entry_script, *, out, now, debug=_debug):
    """Save the config, write the XML, run schtasks /create; returns the success text.

    `out` gets the detailed log lines, `debug` the DEBUG lines (default: diagnostic log). Failures raise as today:
    ValueError (time, kind), RuntimeError (config, XML, schtasks code, timeout).
    When the XML or schtasks step fails, the config goes back to its previous content
    (removed when there was none) and the XML written here is removed; after a timeout
    both are kept, since the task may have been registered.
    """
    formatted_time = normalize_time(spec.time_str)
    debug(f"DEBUG: Time format validated: {spec.time_str} -> {formatted_time}")
    out(f"Time format validated: {spec.time_str} -> {formatted_time}")
    sc_type = schedule_type(spec.schedule)
    config_file, prev_config = _save_config(spec, out=out, now=now)
    executable = sys.executable
    arguments = task_arguments(entry_script, config_file)
    out(f"Task executable: {executable}")
    out(f"Task arguments: {arguments}")
    out(f"Schedule config saved to: {config_file}")
    out(f"Creating scheduled task with retention_count={spec.retention_count}, "
        f"log_retention_days={spec.log_retention_days}")
    user = run_as_user(os.environ)
    xml = build_task_xml(executable, arguments, formatted_time, sc_type, user, spec.monthly_day, now=now)
    xml_file = None
    try:
        xml_file = _write_xml(spec.task_name, xml, out=out)
        cmd_args = build_create_cmd(spec.task_name, xml_file)
        out(f"schtasks command (XML-based): {' '.join(cmd_args)}")
        out(f"Task will run as: {user} with HIGHEST privileges")
        out("Using XML file to avoid command length limitations")
        _run_create(cmd_args, out=out, debug=debug)
    except RuntimeError as e:
        if isinstance(e.__context__, proc.TimeoutExpired):  # may be registered: keep both (F3)
            _log.info("create_task timeout config_kept=True task=%s", spec.task_name)
        else:
            _rollback(spec.task_name, prev_config, xml_file)
        raise
    out(f"Task creation successful: {spec.task_name} (with admin privileges)")
    debug("DEBUG: Returning success message from create_scheduled_task")
    return _success_text(spec, formatted_time, user)


def _rollback(task_name, prev_config, xml_file):
    """Config back to `prev_config` (removed when None), this run's XML removed; OSError only logged."""
    try:
        schedule_config.restore_raw(task_name, prev_config)
        if xml_file is not None and os.path.exists(xml_file):
            os.remove(xml_file)
        _log.info("create_task rolled_back=True task=%s had_config=%s",
                  task_name, prev_config is not None)
    except OSError as e:
        _log.warning("create_task rollback_failed task=%s error=%s: %s",
                     task_name, type(e).__name__, e)


def _save_config(spec, *, out, now):
    try:
        prev_config = schedule_config.read_raw(spec.task_name)
        config_file = schedule_config.save(spec, now=now)
        out(f"Schedule configuration saved to: {config_file}")
        return config_file, prev_config
    except Exception as e:
        error_msg = f"Failed to save schedule configuration: {str(e)}"
        out(error_msg)
        raise RuntimeError(error_msg) from e


def _write_xml(task_name, xml, *, out):
    try:
        xml_file = write_task_xml(task_name, xml)
        out(f"Task XML file created: {xml_file}")
        return xml_file
    except Exception as e:
        error_msg = f"Failed to create task XML file: {str(e)}"
        out(error_msg)
        raise RuntimeError(error_msg) from e


def _run_create(cmd_args, *, out, debug):
    """schtasks /create with the system encoding read now; RuntimeError on failure."""
    system_encoding = locale.getpreferredencoding()
    try:
        debug(f"DEBUG: Executing schtasks command: {' '.join(cmd_args)}")
        result = proc.run(cmd_args, check=False, shell=False, capture_output=True, text=True,
                          encoding=system_encoding, errors='replace', timeout=180)
        debug(f"DEBUG: schtasks command completed with returncode={result.returncode}")
        out(f"schtasks return code: {result.returncode}")
        if result.stdout:
            out(f"schtasks stdout: {result.stdout}")
        if result.stderr:
            out(f"schtasks stderr: {result.stderr}")
        if result.returncode != 0:
            raise RuntimeError(explain_error(result.returncode, result.stderr, result.stdout))
        debug("DEBUG: schtasks command succeeded (returncode=0)")
    except proc.TimeoutExpired:
        out("schtasks command timed out")
        raise RuntimeError("Task creation timed out. The system may be busy.")


def _success_text(spec, formatted_time, user):
    return (
        f"Scheduled task '{spec.task_name}' created successfully!\n\n"
        f"Schedule: {spec.schedule} at {formatted_time}\n"
        f"User: {spec.user}\n"
        f"Destination: {spec.dest}\n"
        f"Retention: {spec.retention_count} backups, {spec.log_retention_days} days logs\n"
        f"Hidden files: {'Included' if spec.include_hidden else 'Excluded'}\n"
        f"System files: {'Included' if spec.include_system else 'Excluded'}\n\n"
        f"Run Level: Administrator (Highest Privileges)\n"
        f"Run As: {user}"
    )
