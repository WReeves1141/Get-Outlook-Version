"""
This script gets the installed Office app version(s).
It is designed for Windows workstations.

It returns a dictionary with the office version,
or it returns the error message.

"""

__author__ = "Walter Reeves"

from winreg import (CloseKey, EnumKey, HKEY_CLASSES_ROOT, HKEY_LOCAL_MACHINE,
                    OpenKey, QueryInfoKey, QueryValueEx)


def entrypoint() -> dict:
    """Gets the outlook version from the Windows registry.
       Returns a dictionary with the version values."""

    search_locations = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    results = {}

    # Change as needed
    program = "Office"

    # Gets the office program version(s).
    for subkey_path in search_locations:
        key = OpenKey(HKEY_LOCAL_MACHINE, subkey_path)

        for i in range(QueryInfoKey(key)[0]):
            subkey = OpenKey(key, EnumKey(key, i))

            try:
                display_name = QueryValueEx(subkey, "DisplayName")[0]

                if program.lower() in display_name.lower() and '{' not in EnumKey(key, i):
                    results[EnumKey(key, i)] = display_name

            except Exception as e:
                results["office_error"] = e

            CloseKey(subkey)
        CloseKey(key)

    if len(results.keys()) > 1 and "office_error" in results.keys():
        results.pop("office_error")

    try:
        key = OpenKey(HKEY_CLASSES_ROOT, r"Outlook.Application\CurVer")
        value, _ = QueryValueEx(key, None)

        results["outlook_version"] = value

    except Exception as e:
        results["outlook_error"] = f"{e}"

    return results


# Testing
# Note: Outlook  must be installed for this to work.
print(entrypoint())
