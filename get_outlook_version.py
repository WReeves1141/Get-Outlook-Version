"""
This script gets the installed Office app version(s).
It is designed for Windows workstations.

It returns a dictionary with the outlook version,
or it returns the error message.

"""

__author__ = "Walter Reeves"

from winreg import HKEY_CLASSES_ROOT, OpenKey, QueryValueEx


def entrypoint() -> dict:
    """Gets the outlook version from the Windows registry.
       Returns a dictionary with the version values."""

    version = {}

    try:
        key = OpenKey(HKEY_CLASSES_ROOT, r"Outlook.Application\CurVer")
        value, _ = QueryValueEx(key, None)

        version["outlook_version"] = value

    except Exception as e:
        version["error"] = f"{e}"

    return version


# Testing
# print(get_outlook_version())
