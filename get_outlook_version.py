"""
This script gets the installed Office app version(s).
It is designed for Windows workstations.

It returns a dictionary with the outlook version,
or it returns the error message.

"""

from winreg import HKEY_CLASSES_ROOT, OpenKey, QueryValueEx

__author__ = "Walter Reeves"


def entrypoint() -> dict:
    """Gets the outlook version from the Windows registry."""
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
