# -*- coding: utf-8 -*-

"""
save the assembly as .jt
"""

import sys

import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Runtime.InteropServices as SRI
from System import Console
from System.IO import Directory
from System.IO.Path import Combine

import jt

__project__ = "save_as_jt"
__author__ = "recs"
__version__ = "0.0.0"
__update__ = "2020-11-13"


def raw_input(message):
    Console.WriteLine(message)
    return Console.ReadLine()


def is_exist(path_to_check):
    return Directory.Exists(path_to_check)


def makedirs(path_to_make):
    Directory.CreateDirectory(path_to_make)


def userprofile():
    return System.Environment.GetEnvironmentVariable("USERPROFILE")


def username():
    return System.Environment.UserName


def combine(path1, path2):
    return Combine(path1, path2)


def main():
    try:
        application = SRI.Marshal.GetActiveObject("SolidEdge.Application")

        assembly = application.ActiveDocument
        print("Part: %s" % assembly.Name)
        assert assembly.Type == 3, "This macro only works on .asm document." 

        jt_filename = assembly.name[:-4] + ".jt"
        print("JT name : %s" % jt_filename)

        path_download = userprofile() + "\\Downloads" + "\\solidedgeJT\\"
        if not is_exist(path_download):
            makedirs(path_download)
        new_name = combine(path_download, jt_filename)
        print("wait...")
        jt.save_as_jt(assembly, new_name)
        print(r"The downloaded file is in: %s", new_name)

    except AssertionError as err:
        print(err.args)

    except Exception as ex:
        print(ex.args)

    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def confirmation(func):
    response = raw_input("""Would you like to make a .jt file of this assembly? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes", "oui"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    print(
        "%s\n--author:%s --version:%s --last-update :%s"
        % (__project__, __author__, __version__, __update__)
    )
    confirmation(main)
