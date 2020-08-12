# -*- coding: utf-8 -*-

"""
hide hardware and save the assembly as .jt or parasolid
"""

import sys

import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import jt

import System
import System.Runtime.InteropServices as SRI
from System import Console
from System.IO import Directory
from System.IO.Path import Combine


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


def generate_jt():
    try:
        application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
        print("Author: recs@premiertech.com")
        print("Last update: 2020-06-29")
        print("version solidedge: %s" % application.Value)

        assert application.Value in [
            "Solid Edge ST7",
            "Solid Edge 2019",
        ], "Unvalid version of solidedge"

        user = username()
        print("\nUser: %s" % user)
        if user.lower() in [
            "alba",
            "bouc11",
            "lapc3",
            "peld6",
            "fouj3",
            "cotk2",
            "nunk",
            "beam",
            "boum3",
            "morm8",
            "benn2",
            "recs",
            "gils2",
            "albp",
            "tres2",
        ]:
            print("Autorized user ID")
        else:
            print("user with no valid permissions.")
            sys.exit()

        asm = application.ActiveDocument
        print("part: %s\n" % asm.Name)

        # asm.Type =>  plate :4 , assembly : 3, partdocument: 1
        assert asm.Type == 3, "This macro only works on .asm not %s" % asm.Name[-4:]

        # --------------------
        selectSet = application.ActiveSelectSet
        selectSet.SuspendDisplay()
        selectSet.RemoveAll()

        objQueries = asm.Queries

        if objQueries.item("Hardware"):
            # Check hardware exists.
            objQuery = objQueries.Item("Hardware")
            print("Query: Hardware activated")
        else:
            # TODO: [1] fix if no Query hardware exist
            # create a new query
            objQuery = objQueries.Add("Hardware")
            objQuery.Scope = SolidEdgeAssembly.QueryScopeConstants.seQueryScopeAllParts
            objQuery.SearchSubassemblies = True

            # String to be set as "Value"
            createria = "HARDWARE"
            # Add Criteria to above query
            objQuery.AddCriteria(
                SolidEdgeAssembly.QueryPropertyConstants.seQueryPropertyName,
                "Category",
                SolidEdgeAssembly.QueryConditionConstants.seQueryConditionContains,
                createria,
            )
            print("Query: Hardware created")

        # TODO: [1] print quantites by queries category
        print("Hardware count: %s" % objQuery.MatchesCount.ToString())
        objSelectSet = asm.SelectSet

        # print(dir(objSelectSet))
        for occurrence in objSelectSet:
            if occurrence.Type == -1879909117:
                occurrence.Visible = False

        # Re-enable selectset UI display.
        selectSet.ResumeDisplay()
        #  Manually refresh the selectset UI display.
        selectSet.RefreshDisplay()

        jt_filename = asm.name[:-4] + ".jt"
        print("JT name : %s" % jt_filename)

        path_download = userprofile() + "\\Downloads" + "\\solidedgeJTs\\"
        if not is_exist(path_download):
            makedirs(path_download)
        new_name = combine(path_download, jt_filename)

        # jt.save_as_jt(asm, new_name)

        # TODO: [1] print volume of the file create
        # TODO: [1] print the timing to create file

    except AssertionError as err:
        print(err.args)

    except Exception as ex:
        print(ex.args)

    finally:
        raw_input("\nPress any key to exit...")
        sys.exit()


def confirmation(func):
    response = raw_input("""Make a .jt format file? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes", "oui"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(generate_jt)
