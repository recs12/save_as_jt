import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System

def save_as(document, NewName):
    SaveAs( document,
    NewName,
    # Optional ByVal IsATemplate As Variant, _
    # Optional ByVal FileFormat As Variant, _
    # Optional ByVal ReadOnlyEnforced As Variant, _
    # Optional ByVal ReadOnlyRecommended As Variant, _
    # Optional ByVal NewStatus As Variant, _
    # Optional ByVal CreateBackup As Variant, _
    # Optional ByVal UpdateLinkInContainer As Variant, _
    # Optional ByVal UpdateAllLinksInContainer As Variant _
    ) 