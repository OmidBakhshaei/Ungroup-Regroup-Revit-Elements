import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from pyrevit import forms, revit
import ctypes
import sys
import os
from System.Collections.Generic import List

# If the script is run with shift-click, open the script file and exit
if __shiftclick__:
    os.startfile(__file__)
    sys.exit()

uiapp = __revit__
uidoc = uiapp.ActiveUIDocument

# Get the active document and selected elements, if any
if uidoc:
    doc = uiapp.ActiveUIDocument.Document
    selection = [doc.GetElement(x) for x in uidoc.Selection.GetElementIds()]

# Get the current Revit file name
revit_file_name = "_".join(doc.Title.split("_")[:-1])

# Get the Revit application object
app = doc.Application

# Get the user's display name
def get_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3

    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)

    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, nameBuffer, size)
    return nameBuffer.value


user = get_display_name()

selected_elements = selection

# # element.GroupId : returns the id of the group in which the element is a member, returns -1 if the element is not a member of any group
# Create a list of group IDs that have at least one of the selected elements
groups_to_ungroup_ids = list(
    set(
        [
            element.GroupId.IntegerValue
            for element in selected_elements
            if element.GroupId.IntegerValue != -1
        ]
    )
)

# Get the actual elements by their IDs
groups_to_ungroup = [doc.GetElement(ElementId(id)) for id in groups_to_ungroup_ids]

# Filter out invalid group IDs and get unique parent group IDs
parent_groups_to_ungroup_ids = list(
    set(
        [
            group.GroupId.IntegerValue
            for group in groups_to_ungroup
            if group.GroupId.IntegerValue != -1
        ]
    )
)

# Get the actual elements by their IDs
parent_groups_to_ungroup = [
    doc.GetElement(ElementId(id)) for id in parent_groups_to_ungroup_ids
]

# Initialize empty lists for storing group information
group_names = []
group_members_ids = []
parent_group_names = []
parent_group_members_lists_of_ids = []

# Ungroup the groups
with Transaction(doc, "Ungroup") as t:

    t.Start()
    # Iterate over parent groups to ungroup them and save their information
    for group in parent_groups_to_ungroup:
        parent_group_names.append(group.Name)
        parent_group_members_lists_of_ids.append(group.GetMemberIds())
        group.UngroupMembers()

    # Iterate over child groups to ungroup them and save their information
    for group in groups_to_ungroup:
        group_names.append(group.Name)
        group_members_ids.append(group.GetMemberIds())
        group.UngroupMembers()

    t.Commit()

# Convert the list of lists of ElementIds to a list of lists of integers
group_members_ids_int = [
    [id.IntegerValue for id in inner_list] for inner_list in group_members_ids
]
parent_group_members_lists_of_ids_int = [
    [id.IntegerValue for id in inner_list]
    for inner_list in parent_group_members_lists_of_ids
]

# Initialize empty list for storing new group IDs
new_group_ids = []

# Regroup the ugrouped groups
with Transaction(doc, "Regroup") as t:
    t.Start()

    # Iterate over old group members to regroup them
    for index, list_of_member_ids in enumerate(group_members_ids_int):
        new_group = doc.Create.NewGroup(
            List[ElementId]([ElementId(id) for id in list_of_member_ids])
        )
        new_group_ids.append(new_group.Id.IntegerValue)
        new_group.GroupType.Name = group_names[index] + " AUTOMATICALLY"

    t.Commit()


# Iterate over parent group members to reassign new group IDs to them
for index, list_of_member_ids in enumerate(parent_group_members_lists_of_ids_int):
    for i, id in enumerate(list_of_member_ids):
        if id in groups_to_ungroup_ids:
            parent_group_members_lists_of_ids_int[index][i] = new_group_ids[
                groups_to_ungroup_ids.index(id)
            ]


with Transaction(doc, "Regroup the parents") as t:
    t.Start()

    # Iterate over parent groups to create new groups and regroup them
    for index, list_of_member_ids in enumerate(parent_group_members_lists_of_ids_int):
        new_group = doc.Create.NewGroup(
            List[ElementId]([ElementId(id) for id in list_of_member_ids])
        )
        new_group.GroupType.Name = parent_group_names[index] + " AUTOMATICALLY"

    t.Commit()