import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import *
from datetime import datetime
from pyrevit import forms, revit
import ctypes
import sys
import os
from System.Collections.Generic import List
import csv
import Core

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

if not selection:
    with forms.WarningBar(title='Select the model groups which you would like to ungroup:'):
        groups_category = BuiltInCategory.OST_IOSModelGroups
        selection = revit.pick_elements_by_category(groups_category)

if not selection:
    forms.alert('No groups were selected!', title='Group Selection', ok =True)
    sys.exit()

# Get the current Revit file name
revit_file_name = "_".join(doc.Title.split("_")[:-1])
revit_file_full_name = doc.Title

# Get the Revit application object
app = doc.Application

# Get the current date
current_date = datetime.now().date()

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

if "DETACHED" in revit_file_full_name.upper() and "OMID" not in user.upper():
    forms.alert('This tool is not applicable to detached files!', title='Ungroup', ok =True)
    sys.exit()

Core.log("Ungroup", user, revit_file_name)
Core.log_overall()

# Check if the selection is a list, and if it's not, convert it to a list
if not isinstance(selection, list):
    groups_to_ungroup = [selection]

# Iterate over the elements in the selection to check if they are all "Model Groups"
for element in selection:
    try:
        if element.Category.Name != "Model Groups":
            forms.alert('You have selected one of the {}!\n\nPlease make sure to select a "Model Group" and run the tool again.'.format(element.Category.Name), title='Group Selection', ok =True)
            sys.exit()
    except:
        pass

group_collector = FilteredElementCollector(doc).OfClass(Group)
groups_names = [group.Name for group in group_collector]
one_has_parent = False
has_dimension = False

# Iterate over the elements in the selection to get the actual elements by their IDs
for selection in selection: 
    groups_to_ungroup = selection
    # Check if the groups to ungroup is a list or not, and get the group ID(s)
    if type(groups_to_ungroup) is list:
        for group in groups_to_ungroup:
            count = groups_names.count(group.Name)
            if count > 1:
                for one_group in group_collector:
                    if one_group.Name == group.Name:
                        if one_group.GroupId.IntegerValue != -1:
                            one_has_parent = True
                            break
                        else:
                            dependent_elements = one_group.GetDependentElements(None)
                            for dependent_element in dependent_elements:
                                element = doc.GetElement(dependent_element)
                                if isinstance(element, Dimension):
                                    has_dimension = True

        group_id = selection[0].Id.IntegerValue
        # Filter out invalid group IDs and get unique parent group IDs
        parent_groups_to_ungroup_ids = list(set([group.GroupId.IntegerValue for group in groups_to_ungroup if group.GroupId.IntegerValue != -1]))
    else:
        group_id = selection.Id.IntegerValue
        count = groups_names.count(selection.Name)
        if count > 1:
            for one_group in group_collector:
                if one_group.Name == group.Name:
                    if one_group.GroupId.IntegerValue != -1:
                        one_has_parent = True
                        break
                    else:
                        dependent_elements = one_group.GetDependentElements(None)
                        for dependent_element in dependent_elements:
                            element = doc.GetElement(dependent_element)
                            if isinstance(element, Dimension):
                                has_dimension = True

        # Filter out invalid group IDs and get unique parent group IDs
        parent_groups_to_ungroup_ids = [groups_to_ungroup.GroupId.IntegerValue] if groups_to_ungroup.GroupId.IntegerValue != -1 else []

    if one_has_parent or parent_groups_to_ungroup_ids:
        forms.alert("This tool is designed to be used with groups that contain multiple instances, with none of them being hosted within other groups.")
        sys.exit()
    if count > 1 and has_dimension:
        forms.alert("Please be cautious as some of the elements in other instances of this group have dimensions tied to them.\n\nMoving the elements within this group could potentially lead to the loss of these dimensions during the regrouping process.")
    elif count > 1:
        forms.alert("Please note that this group has multiple instances within the project.\n\nIt is crucial to ba aware that any modifications made to the elements will affect all other instances of the group during the regrouping process.")
    # Get the actual parent groups to ungroup by their IDs
    parent_groups_to_ungroup = [doc.GetElement(ElementId(id)) for id in parent_groups_to_ungroup_ids]

    # Initialize empty lists for storing group information
    group_names = []
    group_members_ids = []
    parent_group_names = []
    parent_group_members_lists_of_ids = []
    group_type_ids = []
    group_centerpoints = []

    # Ungroup the groups
    with Transaction(doc, 'Ungroup') as t:

        t.Start()

        # Iterate over parent groups to ungroup them and save their information
        for group in parent_groups_to_ungroup:
            parent_group_names.append(group.Name)
            parent_group_members_lists_of_ids.append(group.GetMemberIds())
            group.UngroupMembers()

        # Iterate over child groups to ungroup them and save their information
        if type(groups_to_ungroup) is list:
            for group in groups_to_ungroup:
                group_names.append(group.Name)
                group_type_ids.append(group.GroupType.Id.IntegerValue)
                group_centerpoint = group.Location.Point
                group_centerpoints.append([group_centerpoint.X, group_centerpoint.Y, group_centerpoint.Z])
                group_members_ids.append(group.GetMemberIds())
                group.UngroupMembers()
        else:
            group_names.append(groups_to_ungroup.Name)
            group_type_ids.append(group.GroupType.Id.IntegerValue)
            group_centerpoint = group.Location.Point
            group_centerpoints.append([group_centerpoint.X, group_centerpoint.Y, group_centerpoint.Z])
            group_members_ids.append(groups_to_ungroup.GetMemberIds())
            groups_to_ungroup.UngroupMembers()

        t.Commit()

    # Convert the list of lists of ElementIds to a list of lists of integers
    group_members_ids_int = [[id.IntegerValue for id in inner_list] for inner_list in group_members_ids][0]

    # Convert the list of parent group member IDs to a list of integers, and get the parent group names as strings
    parent_group_names_str = []
    parent_group_members_lists_of_ids_int = []
    if parent_group_members_lists_of_ids:
        parent_group_members_lists_of_ids_int = [[id.IntegerValue for id in inner_list if id.IntegerValue != group_id] for inner_list in parent_group_members_lists_of_ids][0]
        for group_name in parent_group_names:
            parent_group_names_str.append(str(group_name))

    # Store the group information
    Core.write_folder(
        "Group-{}-{}-{}".format(group_names[0], user, revit_file_name),
        "Groups",
        str(group_names),
        str(group_members_ids_int),
        str(parent_group_names_str),
        str(parent_group_members_lists_of_ids_int),
        str(group_type_ids),
        str(group_centerpoints)
    )

# --------------------------------------------------------------------------------
Core.log_done(user)
