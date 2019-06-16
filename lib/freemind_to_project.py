#!/usr/bin/env python3
#
# freemind_to_project.py
#
# Copyright LiKneu 2019
#

import pprint   # Similar to Data::Dumper
import lxml.etree as ET # handling XML

def main(args):
    return 0

def get_pj_path(fm_path):
    '''Takes the XPATH of the Freemind parent element and returns the XPATH
    of the Freemind element'''
    pj_path = fm_path.replace('map', 'Project')
    pj_path = pj_path.replace('node', 'Tasks', 1)
    pj_path = pj_path.replace('node', 'Task')
    return pj_path

def node_to_note(pj_root, pj_uid, pj_text):
    '''Stores the note information of Freemind into a <Note> tag of Project'''
    # Create the XPATH search pattern to find the Project UID...
    xpath_pattern = '//UID[text()="' + str(pj_uid) + '"]'
    task_id = pj_root.xpath(xpath_pattern)
    # ...to get the parent node of this UID...
    pj_parent = task_id[0].getparent()
    # ...to attach the <Notes> to it.
    note = ET.SubElement(pj_parent, "Notes")
    note.text = pj_text

def node_to_task(pj_parent, pj_name, pj_uid, pj_level=1):
    '''Creates <Task> element with the necessary subelements and attaches it to
    a given parent element.
    
    # Example of a minimalistic task as it was produced from a Freemap
    # Project 2003 export
    <Task>
    <UID>0</UID>
    <ID>1</ID>
    <Type>1</Type>
    <IsNull>0</IsNull>
    <OutlineNumber>1</OutlineNumber>
    <OutlineLevel>0</OutlineLevel>
    <Name>Stamm</Name>
    <FixedCostAccrual>1</FixedCostAccrual>
    <RemainingDuration>PT8H0M0S</RemainingDuration>
    <Estimated>1</Estimated>
    <PercentComplete>0</PercentComplete>
    <Priority/>
    </Task>
    '''

    task = ET.SubElement(pj_parent, "Task")
    
    # Name (occurence: min = 0, max = 1)
    #   The name of the task.
    name = ET.SubElement(task, "Name")
    # Project doesn't allow linebreaks in task titles so we remove them here
    name.text = pj_name.replace('\n', ' ')
    
    # UID (occurences: min = max = 1)
    #   The UID element is a unique identifier
    uid = ET.SubElement(task, "UID")
    uid.text = str(pj_uid)
    
    # ID  (occurence: min = 0, max = 1)
    #   For Resource, ID is the position identifier of the resource within the
    #   list of resources. For Task, it is the position identifier of the task
    #   in the list of tasks.
    pj_id = ET.SubElement(task, "ID")
    pj_id.text = "1"

    # Type
    #   0 = fixed units
    #   1 = fixed duration
    #   2 = fixed work
    pj_type = ET.SubElement(task, "Type")
    pj_type.text = "1"

    # IsNull
    #   0 = task or ressource is not null
    #   1 = task or ressource is null
    isnull = ET.SubElement(task, "IsNull")
    isnull.text = "0"
    
    # OutlineNumber
    #   Indicates the exact position of a task in the outline. For example,
    #   "7.2" indicates that a task is the second subtask under the seventh
    #   top-level summary task.
    outlinenumber = ET.SubElement(task, "OutlineNumber")
    outlinenumber.text = "1"

    # OutlineLevel
    #   The number that indicates the level of a task in the project outline
    #   hierarchy.
    outlinelevel = ET.SubElement(task, "OutlineLevel")
    outlinelevel.text = str(pj_level)

    # FixedCostAccrual
    #   Indicates how fixed costs are to be charged, or accrued, to the cost
    #   of a task. (no info on Microsofts webpage what values are allowed)
    fixedcostaccrual = ET.SubElement(task, "FiexedCostAccrual")
    fixedcostaccrual.text = "1"
    
    # RemainingDuration (occurence: min = 0, max = 1)
    #   The amount of time required to complete the unfinished portion of a
    #   task. Remaining duration can be calculated in two ways, either based on
    #   percent complete or on actual duration.
    # TODO: check if RemainingDuration is necessary
    remainingduration = ET.SubElement(task, "RemainingDuration")
    remainingduration.text = "PT8H0M0S" # 8 Hours, 0 Minutes, 0 seconds 
    
    # Estimated (occurence: min = 0, max = 1)
    #   Indicates whether the task's duration is flagged as an estimate.
    #   0 = not estimated (i.e. precise)
    #   1 = estimated
    estimated = ET.SubElement(task, "Estimated")
    estimated.text = "1"
    
    # PercentComplete (occurence: min = 0, max = 1)
    #   The percentage of the task duration completed.
    percentcomplete = ET.SubElement(task, "PercentComplete")
    percentcomplete.text = "0"
    
    # Priority (occurence: min = 0, max = 1)
    #   Indicates the level of importance assigned to a task, with 500 being
    #   standard priority; the higher the number, the higher the priority.
    priority = ET.SubElement(task, "Priority")
    
def to_project(input_file, output_file):
    '''Converts the Freeemind XML file into an MS Project XML file.'''

    print('Converting to Project')
    print('Input file : ' + input_file)
    print('Output file: ' + output_file)
    
    # Prepare the Fremmind XML tree
    # Read freemind XML file into variable
    fm_tree = ET.parse(input_file)
    # Get the root element of the Freemind XML tree
    fm_root = fm_tree.getroot()
    
    # Prepare the root element of the new Project XML file
    attrib = {'xmlns':'http://schemas.microsoft.com/project'}
    pj_root = ET.Element('Project', attrib)
    
    # Based on the Project root element we define the Project tree
    pj_tree = ET.ElementTree(pj_root)
    
    # Add the Project <Title> element
    pj_title = ET.SubElement(pj_root, 'Title')
    pj_title.text = fm_root.xpath('/map/node/@TEXT')[0]
    
    # Add the Project <Tasks> element
    pj_tasks = ET.SubElement(pj_root, 'Tasks')
    pj_path = pj_tree.getpath(pj_tasks)
    
    # Dict holding mapping table of Freemind and Project UIDs
    uid_mapping = {}
    
    # UID in Project starts with 0
    pj_uid = 0
    
    for fm_node in fm_root.iter('node'):
        # Determine the parent of the present Freemind element
        fm_parent = fm_node.getparent()
        # Determine the XPATH of the Freemind parent element
        fm_parent_path = fm_tree.getpath(fm_parent)
        # Determine the XPATH of the Project parent element from the XPATH of
        # the Freemind parent element
        pj_parent_path = get_pj_path(fm_parent_path)
        print("pj parent path:", pj_parent_path)
        # Determine the Project parent element based on its XPATH
        pj_parent = pj_tree.xpath(pj_parent_path)[0]
        # Get the Project text from the Freemind node TEXT attribute
        pj_name = fm_node.get("TEXT")
        # Get the Freemind ID from its attribute
        fm_id = fm_node.get("ID")
        # Add Freemind ID and Project UID to mapping table (Dictionary)
        uid_mapping[fm_id] = pj_uid
        # calculate level with help of XPATH
        # Count number of dashed as indicator for the depth of the structure
        # -1 is for the 1st that is not needed here
        pj_level = pj_tree.getpath(pj_parent).count('/')-1
        node_to_task(pj_parent=pj_parent, pj_name=pj_name, pj_uid=pj_uid, pj_level=pj_level)
        
        # Check if node has an attached note <richcontent>
        fm_note = fm_node.xpath('normalize-space(./richcontent)')
        # If yes, remove all html tags and store the remaining text in a
        # Project <Note> tag
        if fm_note:
            node_to_note(pj_root=pj_root, pj_uid=pj_uid, pj_text=fm_note)
        
        pj_uid += 1

    # Write the Project XML tree to disc
    pj_tree.write(output_file, pretty_print=True, xml_declaration=True, encoding="utf-8")


if __name__ == '__main__':
    import sys
    sys.exit(main(sys.argv))
