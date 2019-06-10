#!/usr/bin/env python3
#
# Copyright LiKneu 2019
#
# Converts Freemind XML files into other file formats like:
#   * Excel
#
import sys  # command line options

from lib.freemind_to_excel import to_excel
from lib.freemind_to_project import to_project

def main():
    script = sys.argv[0]    # filename of this script
    if len(sys.argv)==1:    # no arguments so print help message
        print('''Usage: main.py action filename
        action must be one of --excel --project''')
        return
    
    action = sys.argv[1]
    # check if user input of action is an allowed/defined one
    assert action in ['--excel', '--project'], \
        'Action is not one of --excel --project: ' + action
    
    input_file = sys.argv[2]    # filename of the to be converted mindmap
    output_file = sys.argv[3]   # filename of the export file
    process(action, input_file, output_file)

def process(action, input_file, output_file):
    '''Processes user input.'''
    if action == '--excel':
        to_excel(input_file, output_file)
    elif action == '--project':
        to_project(input_file, output_file)

if __name__ == '__main__':
    main()
