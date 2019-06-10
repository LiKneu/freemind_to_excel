#!/usr/bin/env python3

def to_excel(input_file, output_file):
    '''Converts the Freeemind XML file into an Excel sheet.'''

    import lxml.etree as ET # handling XML
    import openpyxl         # handling Excel files
    from tqdm import tqdm   # progress bar

    print('Converting to Excel')
    print('Input file : ' + input_file)
    print('Output file: ' + output_file)
    
    wb = openpyxl.Workbook()    # Create Excel workbook object
    
    # Get the active sheet (Excel creates sheet 'Sheet') automatically with the
    # creation of the workbook
    sheet = wb.active
    
    tree = ET.parse(input_file) # Read freemind XML file into variable

    root = tree.getroot()       # Get the root element of the XML file

    # Get text of the root node with a XPATH ...
    sheet_title = root.xpath('/map/node/@TEXT')
    sheet.title = sheet_title[0]   # ..and name the Excel sheet like it
    
    base = []   # Create list to hold the cells for export to EXCEL
    headings = ['Level 0']
    max_levels = 0

    for node in tqdm(root.iter('node')):
        path = tree.getpath(node)   # Get the xpath of the node
        
        # Count number of dashed as indicator for the depth of the structure
        # -2 is for the 1st 2 levels that are not needed here
        nr_dashes = path.count('/') - 2
        
        if nr_dashes > max_levels:  # To generate headings for each column we
            max_levels = nr_dashes  # need to know the max depth of the tree
            # Add a heading to the list for each level of the tree
            headings.append('Level ' + str(max_levels))
        
        # Count numbers of elements already in the list
        nr_base = len(base)
        
        # If we have less dashes than elements we now that we jumped back a level
        if nr_dashes < nr_base:
            # And so we reduce the list to the same number like levels
            base = base[:nr_dashes]

        # Append the text of the element to the list
        base.append(node.get('TEXT'))
            
        # If there are no children of type 'node' below the present node add
        # data to Excel sheet
        if not node.xpath('.//node'):
            sheet.append(base)

    # Insert an empty row on top of all rows to make room for header cells
    sheet.insert_rows(1)
    
    col = 1 # First row in Excel is 1 (not 0)
    # Since it seems not possible to inser a whole row with data at the top of
    # a shhet we have to iterate through the columns and write cells separately
    for heading in headings:
        sheet.cell(row=1, column=col).value = heading
        col+=1
    
    # Add autofilter
    # We use sheet.dimensions to let openpyxl determin the number of rows & columns
    sheet.auto_filter.ref = sheet.dimensions
    
    sheet.freeze_panes = 'A2'
    
    wb.save(output_file)    # Save the workbook we have created above to disc
