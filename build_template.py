import xlsxwriter
import os

def create_ncr_template(filename='NCR_Label_Generator.xlsx'):
    # Create the workbook
    workbook = xlsxwriter.Workbook(filename)

    # --- 1. SETUP INPUT SHEET ---
    ws_input = workbook.add_worksheet('Input')

    # --- DEFINE FORMATS ---
    # 1. Header Format (LOCKED)
    header_format = workbook.add_format({
        'bold': True, 
        'bg_color': '#D3D3D3', 
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'locked': True  # Headers cannot be edited
    })

    # 2. Input Cell Format (UNLOCKED)
    # We explicitly unlock these so the user can type here while the rest of the sheet is protected
    unlocked_left = workbook.add_format({
        'align': 'left',
        'valign': 'top',
        'locked': False 
    })

    unlocked_wrap = workbook.add_format({
        'text_wrap': True,
        'align': 'left',
        'valign': 'top',
        'locked': False
    })

    # Add Headers
    headers = ['Part #', 'Lot #', 'Serial #', 'NCR #', 'Disposition', 'Rejection Reason', 'Inspected By', 'Comments']
    for col, text in enumerate(headers):
        ws_input.write(0, col, text, header_format)
    
    # --- COLUMN WIDTHS & ALIGNMENT ---
    # Apply the UNLOCKED formats to the data columns
    ws_input.set_column('A:E', 15, unlocked_left)
    ws_input.set_column('F:F', 30, unlocked_wrap)
    ws_input.set_column('G:G', 20, unlocked_left)
    ws_input.set_column('H:H', 55, unlocked_wrap)

    # --- DROPDOWN LIST ---
    ws_input.data_validation('E2:E1000', {
        'validate': 'list',
        'source': ['RTV', 'Rework', 'Use as is', 'Sort', 'Scrap'],
        'input_title': 'Select Disposition',
        'input_message': 'Choose from the list'
    })

    # --- 2. SETUP LABELS SHEET ---
    ws_labels = workbook.add_worksheet('Labels')

    # --- MARGINS ---
    ws_labels.set_margins(
        left=0.157,
        right=0.157,
        top=0.457, 
        bottom=0.512
    )

    # --- COLUMN WIDTHS (SPLIT GRID) ---
    ws_labels.set_column_pixels('A:B', 178) # Label 1
    ws_labels.set_column_pixels('C:C', 15)  # Gap
    ws_labels.set_column_pixels('D:E', 178) # Label 2

    # --- ROW HEIGHTS (VARIABLE DISTRIBUTION) ---
    for label_i in range(10): 
        base = label_i * 5
        ws_labels.set_row(base, 29.64)     # Row 1
        ws_labels.set_row(base + 1, 29.64) # Row 2
        ws_labels.set_row(base + 2, 20.0)  # Row 3 (Small)
        ws_labels.set_row(base + 3, 20.0)  # Row 4 (Small)
        ws_labels.set_row(base + 4, 48.92) # Row 5 (Large)

    # --- PRINT SCALING ---
    ws_labels.set_print_scale(100)

    # --- INSERT REFERENCE IMAGE ---
    img_path = 'label_map_reference.png'
    
    if os.path.exists(img_path):
        ws_input.insert_image('I11', img_path, {
            'x_scale': 0.305, 
            'y_scale': 0.305,
            'x_offset': 0, 
            'y_offset': 0
        })
        print(f"Successfully created '{filename}' with Locked Map Image.")
    else:
        print(f"Successfully created '{filename}' (Map image not found, skipping).")

    # --- ENABLE PROTECTION ---
    # This prevents users from selecting/deleting the visual map or notes
    # They can ONLY edit the cells we marked as 'locked': False
    ws_input.protect('', {
        'select_locked_cells': True,
        'select_unlocked_cells': True,
        'objects': False,    # User cannot move/delete images
        'scenarios': False
    })

    workbook.close()

if __name__ == "__main__":
    create_ncr_template()