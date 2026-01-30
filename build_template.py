import xlsxwriter
import os

def create_ncr_template(filename='NCR_Label_Generator.xlsx'):
    # Create the workbook
    workbook = xlsxwriter.Workbook(filename)

    # --- 1. SETUP INPUT SHEET ---
    ws_input = workbook.add_worksheet('Input')

    # Define Formats
    header_format = workbook.add_format({
        'bold': True, 
        'bg_color': '#D3D3D3', 
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Standard Left Align
    left_format = workbook.add_format({
        'align': 'left',
        'valign': 'top' 
    })

    # Wrap Text + Left Align
    wrap_format = workbook.add_format({
        'text_wrap': True,
        'align': 'left',
        'valign': 'top'
    })

    # Add Headers
    headers = ['Part #', 'Lot #', 'Serial #', 'NCR #', 'Disposition', 'Rejection Reason', 'Inspected By', 'Comments']
    for col, text in enumerate(headers):
        ws_input.write(0, col, text, header_format)
    
    # --- COLUMN WIDTHS & ALIGNMENT ---
    
    # Cols A-E: Part, Lot, Serial, NCR, Disposition -> Width 15
    ws_input.set_column('A:E', 15, left_format)
    
    # Col F: Rejection Reason -> Width 30
    ws_input.set_column('F:F', 30, wrap_format)
    
    # Col G: Inspected By -> Width 20
    ws_input.set_column('G:G', 20, left_format)
    
    # Col H: Comments -> Width 55
    ws_input.set_column('H:H', 55, wrap_format)

    # --- DROPDOWN LIST (DATA VALIDATION) ---
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
    # Total Label Height must equal 148.2 points to match physical paper.
    # We redistribute this height to give "Comments" more room.
    
    for label_i in range(10): # Loop through 10 labels
        base = label_i * 5
        ws_labels.set_row(base, 29.64)     # Row 1: Part # / Lot
        ws_labels.set_row(base + 1, 29.64) # Row 2: Serial # / NCR
        ws_labels.set_row(base + 2, 20.0)  # Row 3: Insp / Disposition (Small)
        ws_labels.set_row(base + 3, 20.0)  # Row 4: Reason (Small)
        ws_labels.set_row(base + 4, 48.92) # Row 5: Comments (Maximized)

    # --- PRINT SCALING ---
    ws_labels.set_print_scale(100)

    print(f"Successfully created '{filename}' with Variable Row Heights (Max Comments).")
    workbook.close()

if __name__ == "__main__":
    create_ncr_template()