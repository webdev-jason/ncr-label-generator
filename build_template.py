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

    # Standard Left Align for normal fields (Part, Lot, etc.)
    left_format = workbook.add_format({
        'align': 'left',
        'valign': 'top' # Aligns text to top so it matches long comments
    })

    # Wrap Text + Left Align for long fields (Rejection, Comments)
    wrap_format = workbook.add_format({
        'text_wrap': True,
        'align': 'left',
        'valign': 'top'
    })

    # Add Headers
    headers = ['Part #', 'Lot #', 'Serial #', 'NCR #', 'Rejection Reason', 'Inspected By', 'Comments']
    for col, text in enumerate(headers):
        ws_input.write(0, col, text, header_format)
    
    # --- COLUMN WIDTHS & ALIGNMENT ---
    
    # A, B, C, D (Part, Lot, Serial, NCR) -> Width 15 + Left Align
    ws_input.set_column('A:D', 15, left_format)
    
    # E (Rejection Reason) -> Width 35 (Smaller) + Wrap Text + Left Align
    ws_input.set_column('E:E', 35, wrap_format)
    
    # F (Inspected By) -> Width 25 (Bigger) + Left Align
    ws_input.set_column('F:F', 25, left_format)
    
    # G (Comments) -> Width 60 + Wrap Text + Left Align
    ws_input.set_column('G:G', 60, wrap_format)

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

    # --- ROW HEIGHTS (5 ROWS PER LABEL) ---
    for row in range(50):
        ws_labels.set_row(row, 29.64)

    # --- PRINT SCALING ---
    ws_labels.set_print_scale(100)

    print(f"Successfully created '{filename}' with Adjusted Columns (Rejection 35, Inspected 25).")
    workbook.close()

if __name__ == "__main__":
    create_ncr_template()