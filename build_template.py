import xlsxwriter
import os

def create_ncr_template(filename='NCR_Label_Generator.xlsx'):
    # Create the workbook
    workbook = xlsxwriter.Workbook(filename)

    # --- 1. SETUP INPUT SHEET ---
    ws_input = workbook.add_worksheet('Input')

    # Add Headers
    headers = ['Part #', 'Lot #', 'Serial #', 'NCR #', 'Rejection Reason', 'Inspected By', 'Comments']
    
    header_format = workbook.add_format({
        'bold': True, 
        'bg_color': '#D3D3D3', 
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    for col, text in enumerate(headers):
        ws_input.write(0, col, text, header_format)
    
    ws_input.set_column('A:G', 20)
    
    button_note_format = workbook.add_format({'font_color': 'red', 'bold': True})
    ws_input.write(2, 8, "NOTE: Open Excel and insert Button here linked to the Macro", button_note_format)

    # --- 2. SETUP LABELS SHEET ---
    ws_labels = workbook.add_worksheet('Labels')

    # --- MARGINS (PRECISION TUNED) ---
    # Previous: 11.2mm. Goal: Increase slightly.
    # New Top: 11.6mm = 0.457 inch.
    # Bottom: 13mm = 0.512 inch
    # Left/Right: 4mm = 0.16 inch
    ws_labels.set_margins(
        left=0.157,
        right=0.157,
        top=0.457, 
        bottom=0.512
    )

    # --- COLUMN WIDTHS (LOCKED) ---
    # User confirmed these are good. Kept at 356px.
    
    ws_labels.set_column_pixels('A:A', 356) # Label Left
    ws_labels.set_column_pixels('B:B', 15)  # Gap (4mm)
    ws_labels.set_column_pixels('C:C', 356) # Label Right

    # --- ROW HEIGHTS (MICRO-ADJUSTED) ---
    # Previous: 148.5pts. 
    # Reduced to 148.2pts to compensate for the Top Margin increase.
    # This ensures the bottom label doesn't get pushed to Page 2.
    
    for row in range(100):
        ws_labels.set_row(row, 148.2)

    # --- PRINT SCALING ---
    ws_labels.set_print_scale(100)

    print(f"Successfully created '{filename}' with PRECISION layout (Top 11.6mm, Height 148.2pts).")
    workbook.close()

if __name__ == "__main__":
    create_ncr_template()