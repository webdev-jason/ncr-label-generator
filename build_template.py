import xlsxwriter
import os

def create_ncr_template(filename='NCR_Label_Generator.xlsx'):
    # Create the workbook
    workbook = xlsxwriter.Workbook(filename)

    # --- 1. SETUP INPUT SHEET ---
    ws_input = workbook.add_worksheet('Input')

    # Add Headers
    headers = ['Part #', 'Lot #', 'Serial #', 'NCR #', 'Rejection Reason', 'Inspected By', 'Comments']
    
    # Format: Light Grey background, Bold text, Border
    header_format = workbook.add_format({
        'bold': True, 
        'bg_color': '#D3D3D3', 
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Write headers to Row 1 (Index 0)
    for col, text in enumerate(headers):
        ws_input.write(0, col, text, header_format)
    
    # Set column widths for Input sheet (Columns A-G)
    ws_input.set_column('A:G', 20)
    
    # Add a visual placeholder for the button
    button_note_format = workbook.add_format({'font_color': 'red', 'bold': True})
    ws_input.write(2, 8, "NOTE: Open Excel and insert Button here linked to the Macro", button_note_format)

    # --- 2. SETUP LABELS SHEET ---
    ws_labels = workbook.add_worksheet('Labels')

    # Set Page Margins for Quill 7-32205 (0.5" top/bottom, 0.16" sides)
    ws_labels.set_margins(left=0.16, right=0.16, top=0.5, bottom=0.5)

    # Set Column Widths 
    # 4.1 inches is approx 46 chars in Excel standard font
    # 0.15 inches is approx 1.5 chars
    ws_labels.set_column('A:A', 46) 
    ws_labels.set_column('B:B', 1.5)
    ws_labels.set_column('C:C', 46)

    # Set Row Height for 100 rows (2 inches = 144 points)
    for row in range(100):
        ws_labels.set_row(row, 144)

    print(f"Successfully created '{filename}'")
    print("Next Steps:")
    print("1. Open the Excel file.")
    print("2. Save it as a Macro-Enabled Workbook (.xlsm).")
    print("3. Press ALT+F11 and paste the code from 'label_logic.vba'.")

    workbook.close()

if __name__ == "__main__":
    create_ncr_template()