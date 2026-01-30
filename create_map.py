import matplotlib.pyplot as plt
import matplotlib.patches as patches

def create_label_map():
    # 1. Setup the "Paper" (8.5 x 11 inches)
    fig, ax = plt.subplots(figsize=(8.5, 11))
    
    # Remove graph axes ticks/labels, but keep the frame area
    ax.set_xlim(0, 8.5)
    ax.set_ylim(0, 11)
    ax.axis('off')
    
    # Set background to white (Paper color)
    fig.patch.set_facecolor('white')

    # --- NEW: ADD BLACK BORDER AROUND PAPER ---
    # Draw a rectangle the exact size of the paper with a black edge and no fill
    border = patches.Rectangle(
        (0, 0), 8.5, 11, 
        linewidth=4, edgecolor='black', facecolor='none', zorder=10
    )
    ax.add_patch(border)
    # ------------------------------------------
    
    # 2. Define Margins & Label Specs (Matching your Excel Template)
    margin_left = 0.16
    margin_top = 0.5
    label_width = 4.0
    label_height = 2.0
    col_gap = 0.18 
    
    # 3. Draw the 10 Boxes
    row_counter = 2
    
    for r in range(5): 
        y_pos = 11 - margin_top - ((r + 1) * label_height)
        
        # Left Column
        rect_left = patches.Rectangle(
            (margin_left, y_pos), 
            label_width, label_height, 
            linewidth=2, edgecolor='black', facecolor='#f9f9f9'
        )
        ax.add_patch(rect_left)
        
        ax.text(
            margin_left + (label_width/2), 
            y_pos + (label_height/2), 
            f"Row {row_counter}", 
            horizontalalignment='center', verticalalignment='center', 
            fontsize=40, fontweight='bold', fontname='Arial', color='#333333'
        )
        
        # Right Column
        rect_right = patches.Rectangle(
            (margin_left + label_width + col_gap, y_pos), 
            label_width, label_height, 
            linewidth=2, edgecolor='black', facecolor='#f9f9f9'
        )
        ax.add_patch(rect_right)
        
        ax.text(
            margin_left + label_width + col_gap + (label_width/2), 
            y_pos + (label_height/2), 
            f"Row {row_counter + 1}", 
            horizontalalignment='center', verticalalignment='center', 
            fontsize=40, fontweight='bold', fontname='Arial', color='#333333'
        )
        
        row_counter += 2

    # 4. Save
    output_filename = 'label_map_reference.png'
    # Added bbox_inches='tight' and pad_inches=0 to ensure the border isn't cut off
    plt.savefig(output_filename, dpi=300, bbox_inches='tight', pad_inches=0)
    print(f"Successfully generated '{output_filename}' with a border.")

if __name__ == "__main__":
    create_label_map()