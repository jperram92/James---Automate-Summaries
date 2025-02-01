import argparse
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import logging
import sys
import os

# Set up logging
logging.basicConfig(filename='excel_to_ppt.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def validate_input(input_file, template_file):
    """Validate input files and structure"""
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Excel file {input_file} not found")
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"Template file {template_file} not found")

    required_columns = ['Section', 'Title', 'Description', 'Priority']
    df = pd.read_excel(input_file)
    missing_cols = [col for col in required_columns if col not in df.columns]
    
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
    
    return df

def add_priority_badge(slide, priority):
    """Add colored priority badge to slide"""
    priority_colors = {
        'High': RGBColor(198, 31, 60),    # Red
        'Medium': RGBColor(255, 192, 0),  # Amber
        'Low': RGBColor(59, 168, 85)      # Green
    }
    
    left = Inches(8.5)
    top = Inches(0.2)
    width = Inches(1)
    height = Inches(0.4)
    
    shape = slide.shapes.add_shape(
        1,  # Rectangle
        left, top, width, height
    )
    
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = priority_colors.get(priority, RGBColor(128, 128, 128))
    
    text_frame = shape.text_frame
    p = text_frame.paragraphs[0]
    p.text = priority
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = 1  # Center alignment

def create_ppt(input_file, template_file, output_file):
    """Main function to create PowerPoint from Excel"""
    try:
        df = validate_input(input_file, template_file)
        prs = Presentation(template_file)
        
        # Use the first slide layout for all new slides
        slide_layout = prs.slide_layouts[1]
        
        for _, row in df.iterrows():
            slide = prs.slides.add_slide(slide_layout)
            
            # Set title
            title = slide.shapes.title
            title.text = f"{row['Section']} - {row['Title']}"
            
            # Set content
            content = slide.placeholders[1]
            content.text = row['Description']
            
            # Add diagram if specified
            if 'Diagram Needed' in row and pd.notna(row['Diagram Needed']):
                img_path = f"diagrams/{row['Diagram Needed']}.png"
                if os.path.exists(img_path):
                    left = Inches(1)
                    top = Inches(2)
                    height = Inches(4)
                    slide.shapes.add_picture(img_path, left, top, height=height)
            
            # Add priority badge
            add_priority_badge(slide, row['Priority'])
        
        prs.save(output_file)
        print(f"Successfully created {output_file} with {len(df)} slides")
        
    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        sys.exit(f"Error: {str(e)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert Excel to PowerPoint')
    parser.add_argument('--input', required=True, help='Input Excel file')
    parser.add_argument('--template', required=True, help='PowerPoint template file')
    parser.add_argument('--output', default='output.pptx', help='Output PowerPoint file')
    
    args = parser.parse_args()
    
    create_ppt(
        input_file=args.input,
        template_file=args.template,
        output_file=args.output
    )