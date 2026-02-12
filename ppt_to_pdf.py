import sys
import os
import comtypes.client

def convert_ppt_to_pdf(input_folder, output_folder=None):
    """
    Converts all PowerPoint files in the specified folder to PDF.
    
    Args:
        input_folder (str): Path to the folder containing PPT/PPTX files.
        output_folder (str): Path to save PDFs. Defaults to input_folder.
    """
    
    # It is considered best practice to use absolute paths for COM automation
    input_folder = os.path.abspath(input_folder)
    if output_folder:
        output_folder = os.path.abspath(output_folder)
    else:
        output_folder = input_folder

    # Check if input directory exists
    if not os.path.exists(input_folder):
        print(f"Error: The directory {input_folder} does not exist.")
        return

    # Create output directory if it doesn't exist
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except OSError as e:
            print(f"Error creating output directory: {e}")
            return

    # PDF format code for PowerPoint SaveAs
    format_type = 32 

    powerpoint = None
    try:
        # Initialize PowerPoint application
        # It is advisable to keep the app invisible for background processing
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1 

        files = [f for f in os.listdir(input_folder) if f.lower().endswith((".ppt", ".pptx"))]

        if not files:
            print("No PowerPoint files found in the specified directory.")
            return

        print(f"Found {len(files)} files to convert...")

        for filename in files:
            input_path = os.path.join(input_folder, filename)
            
            # Construct output filename
            file_name_without_ext = os.path.splitext(filename)[0]
            output_filename = f"{file_name_without_ext}.pdf"
            output_path = os.path.join(output_folder, output_filename)

            try:
                # Open the presentation
                deck = powerpoint.Presentations.Open(input_path)
                
                # Save as PDF
                deck.SaveAs(output_path, format_type)
                print(f"Successfully converted: {filename}")
                
                # Close the presentation
                deck.Close()
            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

    except Exception as e:
        print(f"An error occurred initializing PowerPoint: {e}")
    
    finally:
        # Ensure PowerPoint quits properly to free resources
        if powerpoint:
            powerpoint.Quit()

if __name__ == "__main__":
    # Configuration
    # Set the target folder to the directory where this script is located
    target_folder = os.path.dirname(os.path.abspath(__file__))
    
    # Output to the same folder (None defaults to input_folder in the function)
    destination_folder = None 
    
    print(f"Scanning for PPT files in: {target_folder}")

    convert_ppt_to_pdf(target_folder, destination_folder)
