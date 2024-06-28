import os
import comtypes.client
import streamlit as st
from PIL import Image
import zipfile
import base64
import datetime

# Initialize COM
comtypes.CoInitialize()

def convert_pptx_to_images(pptx_path, output_dir):
    try:
        # Initialize PowerPoint application
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1

        # Open the PowerPoint presentation
        presentation = powerpoint.Presentations.Open(pptx_path)

        # Ensure output directory exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Iterate through each slide in the presentation
        image_paths = []
        for i, slide in enumerate(presentation.Slides):
            # Save each slide as a PNG image
            output_path = os.path.join(output_dir, f"Slide_{i + 1}.png")
            # Convert to absolute path to avoid issues
            output_path = os.path.abspath(output_path)
            slide.Export(output_path, "PNG")
            image_paths.append(output_path)

        # Close the presentation
        presentation.Close()

        # Quit the PowerPoint application
        powerpoint.Quit()

        return image_paths

    except Exception as e:
        st.error(f"Error during PowerPoint conversion: {e}")
        return []

    finally:
        # Ensure COM is uninitialized
        comtypes.CoUninitialize()

# Function to generate HTML for downloading files
def get_binary_file_downloader_html(bin_file, file_label='File', button_text='Download'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">{button_text}</a>'
    return href

# Streamlit app
st.title("PowerPoint to Image Converter")

uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded_file is not None:
    # Ensure the temp directory exists
    temp_dir = os.path.abspath("temp")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    # Save the uploaded file temporarily
    temp_pptx_path = os.path.join(temp_dir, uploaded_file.name)
    with open(temp_pptx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Ensure the temp file is accessible
    if not os.path.exists(temp_pptx_path):
        st.error(f"The file '{temp_pptx_path}' could not be saved.")
    else:
        try:
            # Extract base filename without extension
            base_filename = os.path.splitext(uploaded_file.name)[0]
            # Create folder name with timestamp for uniqueness
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            folder_name = f"{base_filename}_{timestamp}"

            # Specify the output directory based on the folder name
            output_dir = os.path.abspath(folder_name)

            consent_button = st.button("Click to consent and download")
            if consent_button:
                image_paths = convert_pptx_to_images(temp_pptx_path, output_dir)

            if image_paths:
                # Display the images
                st.header("Converted Slides")
                for image_path in image_paths:
                    st.image(Image.open(image_path))


        except Exception as e:
            st.error(f"An error occurred: {e}")

        finally:
            # Remove the temporary PowerPoint file
            os.remove(temp_pptx_path)

# Run Streamlit app
if __name__ == "__main__":
    st.set_page_config(layout="wide")
    st.write("""
        # PowerPoint to Image Converter
        Upload your PowerPoint file and get each slide as an image.
    """)
    st.write("""
        ## Instructions
        - Click the **Browse files** button to upload a `.pptx` file.
        - Wait for the file to be processed and the images to be displayed.
        - Click **Click to consent and download** to download the converted images as a ZIP file.
    """)
