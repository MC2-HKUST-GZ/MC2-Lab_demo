import os
from PIL import Image

# Specify the source and destination directories
source_directory = r'C:\Users\wenjy\Desktop\MC2-Lab-main\images\Publication'
destination_directory = r'C:\Users\wenjy\Desktop\MC2-Lab-main\images\Publication_png'

# Create the destination directory if it doesn't exist
if not os.path.exists(destination_directory):
    os.makedirs(destination_directory)

# Loop through all files in the source directory
for filename in os.listdir(source_directory):
    # Check if the file is an image (you can add more formats if needed)
    if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')):
        # Construct full file path
        file_path = os.path.join(source_directory, filename)
        
        # Open an image file
        try:
            with Image.open(file_path) as img:
                # Convert image to PNG format
                png_filename = f"{os.path.splitext(filename)[0]}.png"  # Change the extension to .png
                png_file_path = os.path.join(destination_directory, png_filename)
                
                # Save the image as PNG
                img.save(png_file_path, 'PNG')
                print(f"Converted {filename} to {png_filename}")
        except Exception as e:
            print(f"Could not convert {filename}: {e}")

print("All images have been converted to PNG format and saved to the new location.")