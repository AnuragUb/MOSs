import os
import shutil

# Source and destination directories
source_dir = os.path.dirname(os.path.abspath(__file__))
dest_dir = os.path.join(source_dir, 'MOS Offline')

# Create destination directories if they don't exist
for path in ['static/css', 'static/js', 'templates', 'uploads']:
    full_path = os.path.join(dest_dir, path)
    os.makedirs(full_path, exist_ok=True)

# List of static files to copy (CSS)
css_files = [
    ('static/css/style.css', 'static/css/style.css'),
    ('static/css/topbar.css', 'static/css/topbar.css'),
]

# Copy CSS files
for src, dst in css_files:
    src_path = os.path.join(source_dir, src)
    dst_path = os.path.join(dest_dir, dst)
    
    if os.path.exists(src_path):
        shutil.copy2(src_path, dst_path)
        print(f"Copied: {src} -> {dst}")
    else:
        print(f"Warning: Source file {src} does not exist")

print("File copying complete!") 