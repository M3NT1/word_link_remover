# file_utils.py

import os

def get_file_path(directory, filename):
    return os.path.join(directory, filename)

def create_output_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)