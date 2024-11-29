import os
import fnmatch

def delete_localized_resx_files(directory, base_file_name):
    """
    Deletes all .resx files in the specified directory that start with the base_file_name
    but excludes the language-neutral one (e.g., AboutBox.resx).
    
    Args:
        directory (str): The path to the directory containing the .resx files.
        base_file_name (str): The base file name to match (e.g., "AboutBox").
    """
    if not os.path.isdir(directory):
        print(f"Error: The specified directory '{directory}' does not exist.")
        return
    
    # Pattern to match localized .resx files (e.g., AboutBox.ar.resx, AboutBox.fr.resx, etc.)
    localized_pattern = f"{base_file_name}.*.resx"
    language_neutral_file = f"{base_file_name}.resx"
    
    # Iterate through the files in the directory
    for file_name in os.listdir(directory):
        file_path = os.path.join(directory, file_name)
        
        # Check if the file matches the localized pattern but is not the language-neutral file
        if fnmatch.fnmatch(file_name, localized_pattern) and file_name != language_neutral_file:
            try:
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            except OSError as e:
                print(f"Error deleting file {file_path}: {e}")
    
    print("Localized .resx file deletion completed.")

# Example usage
directory_path = "../"
delete_localized_resx_files(directory_path, "AboutBox")
delete_localized_resx_files(directory_path, "Forge")
delete_localized_resx_files(directory_path, "Forge")
delete_localized_resx_files(directory_path, "GenerateUserControl")
delete_localized_resx_files(directory_path, "RAGControl")
delete_localized_resx_files(directory_path, "PasswordPrompt")