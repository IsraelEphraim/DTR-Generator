from cx_Freeze import setup, Executable

# Specify the list of files and options for the build
build_exe_options = {
    "packages": ["webbrowser", "pandas", "openpyxl"],  # Add any required packages
    "excludes": [],
    "include_files": ["template", "static"]  # Add any additional files or directories
}

setup(
    name="EastWest TimeKeeping",
    version="1.0",
    description="Topserve Time Keeping for EastWest",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base="Win32GUI")]
)
