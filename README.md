# Easy Access Sheet Toolkit
*April 2025*
Note: This project has been significantly updated, changed, and expanded in the meantime, but is not in a 'shareable' state yet. Once it is (if it is...) I'll update the README... someday!

*February 2025*

Note: only tested on windows systems (actually, only tested on my personal computer with very specific data and settings)

The Easy Access Sheet Toolkit is a python script with a built-in CLI developed to automate the process of creating sheets of data per faculty from a university's CopyRight tool.

This script does the following:
- Read data exported by the CopyRight tool
- Process the data into a standard format
- Add a few extra columns to the data to improve the workflow
- Enrich each item with data from OSIRIS and People Pages
- Export various sheets:
    - per faculty
    - all items
    - only items that have been changed
- format & style those excel sheets for easy usage
- Read back the data from all the sheets to produce a CopyRight 'import' sheet -- currently under development.

Take a look at the cli help function:
![image](https://github.com/user-attachments/assets/c2038652-3e68-4ebf-9bec-e677721e001a)

# Setting up things to use the toolkit

Here are some basic instructions on how to get started with the toolkit. For more details, take a look in the source code -- it contains a lot of comments. You probably will want to change things in the script to match your own workflows anyway. Feel free to get in touch if you have questions.

### Settings: settings.yaml

settings.yaml contains most settings for the app. 

The file should include a hierarchy of the universities faculties' and programmes. This is necessary to split the results per faculty and/or programme because the CopyRight tool does not have information about which faculty a programme belongs to.
The included settings.yaml file includes all the settings used at the University of Twente. If you do not change the information, the script will not be able to recognize any file, and it will then move all items to the 'Unmapped' faculty.

### Adding data from the CopyRight tool

Go to SURF's CopyRight tool, filter the data you want to export, press right click on the sheet -> press download as ... -> select data --> press export --> press 'Click here to download your data file'.
Put it in the folder called 'raw_copyright_data' (or whatever name you changed it to in settings.yaml) in the same directory as the run.py file, and put the downloaded file in it. 

### Running the cli

Run the cli with standard settings to create the initial sheets:

    uv run easy_access_cli.py

If all is well, you'll see something like this:
![image](https://github.com/user-attachments/assets/0724f886-f067-45d7-98f3-67e9eff8e25e)
If something is wrong, you'll either get a short warning/error message with the specific issue; or you'll get a giant list of detailed errors. Probably the last. 

If you want to check out all available cli options use the --help option (see the screenshot above):

    uv run easy_access_cli.py --help


# uv? What is that?

[uv](https://docs.astral.sh/uv/getting-started/installation/) is an all-in-one python manager. It's a great way to run python scripts and manage dependencies: very easy, very fast, and it runs completely separate from any other python installations on your computer. I -highly- recommend using it in general, but especially for this script.
Of course, you can use any python installation you want to run the script, but using uv is definitely the easiest way to get started.

## Installing uv

On windows, install uv by opening Powershell (press windows key, type 'powershell', enter) and pasting the following line:

    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

and then press enter to install.
Once uv is done installing, close PowerShell and start it up again.
Then type in:

    uv python install

press enter, and you've fully installed uv!
Run  the help info of the easy access script with the following command:

    uv run run.py --help

To run the script using default settings, don't include --help:

    uv run run.py
