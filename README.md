# Easy Access Sheet Toolkit
*September 2024*

Note: only tested on windows systems (actually, only tested on my personal computer with very specific data and settings)

The Easy Access Sheet Toolkit is a python script with a built-in CLI developed to automate the process of creating sheets of data per faculty from a university's CopyRight tool.

This script can do the following:
- Read data from a CopyRight export
- Process the data into a standard format for usage by people who check the documents on canvas
- Add a few extra columns to the data to improve the workflow
- Export various sheets:
    - per faculty
    - all items
    - only items that have been changed
- format & style those excel sheets for easy usage
- Read back the data from all the sheets to produce a CopyRight 'import' sheet -- currently under development.


# Setting up things to use the toolkit

Here are some sort instructions on how to get started with the toolkit. For more details, take a look in the easy_access_cli.py file -- it contains a lot of comments. You probably will want to change things in the script to match your own workflows anyway. Feel free to get in touch if you have questions.

### Settings: settings.env & department_mapping.json

In settings.env you can change where the script will look for the CopyRight export, and where it will create the output files. You can just leave this as is without issues.

department_mapping.json is a simple json file that maps the programme names in the column 'Department' to their corresponding faculties. This is necessary because CopyRight does not have information about which faculty a programme belongs to.
An example file is included, with the mapping for the University of Twente.

**You will need to change this to match your universities' programme names to faculty names to get useful output per faculty!!**
If you do not do so, the script will match all items to the 'Unmapped' faculty.


### Export data from CopyRight

Go to SURF's CopyRight tool, filter the data you want to export, press right click on the sheet -> press download as ... -> select data --> press export --> press 'Click here to download your data file'.
Make a folder called 'raw_copyright_data' (or whatever name you changed it to in settings.env) in the same directory as the easy_access_cli.py file, and put the downloaded file in it.

### Running the cli

Run the cli with standard settings to create the initial sheets:

    uv run easy_access_cli.py

Or, if you want to check out all available cli options:

    uv run easy_access_cli.py --help


# uv? What is that?

UV is an all-in-one python manager. It's a great way to run python scripts and manage dependencies: very easy, very fast, and it runs completely separate from any other python installations on your computer. I -highly- recommend using it in general, but especially for this script.
Of course, you can use any python installation you want to run the script, but using uv is definitely the easiest way to get started.

## Installing uv

Install uv by opening Powershell (press windows key, type 'powershell', enter) and pasting the following line:

    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

and then press enter to install.
Once uv is done installing, close PowerShell and start it up again.
Then type in:

    uv python install

press enter, and you've fully installed uv!
Run  the help info of the easy access script with the following command:
    uv run easy_access_cli.py --help

To run the script using default settings, use the following command:
    uv run easy_access_cli.py
