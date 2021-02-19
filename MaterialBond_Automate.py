# Import from packages
import os
import argparse
import logging
from glob import glob
from shutil import move
from xml.etree import ElementTree
from pywinauto import application

# Import from modules
from MyFunctions import initialise_app, finalise_app, handle_exception

# Initialise project
CURR_DIR, CURR_FILE = os.path.split(__file__)
PROJ_NAME = CURR_FILE.split('.')[0]

# Get command line arguments
my_arg_parser = argparse.ArgumentParser(description=f"{PROJ_NAME}")
my_arg_parser.add_argument("--app", help="Enter path to application")
my_arg_parser.add_argument("--xml", help="Enter path to application")
my_arg_parser.add_argument("--log", help="DEBUG to enter debug mode")
args = my_arg_parser.parse_args()

# Initialise app
initialise_app(PROJ_NAME, args.log)
logger = logging.getLogger("my_logger")

# # Get environment variables
# env_var1 = os.getenv("env_var1")
# env_var2 = os.getenv("env_var2")
# if env_var1 == None or env_var2 == None:
#     handle_exception("Missing environment variables!")

# Declare variables
xml_folder = args.xml
xml_folder_archive = xml_folder + "/archive"
data_dict = {}
app_path = args.app

##################################################
# Functions
##################################################


##################################################
# Main
##################################################
os.makedirs(xml_folder_archive, exist_ok=True)

xmls = sorted(glob(xml_folder + "/*.xml"), key=os.path.getmtime)

if len(xmls) == 0:
    finalise_app("XML folder empty! No part to process")

tree = ElementTree.parse(xmls[0])
root = tree.getroot()
for child1 in root:
    for child2 in child1:
        data_dict[child2.tag] = child2.text

app = application.Application(backend="uia").start(app_path)
app["Material Bond"]["Serial NumberEdit"].type_keys(data_dict["SerialNumber"])
app["Material Bond"]["Material AEdit"].type_keys(data_dict["MaterialA"])
app["Material Bond"]["Material BEdit"].type_keys(data_dict["MaterialB"])

move(xmls[0], xml_folder_archive + '/' + os.path.basename(xmls[0]))

finalise_app()