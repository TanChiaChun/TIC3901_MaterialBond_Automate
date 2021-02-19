# Import from packages
import os
import argparse
import logging
from glob import glob
from shutil import move
from xml.etree import ElementTree
from pywinauto import application
from pywinauto.findwindows import ElementNotFoundError

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
# Validate arguments
if not os.path.exists(app_path):
    handle_exception("Application not found!")
if not os.path.exists(xml_folder):
    handle_exception("XML folder not found!")

# Create archive folder
os.makedirs(xml_folder_archive, exist_ok=True)

# Get list of XMLs and sort descending by date
xmls = sorted(glob(xml_folder + "/*.xml"), key=os.path.getmtime)
if len(xmls) == 0:
    finalise_app("XML folder empty! No part to process")

# Parse XML and add data to dict
selected_xml_path = xmls[0]
selected_xml = os.path.basename(selected_xml_path)
logger.info("Read " + selected_xml)
try:
    tree = ElementTree.parse(selected_xml_path)
    root = tree.getroot()
    for child1 in root:
        for child2 in child1:
            data_dict[child2.tag] = child2.text
except ElementTree.ParseError:
    handle_exception("XML file not recognised!")
except Exception:
    handle_exception("XML file error!")

# Start external app and input data
logger.info("Starting " + os.path.basename(app_path))
app = application.Application(backend="uia").start(app_path)
try:
    logger.debug("Inputting Serial Number: " + data_dict["SerialNumber"])
    app["Material Bond"]["Serial NumberEdit"].type_keys(data_dict["SerialNumber"])
    logger.debug("Inputting Material A: " + data_dict["MaterialA"])
    app["Material Bond"]["Material AEdit"].type_keys(data_dict["MaterialA"])
    logger.debug("Inputting Material B: " + data_dict["MaterialB"])
    app["Material Bond"]["Material BEdit"].type_keys(data_dict["MaterialB"])
except ElementNotFoundError:
    handle_exception("Error communicating with Material Bond application!")
except Exception:
    handle_exception("Material Bond application error")

# Archive XML
move(selected_xml_path, xml_folder_archive + '/' + selected_xml)
logger.info("Archived "+ selected_xml)

finalise_app()