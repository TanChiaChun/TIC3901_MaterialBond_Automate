# Import from packages
import os
import argparse
import logging
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
xml_file = "data/input/Mes2Oem_OpStart_SN12345.xml"
data_dict = {}
app_path = args.app

##################################################
# Functions
##################################################


##################################################
# Main
##################################################
tree = ElementTree.parse(xml_file)
root = tree.getroot()
for child1 in root:
    for child2 in child1:
        data_dict[child2.tag] = child2.text

app = application.Application(backend="uia").start(app_path)
app["Material Bond"]["Serial NumberEdit"].type_keys(data_dict["SerialNumber"])
app["Material Bond"]["Material AEdit"].type_keys(data_dict["MaterialA"])
app["Material Bond"]["Material BEdit"].type_keys(data_dict["MaterialB"])

finalise_app()