from __future__ import print_function  # better print function
import os                              # file/directory functions
import shutil                          # for copying directories recursively
from datetime import datetime          # for datetime functions

# ----------------------------------------------------------------------
# OVERVIEW: This program is intended to copy Storage folder from API
# servers to a destination folder on this file server. Please check
# the SETTINGS block below before executing.
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# SETTINGS BLOCK - BEGIN
# ----------------------------------------------------------------------

# Destination directory - this is where Storage folders will be copied to:
dest_dir = "/Shared/API_preupgrade_backups/Storage"

# List of all possible paths to Storage directory on TEST machines:
dirs_test = ("/Log Files/API Healthcare/APIHealthcare/Test/Storage",
             "/Log Files/API/APIHealthcare/Test/Storage",
             "/Program Files/API Healthcare/Application Server/Test/All Devices/Storage",
             "/Program Files/API/Application Server/Test/All Devices/Storage",
             "/Program Files/API Healthcare/Application Server/Test/Telephony/Storage",
             "/Program Files/API/Application Server/Test/Telephony/Storage")

# List of all possible paths to Storage directory on LIVE machines:
dirs_live = ("/Log Files/API Healthcare/APIHealthcare/Live/Storage",
             "/Log Files/API/APIHealthcare/Live/Storage",
             "/Program Files/API Healthcare/Application Server/Live/All Devices/Storage",
             "/Program Files/API/Application Server/Live/All Devices/Storage",
             "/Program Files/API Healthcare/Application Server/Live/Telephony/Storage",
             "/Program Files/API/Application Server/Live/Telephony/Storage")

# Above dirs will be tried one by one on each machine until found one that exists.
# Always use forward slashes.

# List of all TEST machines, whether or not having Storage directory:
machines_test = ("LBX-PRI-T-AP1",
                 "LBX-AGT-T-AP1",
                 "LBX-AGT-T-AP2",
                 "LBX-AGT-T-AP3",
                 "LBX-AGT-T-AP4",
                 "LBX-AGT-T-AP5",
                 "LBX-SQLRS-T-AP1",
                 "LBX-AGT-T-AP6",
                 "LBX-RPT-T-AP1",
                 "LBX-WPS-T-WS1",
                 "LBX-WPS-T-WS2",
                 "LBX-WPS-T-WS3",
                 "LBX-WPS-T-WS4",
                 "LBX-WPS-T-WS5",
                 "LBX-WPS-T-WS6",
                 "LBX-WPS-T-WS7",
                 "LBX-WPS-T-WS8"
                 )
# List of all LIVE machines, whether or not having Storage directory:
machines_live = ("LBX-PRI-P-AP1",
                 "LBX-AGT-P-AP1",
                 "LBX-AGT-P-AP2",
                 "LBX-AGT-P-AP3",
                 "LBX-AGT-P-AP4",
                 "LBX-AGT-P-AP5",
                 "LBX-AGT-P-AP6",
                 "LBX-AGT-P-AP7",
                 "LBX-AGT-P-AP8",
                 "LBX-AGT-P-AP9",
                 "LBX-AGT-P-AP10",
                 "LBX-AGT-P-AP11",
                 "LBX-SQLRS-P-AP1",
                 "LBX-TC-P-AP1",
                 "LBX-TC-P-AP2",
                 "LBX-RPT-P-AP1",
                 "LBX-WPS-P-WS1",
                 "LBX-WPS-P-WS2",
                 "LBX-WPS-P-WS3",
                 "LBX-WPS-P-WS4",
                 "LBX-WPS-P-WS5",
                 "LBX-WPS-P-WS6",
                 "LBX-WPS-P-WS7",
                 "LBX-WPS-P-WS8",
                 "LBX-WPS-P-WS9",
                 "LBX-WPS-P-WS10",
                 "LBX-WPS-P-WS11",
                 "LBX-WPS-P-WS12",
                 "LBX-WPS-P-WS13",
                 "LBX-WPS-P-WS14",
                 "LBX-SQLCL-PCL1"
                 )

# ----------------------------------------------------------------------
# SETTINGS BLOCK - END
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# MAIN PROGRAM BLOCK - BEGIN
# ----------------------------------------------------------------------

# Open a log file in append mode
log = open("backup_storage.log", "a")

print("-" * 67)
print("This program will backup STORAGE directory found on API servers")
print("to directory: {}".format(dest_dir))
print("To check exact settings, open this program file in edit mode.")
print("-" * 67)

# Get run mode from user:
run_mode_raw = raw_input("Enter the execution mode [test, live]: ")
run_mode = run_mode_raw.lower().strip()

if run_mode == "test":
    dirs = dirs_test
    machines = machines_test
elif run_mode == "live":
    dirs = dirs_live
    machines = machines_live
else:
    print ("You entered an invalid value for execution mode: '{}'".format(run_mode))
    print ("Re-launch this program and try again.")
    user_input = raw_input("Press Enter to exit...")
    quit()

print ("Thank you. The program will run for all {} machines.".format(run_mode.upper()))
user_input = raw_input("Are you ready to proceed? [y/n]: ")
if user_input.lower() == "y":
    
    # Get current date/time and format it:
    # - short format for use in folder names
    # - long format for use in logging
    raw_current_datetime = datetime.now()
    current_datetime_short = raw_current_datetime.strftime("%Y%m%d-%H%M")
    current_datetime_long = raw_current_datetime.strftime("%Y-%m-%d %H:%M")

    print ("-" * 67, file=log)
    print ("Program started ", current_datetime_long, file=log)
    print ("-" * 67, file=log)

    print ("-" * 67)
    print ("Program started ", current_datetime_long)
    print ("-" * 67)
    
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    # Construct a destination directory for this program run
    dest_dir_thisrun = "{0}/Storage_{1}".format(dest_dir, current_datetime_short)
    if os.path.exists(dest_dir_thisrun):
        # if exists, it was created a minute ago, so safe to delete
        shutil.rmtree(dest_dir_thisrun)
    else:
        try:
            os.mkdir(dest_dir_thisrun)
        except:
            print ("Failed to create a destination directory:", dest_dir_thisrun, file=log)
            print ("-" * 67, file=log)
            print ("Finished.", file=log)
            print ("-" * 67, file=log)
            print ("Failed to create a destination directory:", dest_dir_thisrun)
            user_input = raw_input("Press Enter to exit...")
            quit()
        
    for machine in machines:

        found = 0
        machine_name_length = len(machine)
        extra_space = 20 - machine_name_length
        print("{0}...".format(machine), " " * extra_space, end="")

        for dir in dirs:
            # construct a full path to Storage directory:
            src = "//{0}/c${1}".format(machine, dir)
            if os.path.exists(src):
                found = 1
                # Storage directory found. Copy it into target directory.
                try:
                    dest = dest_dir_thisrun + "/" + machine + "/Storage"
                    shutil.copytree(src, dest)
                    print("Storage copied successfully.")
                    print("{0}...".format(machine), " " * extra_space, "Storage copied successfully.", file=log)
                except:
                    print("Storage copying FAILED. Try copying manually.")
                    print("{0}...".format(machine), " " * extra_space, "Storage copying FAILED. Try copying manually.", file=log)
                break
        if found == 0:
            print("Storage NOT found on this machine.")
            print("{0}...".format(machine), " " * extra_space, "Storage NOT found on this machine.", file=log)

    print ("-" * 67, file=log)
    print ("Finished.", file=log)
    print ("-" * 67, file=log)

    print ("-" * 67)
    print ("Finished.")
    print ("-" * 67)

    user_input = raw_input("Press Enter to exit...")

else:

    print ("Cancelled.")

# ----------------------------------------------------------------------
# MAIN PROGRAM BLOCK - END
# ----------------------------------------------------------------------
