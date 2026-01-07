import os

FOLDER_DDR = "DDR_TEMP"
FOLDER_MOVING = "Moving_TEMP"
FOLDER_POB = "POB_TEMP"
FOLDER_DCR= "DCR_TEMP"

RIGS = ["O1", "O2", "O3"]

INPUT_DIRECTORIES = []

for rig in RIGS:
    INPUT_DIRECTORIES.append(os.path.join(FOLDER_DDR, rig))
    INPUT_DIRECTORIES.append(os.path.join(FOLDER_MOVING, rig))
    INPUT_DIRECTORIES.append(os.path.join(FOLDER_POB, rig))
    INPUT_DIRECTORIES.append(os.path.join(FOLDER_DCR, rig))

# --- مسیرهای خروجی و بکاپ ---
MAIN_OUTPUT_DIR = 'OUTPUT'
MAIN_BACKUP_DIR = 'BACKUP'


# مسیر فایل coordinates_points.json
COORDINATES_POINTS_PATH="coordinates_points.json"



