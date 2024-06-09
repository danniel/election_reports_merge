"""
This script reads all `.xlsx` documents from the `documents_input/` directory
and merges them into a single `.xlsx` document saved in `documents_output/`
"""

import datetime
import logging
import os
import pandas as pd


logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Normalized sheet names
sheet_index_names = (
    (0, "PSI"),
    (1, "A"),
    (2, "B"),
    (3, "C"),
)

result_doc = r"documents_output/Compiled_{}.xlsx".format(
    datetime.datetime.now(datetime.UTC).strftime("%Y%m%d_%H%M%S_UTC")
)

source_dir = r"documents_input/"
source_docs = [
    os.path.join(root, file)
    for root, folder, files in os.walk(source_dir)
    for file in files
    if file.endswith(".xlsx")
]

# Open the output file for writing
with pd.ExcelWriter(result_doc) as writer:
    logger.info("Opened %s for writing.", result_doc)

    # Read data from each input file
    for document in source_docs:
        logger.info("Opened %s for reading.", document)

        # Read data from each input file sheet
        for sheet_num in range(0, 4):
            sheet_name = pd.ExcelFile(document).sheet_names[sheet_num]
            df = pd.read_excel(document, sheet_name)
            df.to_excel(writer, sheet_name=sheet_index_names[sheet_num][1], index=False)

    logger.info("Finished processing all source documents.")
