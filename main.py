"""
This script reads all `.xlsx` documents from the `documents_input/` directory
and merges them into a single `.xlsx` document saved in `documents_output/`
"""

import datetime
import pandas as pd
import os


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

with pd.ExcelWriter(result_doc) as writer:
    for document in source_docs:
        for sheet_num in range(0, 4):
            sheet_name = pd.ExcelFile(document).sheet_names[sheet_num]
            df = pd.read_excel(document)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
