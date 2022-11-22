#!/usr/bin/env python

import sys
import os
import logging
import pandas as pd
import math

from iptcinfo3 import IPTCInfo

# logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.DEBUG)
logger = logging.getLogger("metaphotos")

ENCODING = 'latin-1'
COL_NAMES = ["Name", "Description", "Keywords", "Copyright"]
META_ID_TO_STR = {
    120: 'caption/abstract',
    116: 'copyright notice',
    25: 'keywords',
    15: 'category',
    20: 'supplemental category',
    118: 'contact'
}


def is_empty_or_nan(v):
    res = not v
    if not res and isinstance(v, float):
        res = math.isnan(v)
    return res
    

def main(input_path, fname_xls, output_path=None):

    stats = {
        "Input path": input_path,
        "XLS File": fname_xls,
        "Output path": output_path or "not specified",
        "Num. Lines Read": 0,
        "Num. Files Processed": 0,
        "Num. Files Not Found": 0,
        "Empty Lines Found": [],
        "Num. Files Updated And Saved": 0,
    }
    df = pd.read_excel(fname_xls, index_col=None, names=COL_NAMES)
    logger.debug(f"{df.keys()=} {len(df)=} {df=}")

    for i in range(len(df)):

        basename_img = df.at[i, "Name"]
        # NULL CELLs are interpreted as float('nan')
        if is_empty_or_nan(basename_img):
            # Empty line, skip
            stats["Empty Lines Found"].append(i+2)
            continue
        
        stats["Num. Lines Read"] += 1
        fname_img = os.path.join(input_path, basename_img)
        logger.info(f"\n[START] Processing line {i}, image={fname_img} ...")
        logger.debug(f"{fname_img=} {df.at[i, 'Name']=}")
        if os.path.exists(fname_img):

            stats["Num. Files Processed"] += 1
            # Check if any info is provided, otherwise skip the line
            keywords_cell = df.at[i, 'Keywords']
            caption_cell = df.at[i, 'Description']
            copyright_cell = df.at[i, 'Copyright']
            if is_empty_or_nan(keywords_cell) and is_empty_or_nan(caption_cell) and is_empty_or_nan(keywords_cell):
                stats["Empty Lines Found"].append(i+2)
                continue
                
            iptc = IPTCInfo(fname_img, inp_charset=ENCODING, out_charset=ENCODING)

            # Update keywords if specified
            if not is_empty_or_nan(keywords_cell):
                keywords = [x.strip() for x in keywords_cell.split(',')]
                iptc["keywords"] = keywords

            # Update caption if specified
            if not is_empty_or_nan(caption_cell):
                iptc["caption/abstract"] = caption_cell.strip()

            # Update copyright if specified
            if not is_empty_or_nan(copyright_cell):
                iptc["copyright notice"] = copyright_cell.strip()

            if output_path:
                fname_dest_img = os.path.join(output_path, basename_img)
                iptc.save_as(fname_dest_img)
                stats["Num. Files Updated And Saved"] += 1
                logger.info(f"[END] Processed line {i}, saved image={fname_dest_img}")

            else:
                for k, v in iptc._data.items():

                    key = META_ID_TO_STR.get(k, k)
                    value = v

                    if v and isinstance(v, list) and isinstance(v[0], bytes):
                        value = [x.decode() for x in v]
                    elif isinstance(v, bytes):
                        value = v.decode()

                    logger.info(f"* {key}: {value}")

                logger.info(f"[END] Processed line {i}, not saved image={fname_img}")

        else:
            stats["Num. Files Not Found"] += 1

    print("\n---- STATS ----")
    for k, v in stats.items():
        print(f"- {k}: {v}")
    print()


if __name__ == "__main__":

    if len(sys.argv) < 3:
        print(f"Usage: {sys.argv[0]} <input_path> <filename.xls>")
        sys.exit(100)

    input_path = sys.argv[1]
    if not os.path.isdir(input_path):
        print(f"The directory '{input_path}' does not exist")
        sys.exit(101)

    fname_xls = sys.argv[2]
    if not os.path.exists(fname_xls):
        print(f"The filename '{fname_xls}' does not exist")
        sys.exit(102)

    output_path = None
    if len(sys.argv) > 3:
        output_path = sys.argv[3]
        # Create output dir if not exists
        os.makedirs(output_path, exist_ok=True)
    else:
        logging.basicConfig(format='%(message)s', level=logging.INFO)

    main(input_path, fname_xls, output_path)
