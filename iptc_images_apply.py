#!/usr/bin/env python

import sys
import os
import logging
import pandas as pd

from iptcinfo3 import IPTCInfo

# logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.DEBUG)
logger = logging.getLogger("metaphotos")

COL_NAMES = ["Name", "Description", "Keywords", "Copyright"]
META_ID_TO_STR = {
    120: 'caption/abstract',
    116: 'copyright notice',
    25: 'keywords',
    15: 'category',
    20: 'supplemental category',
    118: 'contact'
}


def main(argv):

    input_path = argv[1]
    fname_xls = argv[2]

    output_path = None
    if len(argv) > 3:
        output_path = argv[3]
        # Create output dir if not exists
        os.makedirs(output_path, exist_ok=True)
    else:
        logging.basicConfig(format='%(message)s', level=logging.INFO)


    stats = {
        "Input path": input_path,
        "XLS File": fname_xls,
        "Output path": output_path or "not specified",
        "Num. Files Read": 0,
        "Num. Files Updated And Saved": 0,
    }
    df = pd.read_excel(fname_xls, index_col=None, names=COL_NAMES)
    logger.debug(f"{df.keys()=} {len(df)=} {df=}")

    for i in range(len(df)):

        basename_img = df.at[i, "Name"]
        if not basename_img:
            # Empty line, skip
            continue
        
        stats["Num. Files Read"] += 1
        fname_img = os.path.join(input_path, basename_img)
        logger.info(f"\n[START] Processing line {i}, image={fname_img} ...")
        logger.debug(f"{fname_img=} {df.at[i, 'Name']=}")
        if os.path.exists(fname_img):

            # Check if any info is provided, otherwise skip the line
            keywords_cell = df.at[i, 'Keywords']
            caption_cell = df.at[i, 'Description']
            copyright_cell = df.at[i, 'Copyright']
            if not keywords_cell and not caption_cell and not keywords_cell:
                continue
                
            iptc = IPTCInfo(fname_img)

            # Update keywords if specified
            if keywords_cell:
                keywords = [x.strip() for x in keywords_cell.split(',')]
                iptc["keywords"] = keywords

            # Update caption if specified
            if caption_cell:
                iptc["caption/abstract"] = caption_cell.strip()

            # Update copyright if specified
            if copyright_cell:
                iptc["copyright notice"] = copyright_cell

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

    print("\n---- STATS ----")
    for k, v in stats.items():
        print(f"- {k}: {v}")
    print()


if __name__ == "__main__":
    main(sys.argv)