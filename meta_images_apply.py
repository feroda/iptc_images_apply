#!/usr/bin/env python
"""
Lo script prende in ingresso/input: 

- il percorso della cartella con le immagini da processare
- il percorso del foglio di calcolo che contiene i metadati
- il percorso della cartella in cui verranno copiati i file processati (se non esiste viene creata)

Esempio file:
Name	Description	Keywords	Copyright Notice
62507-c.jpg;Rare and endangered trees unique to China, 1958;Rare , endangered, trees ,unique, plants, botany, wildlife, China, 1958, 1950s;	Colaimages

Per ogni immagine nella cartella, lo script: 

- leggerà dal foglio di calcolo i metadati da applicare
- se esistono, creerà una copia dell'immagine nella cartella di destinazione con i metadati applicati
- stamperà un report di quanti file sono stati processati
"""

import sys
import os
import logging
import pandas as pd

from iptcinfo3 import IPTCInfo

logging.basicConfig(format='%(levelname)s:%(message)s', level=logging.DEBUG)
logger = logging.getLogger("metaphotos")

COL_NAMES = ["Name", "Description", "Keywords", "Copyright"]


def main(argv):

    input_path = argv[1]
    fname_xls = argv[2]

    output_path = None
    if len(argv) > 3:
        output_path = argv[3]
        # Create output dir if not exists
        os.makedirs(output_path, exist_ok=True)

    stats = {
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
        saved = False
        fname_img = os.path.join(input_path, basename_img)
        logger.debug(f"{fname_img=} {df.at[i, 'Name']=}")
        if os.path.exists(fname_img):

            # Check if any info is provided, otherwise skip the line
            keywords_cell = df.at[i, 'Keywords']
            caption_cell = df.at[i, 'Description']
            copyright_cell = df.at[i, 'Copyright']
            save_the_file = False
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

            for k, v in iptc._data.items():
                if isinstance(v, str):
                    logger.debug("{} {}".format(k, v))
                elif isinstance(v, list):
                    try:
                        logger.debug("{} {}".format(k, [x.decode() for x in v]))
                    except:
                        logger.debug("{} {}".format(k, v))
                else:
                    logger.debug("{} {}".format(k, v.decode()))

            if output_path:
                iptc.save_as(os.path.join(output_path, basename_img))
                stats["Num. Files Updated And Saved"] += 1
                saved = True

            print(f"- Processed line {i} with image={fname_img} {saved=}")

    print(stats)


if __name__ == "__main__":
    main(sys.argv)
