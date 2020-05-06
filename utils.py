from datetime import datetime
from pathlib import Path
import re

re_timestamp = re.compile('[0-9]{8}-[0-9]{6}')

def save_invoice_with_timestamp(df, path):
    # remove old timestamp
    stem = re.sub(r'__[0-9]{8}-[0-9]{6}','', path.stem)
    
    # create new timestamp
    timestamp = str(datetime.now().strftime("%Y%m%d-%H%M%S"))
    
    output = path.parent / (stem + '__' + timestamp + '.xlsx')
    print('save_invoice_with_timestamp: ' + str(output))   
    df.to_excel(output, index=False)
    

def find_latest_invoice_version(invoice_path):
    re_version = re.compile(invoice_path.stem + '__[0-9]{8}-[0-9]{6}' + invoice_path.suffix)
    
    versions = []
    tmp = list(invoice_path.parent.glob("*" + invoice_path.suffix))
    pattern = invoice_path.stem + '__[0-9]{8}-[0-9]{6}' + invoice_path.suffix
    #print(pattern)
    for t in tmp:
        if re.search(pattern, str(t)):
            #print(t)
            versions.append(t)
    
    latest = invoice_path
    if len(versions) > 0:
        latest = sorted(versions)[-1]
    
    print('find_latest_invoice_version: ' + str(latest))   
    return latest

