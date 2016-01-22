#!python

import os
import sys
import re
import random
import zipfile
import shutil
import xml.etree.ElementTree as etree

TEMPLATE_DIR=os.path.join(os.path.dirname(os.path.abspath(__file__)),'template')

def makeWorkbook(path, sheet_data, **format_vals):
  # if sheet data is an etree.Element, dump it to a string
  if type(sheet_data) is etree.Element:
    sheet_data = etree.dump(sheet_data)

  # separate the requested path into directory and filename
  dirname, filename = os.path.split(path)
  # duplicate the workbook template directory
  unpacked_dir = os.path.join(dirname,filename[:filename.rindex('.')])
  i=0
  while True:
    try:
      shutil.copytree(TEMPLATE_DIR, unpacked_dir)
    except (IOError, WindowsError):
      unpacked_dir = re.sub(r'(__\d{'+str(len(str(i)))+'})?$', '__'+str(i), unpacked_dir)
      i += 1
      continue
    break
  
  # generate some ids
  format_vals['workbookId'] = random.randint(0,sys.maxint)
  format_vals['sheetId'] = random.randint(0,sys.maxint)
  format_vals['sheetData'] = sheet_data
  # make sure they're different
  while format_vals['workbookId'] == format_vals['sheetId']:
    format_vals['sheetId'] = random.randint(0,sys.maxint)

  # walk the template dir, filling in all the 
  for tdir, tsubs, tfiles in os.walk(unpacked_dir):
    # for each template file in a dirwalk:
    for tfile in tfiles:
      tpath = os.path.join(tdir,tfile)
      # read in the file
      with open(tpath) as tfin:
        tfstr = tfin.read()
      # fill in the template and overwrite the file
      with open(tpath, 'w') as tfout:
        tfout.write(tfstr.format(**format_vals))
  
  # zip up the template dir into a workbook
  zipf = zipfile.ZipFile(path, 'w')
  zipdir(unpacked_dir, zipf)
  zfp = zipf.fp
  zipf.close()
  print zipf.filename,'is closed:',zfp.closed
  # delete the template dir; no longer needed
  shutil.rmtree(unpacked_dir)


# yoinked from: http://stackoverflow.com/questions/1855095/how-to-create-a-zip-archive-of-a-directory
def zipdir(path, ziph):
  if path.endswith('\\') or path.endswith('/'):
    path = path[:-1]
  # ziph is zipfile handle
  for root, dirs, files in os.walk(path):
    for fname in files:
      fpath = os.path.join(root, fname)
      ziph.write(fpath, fpath[len(path)+1:])

