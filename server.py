#!python

import os
import re
import string
import xlsxwriter
from bottle import __version__ as bottle_version
from bottle import route, post, static_file, request, hook, run
from PIL import Image

TMP_DIR=os.path.dirname(os.path.abspath(__file__))+'/tmp'
print "tmp dir is", TMP_DIR
if not os.path.exists(TMP_DIR):
  print "making tmp dir"
  os.makedirs(TMP_DIR)

print "Using Bottle v{v}".format(v=bottle_version)

@route('/')
def index():
  return static_file('index.html', root=os.path.dirname(__file__))

@post('/imgspreadsheet')
def img_spreadsheet():
  print 'generating spreadsheet'
  # grab uploaded file
  img = request.files.get('pic')
  # get file name
  _, filename = os.path.split(img.filename)
  print 'image:', filename
  # build path to temporary image file
  img_path = os.path.join(TMP_DIR, filename)
  # save temporary copy of uploaded image
  i=0
  while True:
    try:
      img.save(img_path)
    except (IOError, WindowsError):
      img_path = re.sub(r'(__\d{'+str(len(str(i)))+'})?(\.\w+)$', '__'+str(i)+r'\2', img_path)
      i += 1
      continue
    break
  
  # insert rows into an excel workbook
  spreadsheet_filename =  filename[:filename.rindex('.')]+'.xlsx'
  print 'generating workbook:', spreadsheet_filename
  spreadsheet_path = os.path.join(TMP_DIR, spreadsheet_filename)
  # convert image to spreadsheet rows
  img_to_workbook(img_path, spreadsheet_path)
  
  # set up a hook to delete the image and spreadsheet after sending to the user
  #@hook('after_request')
  #def delete_spreadsheet():
  #  os.remove(img_path)
  #  os.remove(spreadsheet_path)
  # return the excel workbook
  return static_file(spreadsheet_filename, root=TMP_DIR, download=spreadsheet_filename)

def img_to_workbook(imgfile, bookfile, subpixels=False):
  PIXEL_SIZE = 15
  
  # create a workbook and worksheet to store our image
  book = xlsxwriter.Workbook(bookfile)
  sheet = book.add_worksheet(os.path.basename(imgfile)[:31])

  # grab the image's pixels
  with Image.open(imgfile) as img:
    img_data = img.getdata()
    img_width, img_height = img.size

  if subpixels:
    for row in range(img_height):
      #print "row", row
      r_row = row*3
      g_row = row*3+1
      b_row = row*3+2
      for col in range(img_width):
        r,g,b = img_data[row*img_height+col]
        # add color parts to the sheet in separate rows
        sheet.write(r_row, col, r)
        sheet.write(g_row, col, g)
        sheet.write(b_row, col, b)
      sheet.conditional_format(r_row, 0, r_row, col, {
        'type': '2_color_scale',
        'min_color': 'black',
        'max_color': 'red',
        'min_value': 0,
        'max_value': 255
      })
      sheet.conditional_format(g_row, 0, g_row, col, {
        'type': '2_color_scale',
        'min_color': 'black',
        'max_color': 'lime',
        'min_value': 0,
        'max_value': 255
      })
      sheet.conditional_format(b_row, 0, b_row, col, {
        'type': '2_color_scale',
        'min_color': 'black',
        'max_color': 'blue',
        'min_value': 0,
        'max_value': 255
      })
      # set each row to one pixel height
      sheet.set_row(r_row, PIXEL_SIZE)
      sheet.set_row(g_row, PIXEL_SIZE)
      sheet.set_row(b_row, PIXEL_SIZE)
    
    # set columns to one pixel width
    sheet.set_column(0, col, PIXEL_SIZE*3)
      
  else:
    for row in range(img_height):
      for col in range(img_width):
        r,g,b = img_data[row*img_height+col]
        # convert color tuple to int
        #color = r*pow(16,4) + g*pow(16,2) + b
        color = (r + g + b) / 3
        # add color to the sheet
        sheet.write(row, col, color)
        
      sheet.conditional_format(row, 0, row, col, {
        'type': '2_color_scale',
        'min_color': 'black',
        'max_color': 'white',
        'min_value': 0,
        'max_value': 255
      })
      # set row to one pixel height
      sheet.set_row(row, PIXEL_SIZE)
    
    # set columns to one pixel width
    sheet.set_column(0, col, PIXEL_SIZE)
  
      
  book.close()
  return book


def genlabel(n):
  label=''
  while n>0:
    n, r = divmod(n-1, 26)
    label = string.ascii_uppercase[r] + label
  return label

  

run(host='localhost', port=8080, debug=True)

