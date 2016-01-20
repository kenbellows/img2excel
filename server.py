#!python

import os
import re
import string
import xml.etree.ElementTree as etree
from bottle import __version__ as bottle_version
from bottle import route, post, static_file, request, hook, run
from PIL import Image
from workbook import workbook

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
  
  # convert image to spreadsheet rows
  sheetData = img_to_sheetData(img_path)
  # insert rows into an excel workbook
  spreadsheet_filename =  filename[:filename.rindex('.')]+'.xlsx'
  print 'spreadsheet:', spreadsheet_filename
  spreadsheet_path = os.path.join(TMP_DIR, spreadsheet_filename)
  workbook.makeWorkbook(spreadsheet_path, sheetData)
  
  # set up a hook to delete the image and spreadsheet after sending to the user
  @hook('after_request')
  def delete_spreadsheet():
    os.remove(img_path)
    os.remove(spreadsheet_path)
  # return the excel workbook
  return static_file(spreadsheet_filename, root=TMP_DIR, download=spreadsheet_filename)

def img_to_sheetData(filename, subpixels=False):
  # grab the image's pixels
  with Image.open(filename) as img:
    img_data = img.getdata()
    img_width, img_height = img.size
    #print "image size:", img.size
  # create root sheetData element
  root = etree.Element('x:sheetData')
  for row in range(img_height):
    #print "row", row
    if subpixels:
      r_row = str(row*3)
      r_row_el = etree.SubElement(root, 'row')
      r_row_el.attrib['r'] = r_row
      g_row = str(row*3+1)
      g_row_el = etree.SubElement(root, 'row')
      g_row_el.attrib['r'] = g_row
      b_row = str(row*3+2)
      b_row_el = etree.SubElement(root, 'row')
      b_row_el.attrib['r'] = b_row
      for col in range(img_width):
        col_label = genlabel(col+1)
        # separate out pixel colors
        r,g,b = img_data[row*img_height+col]
        
        # add a cell to the red row
        r_cell_el = etree.SubElement(r_row_el, 'c')
        # cell label of the form 'A1'
        r_cell_el.attrib['r'] = col_label + r_row
        # style index linked to red color scale
        #r_cell_el.attrib['s'] = '1'
        # create a value tag and add this pixel's red color
        r_val = etree.SubElement(r_cell_el, 'v')
        r_val.text = str(r)
        
        # add a cell to the green row
        g_cell_el = etree.SubElement(g_row_el, 'c')
        # cell label of the form 'A1'
        g_cell_el.attrib['r'] = col_label + g_row
        # style index linked to green color scale style
        #g_cell_el.attrib['s'] = '2'
        # create a value tag and add this pixel's green color
        g_val = etree.SubElement(g_cell_el, 'v')
        g_val.text = str(g)
        
        # add a cell to the blue row
        b_cell_el = etree.SubElement(b_row_el, 'c')
        # cell label of the form 'A1'
        b_cell_el.attrib['r'] = col_label + b_row
        # style index linked to blue color scale style
        #b_cell_el.attrib['s'] = '3'
        # create a value tag and add this pixel's blue color
        b_val = etree.SubElement(b_cell_el, 'v')
        b_val.text = str(b)
    
      # create 3 conitional formatting color scales
      # red conditional formatting
      r_cond_format = etree.Element('conditionalFormatting')
      r_cond_format.attrib['sqref'] = ','.join(map(lambda n: str(n)+':'+str(n), range(1,height*3,3)))
      r_cond_format.append(etree.fromstring('''
        <cfRule type="colorScale" priority="2">
          <colorScale>
            <cfvo type="formula" val="0"/>
            <cfvo type="formula" val="255"/>
            <color rgb="FF000000"/>
            <color rgb="FFFF0000"/>
          </colorScale>
        </cfRule>
      '''))
      
      # green conditional formatting
      g_cond_format = etree.Element('conditionalFormatting')
      g_cond_format.attrib['sqref'] = ','.join(map(lambda n: str(n)+':'+str(n), range(2,height*3+1,3)))
      g_cond_format.append(etree.fromstring('''
        <cfRule type="colorScale" priority="2">
          <colorScale>
            <cfvo type="formula" val="0"/>
            <cfvo type="formula" val="255"/>
            <color rgb="FF000000"/>
            <color rgb="FF00FF00"/>
          </colorScale>
        </cfRule>
      '''))
      
      # blue conditional formatting
      b_cond_format = etree.Element('conditionalFormatting')
      b_cond_format.attrib['sqref'] = ','.join(map(lambda n: str(n)+':'+str(n), range(3,height*3+2,3)))
      b_cond_format.append(etree.fromstring('''
        <cfRule type="colorScale" priority="2">
          <colorScale>
            <cfvo type="formula" val="0"/>
            <cfvo type="formula" val="255"/>
            <color rgb="FF000000"/>
            <color rgb="FF0000FF"/>
          </colorScale>
        </cfRule>
      '''))
      
      cond_format = etree.tostring(r_cond_format) +\
                    etree.tostring(g_cond_format) +\
                    etree.tostring(b_cond_format)
     
    else:
      row_s = str(row)
      row_el = etree.SubElement(root, 'row')
      row_el.attrib['r'] = row_s
      for col in range(img_width):
        col_label = genlabel(col+1)
        #if col%100 == 0:
        #  print "\t"+col_label
        # convert color tuple to int
        r,g,b = img_data[row*img_height+col]
        color = r*pow(16,4) + g*pow(16,2) + b
        #int(''.join(map(lambda c: '{:02x}'.format(c), pixel)), 16)
        
        # add a cell to the red row
        cell_el = etree.SubElement(row_el, 'c')
        # cell label of the form 'A1'
        cell_el.attrib['r'] = col_label + row_s
        # style index linked to red color scale
        #cell_el.attrib['s'] = '1'
        # create a value tag and add this pixel's red color
        val = etree.SubElement(cell_el, 'v')
        val.text = str(color)
      
      # generate the conditional formatting for the cells
      # I'm not sure how to do full color with conditional
      # formatting alone; this is only grayscale
      cond_format_el = etree.Element('conditionalFormatting')
      cond_format_el.attrib['sqref'] = 'A1:'+genlabel(img_width+1)+str(img_height)
      cond_format_el.append(etree.fromstring('''
        <cfRule type="colorScale" priority="2">
          <colorScale>
            <cfvo type="formula" val="0"/>
            <cfvo type="formula" val="16777215"/>
            <color rgb="FF000000"/>
            <color rgb="FFFFFFFF"/>
          </colorScale>
        </cfRule>
      '''))
      
      cond_format = etree.tostring(cond_format_el)
  
  # sheet format properties element
  sheet_format = etree.fromstring('<sheetFormatPr customHeight="1" />')
  if subpixels:
    sheet_format.attrib['defaultColWidth'] = '60'
    sheet_format.attrib['defaultRowHeight'] = '20'
  
  # combine the sheet format properties, sheet data, and
  # conditional formatting to create the resultng xml string
  result = etree.tostring(sheet_format) +\
           etree.tostring(root) +\
           cond_format
  # convert result to single line
  result = re.sub(r'\n\s*', '', result)
  
  if subpixels:
    result = etree.parse

def genlabel(n):
  label=''
  while n>0:
    n, r = divmod(n-1, 26)
    label = string.ascii_uppercase[r] + label
  return label

  

run(host='localhost', port=8080, debug=True)

