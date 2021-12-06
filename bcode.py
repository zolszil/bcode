import pandas as pd
import barcode
from barcode.writer import ImageWriter
from io import BytesIO
from PIL import Image
import os


# required packages: python-barcode, pandas, openpyxl, barcode, Pillow

# ----------------------------------------------
excel_file = 'plates.xlsx'
excel_column = 'CODES'

# these parameters are in pixels:
h_margin = 80
v_margin = 120
h_gap = 200
v_gap = 50

# global variable ------------------------------
filenames = []
# ----------------------------------------------

# reads the first workbook's column specified by name
def read_excel(ef, ec):
    df = pd.read_excel(ef, sheet_name = 0) # reads the first workbook
    codes = df[ec].tolist()
    if len(codes) > 108:
        codes = []
    return codes


# saves barcodes as individual png files,
# codes are obtained in the list
def create_code_images(lst):
    # necessary code class
    C128 = barcode.get_barcode_class('code128')
    szamlalo = 1
    for i in lst:
        code_img = C128(i, writer=ImageWriter())
        filenames.append('vk_' + '0' * (3-len(str(szamlalo))) + str(szamlalo))
        code_img.save(filenames[szamlalo-1])
        szamlalo += 1
    #print(filenames)    


# reduces original size of bacodes and saves them with the same names
def shrink_codes():
    for i in range(len(filenames)):
        im = Image.open(filenames[i]+'.png')
        w, h = im.size
        x1 = 0
        y1 = int(h*0.2)
        x2 = w
        y2 = int(h*0.4)
        y3 = int(h*0.88)
        y4 = int(h*0.95)
        im1 = im.crop((x1, y1, x2, y2))
        im2 = im.crop((x1, y3, x2, y4))

        dst = Image.new('RGB', (w, im1.height + im2.height))
        dst.paste(im1, (0, 0))
        dst.paste(im2, (0, im1.height))
        dst.save(filenames[i]+'.png')


# merges all images to a single file
# with the above defined margins and gaps
def merge_codes():
    # calculation of full size of result image
    im = Image.open(filenames[0]+'.png')
    w, h = im.size
    w_full = 2*h_margin + 4*w + 3*h_gap
    h_full = 2*v_margin + 27*h + 26*v_gap
    dst = Image.new('RGB', (w_full, h_full), (255,255,255))
      
    x = 0 # hor. coordinate of code (0...3)
    y = 0 # ver. coordinate of code (0...26)
    for i in range(len(filenames)):
        im = Image.open(filenames[i]+'.png')
        coor_x = h_margin + x * (w + h_gap)
        coor_y = v_margin + y * (h + v_gap)
        dst.paste(im, (coor_x, coor_y))
        if x == 3:
            x = 0
            y += 1
        else:
            x += 1

    dst.save('barcodes.png')


# delete all temp. image files
def delete_images():
    for i in range(len(filenames)):
        os.remove(filenames[i] + '.png')


if __name__ == "__main__":
    
    print("Excel to barcode-table image generator \n" + 38*"-")
    code_list = read_excel(excel_file, excel_column)
    if len(code_list) == 0:
        print("ERROR: Too many codes in Excel file")
        input()
        quit()
    else:
        print("Number of codes in Excel file: " + str(len(code_list)))
        create_code_images(code_list)
        shrink_codes()
        merge_codes()
        delete_images()
        print("Output image has been created. \n(Press any key...)")
        input()
        
        

    
