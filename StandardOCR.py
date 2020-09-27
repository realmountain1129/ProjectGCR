import cv2
import imutils
from pyimagesearch.transform import four_point_transform
from skimage.filters import threshold_local
import matplotlib.pyplot as plt
import easyocr
import win32com.client as win32
import os
import pandas as pd
import re
import sqlite3 as lite

# OCR
reader_kr = easyocr.Reader(['ko'], gpu='cuda:0')
reader_en = easyocr.Reader(['en'], gpu='cuda:0')

CROP_PATH_2 = r"D:\\crop_stand"


###################################
###  표준설치계약서 OCR을 위한 py  ###
###################################


def auto_crop(path):
    image = cv2.imread(path)
    ratio = 1
    orig = image.copy()
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(gray, 75, 200)

    cnts = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    cnts = imutils.grab_contours(cnts)
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]

    for c in cnts:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)

        if len(approx) == 4:
            screenCnt = approx
            break

    # Find contours of documents
    cv2.drawContours(image, [screenCnt], -1, (0, 255, 0), 2)

    # apply the four point transform to obtain a top-down
    # view of the original image
    warped = four_point_transform(orig, screenCnt.reshape(4, 2) * ratio)

    # convert the warped image to grayscale, then threshold it
    # to give it that 'black and white' paper effect
    warped = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
    T = threshold_local(warped, 11, offset=10, method='gaussian')
    warped = (warped > T).astype("uint8") * 255

    warped = cv2.resize(warped, dsize=(800, 1000), interpolation=cv2.INTER_AREA)
    # show the original and scanned images
    # Apply perspective transform

    return warped


# threshold
def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]


def contours_name(orig, contours, index, df, custnum, db_con):
    for cnt in contours:
        x, y, w, h = 145, 40, 100, 90  # x,y,width,height

        image_crop = orig[y:y + h, x:x + w]
        image = plt.imshow(image_crop, cmap='gray')
        plt.axis('off')
        plt.savefig(os.path.join(CROP_PATH_2, "crop.png"), bbox_inches='tight')

        image_double = cv2.pyrUp(image_crop, dstsize=(w * 2, h * 2), borderType=cv2.BORDER_DEFAULT)

        file1 = open(os.path.join(CROP_PATH_2, "crop.png"), "rb")
        img1 = file1.read()

        # Apply OCR on the cropped image
        text = reader_kr.readtext(image_double)

    try:
        if text[0][1] == df['신청자명'][index]:
            return "신청자명이 일치합니다!"
        else:
            b_crop = lite.Binary(img1)
            db_con.execute("INSERT OR REPLACE INTO name VALUES(?,?,?)", (custnum, df['신청자명'][index], b_crop))
            db_con.commit()
    except IndexError:
        db_con.execute("INSERT OR REPLACE INTO name VALUES(?,?,?)", (custnum, df['신청자명'][index], None))
        db_con.commit()


def contours_address(orig, contours, index, df, custnum, db_con):
    for cnt in contours:
        x, y, w, h = 165, 287, 320, 60  # x,y,width,height

        dst = orig.copy()
        image_crop = orig[y:y + h, x:x + w]
        image = plt.imshow(image_crop, cmap='gray')
        plt.axis('off')
        plt.savefig(os.path.join(CROP_PATH_2, "crop1.png"), bbox_inches='tight')
        file1 = open(os.path.join(CROP_PATH_2, "crop1.png"), "rb")
        img1 = file1.read()
        # cv2.imwrite('test.png', image_crop)
        b_crop = lite.Binary(img1)

        b = image_crop.tostring()

        return b_crop


def contours_address2(orig, contours, index, df, custnum, db_con):
    for cnt in contours:
        x, y, w, h = 330, 45, 470, 60  # x,y,width,height

        dst = orig.copy()
        image_crop = orig[y:y + h, x:x + w]
        image = plt.imshow(image_crop, cmap='gray')
        plt.axis('off')
        plt.savefig(os.path.join(CROP_PATH_2, "crop2.png"), bbox_inches='tight')
        file2 = open(os.path.join(CROP_PATH_2, "crop2.png"), "rb")
        img2 = file2.read()

        b_crop = lite.Binary(img2)
        b = image_crop.tostring()

        image_double = cv2.pyrUp(image_crop, dstsize=(w * 2, h * 2), borderType=cv2.BORDER_DEFAULT)

        # Apply OCR on the cropped image
        text = reader_kr.readtext(image_double)
        # return text[0][1]
        return b_crop


def contours_capacity(orig, contours, index, df, custnum, db_con):
    for cnt in contours:
        x, y, w, h = 680, 220, 90, 50  # x,y,width,height

        dst = orig.copy()
        image_crop = orig[y:y + h, x:x + w]
        image = plt.imshow(image_crop, cmap='gray')
        plt.axis('off')
        plt.savefig(os.path.join(CROP_PATH_2, "crop4.png"), bbox_inches='tight')

        # b_crop = lite.Binary(image_crop)

        image_double = cv2.pyrUp(image_crop, dstsize=(w * 2, h * 2), borderType=cv2.BORDER_DEFAULT)

        # cv2.imwrite('capacity.jpg', image_crop)

        # Apply OCR on the cropped image
        text = reader_en.readtext(image_double)
        r = re.compile("(\d+)(\s|\w)")
        m = r.match(text[0][1])

        # file4 = open(os.path.join(CROP_PATH, "crop4.png"), encoding='utf-16-le')
        # img4 = file4.read()

        try:
            if int(m.groups()[0]) == int(df['실제설치용량'][index]):
                return 'Yes'
            else:
                file4 = open(os.path.join(CROP_PATH_2, "crop4.png"), "rb")
                img4 = file4.read()

                b_crop = lite.Binary(img4)

                # b_crop = lite.Binary(img1)
                db_con.execute("INSERT INTO capacity VALUES(?,?,?)", (custnum, df['실제설치용량'][index], b_crop))
                db_con.commit()

        except AttributeError:
            return '용량이 일치하지 않습니다. 이미지와 비교해주세요! \n'
        except IndexError:
            db_con.execute("INSERT INTO capacity VALUES(?,?,?)", (custnum, df['실제설치용량'][index], None))
            db_con.commit()


# OCR 수행
def standOCR(img_path, img_list, ex_name, db_con):
    df = pd.read_excel(ex_name)
    for img in img_list:
        try:
            num = img.split('.')
            for i in range(0, len(df.index)):
                if df['참여신청번호'][i] == int(num[0]):
                    index = (df.index[i])
                else:
                    continue
            print("auto_Crop 전 : ", os.path.join(img_path, img))
            warped = auto_crop(os.path.join(img_path, img))
            threshold = cv2.threshold(warped, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            contours_name(warped, threshold, index, df, num[0], db_con)

            binary1 = contours_address(warped, threshold, index, df, num[0], db_con)
            binary2 = contours_address2(warped, threshold, index, df, num[0], db_con)

            db_con.execute("INSERT INTO address VALUES(?,?,?,?)", (num[0], df['설치주소'][index], binary1, binary2))
            db_con.commit()

            contours_capacity(warped, threshold, index, df, num[0], db_con)

        except UnboundLocalError:
            print(img + ' Error!\n')
    return
