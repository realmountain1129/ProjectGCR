import errno
import os
import sys

import shutil
from pdf2jpg import pdf2jpg
from wand.image import Image


class convertPdf2Image:
    def __init__(self):
        self.IMAGE_PATH = r"D:\image"
        self.CVT_PATH = r"D:\cvt"
        self.TEMP_PATH = r"D:\cvt_temp"
        self.DESK_PATH = r"D:\desk"
        self.RESULT_PATH = r"D:\result"
        self.CROP_PATH_2 = r"D:\crop_stand"
        self.CROP_PATH_1 = r"D:\crop_buil"

        # jpg 파일 저장할 폴더 & temp 저장할 폴더 생성
        # Hoxy... 그런 이름의 폴더가 있다면... 지워주기...
        try:
            if not (os.path.isdir(self.IMAGE_PATH)):
                os.makedirs(os.path.join(self.IMAGE_PATH))
            elif os.path.isdir(self.IMAGE_PATH):
                shutil.rmtree(self.IMAGE_PATH)
                os.makedirs(os.path.join(self.IMAGE_PATH))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("IMAGE 폴더를 못만들었습니다ㅠㅜ")
                raise

        try:
            if not (os.path.isdir(self.CVT_PATH)):
                os.makedirs(os.path.join(self.CVT_PATH))
            elif os.path.isdir(self.CVT_PATH):
                shutil.rmtree(self.CVT_PATH)
                os.makedirs(os.path.join(self.CVT_PATH))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("CVT 폴더를 못만들었습니다ㅠㅜ")
                raise

        try:
            if not (os.path.isdir(self.TEMP_PATH)):
                os.makedirs(os.path.join(self.TEMP_PATH))
            elif os.path.isdir(self.TEMP_PATH):
                shutil.rmtree(self.TEMP_PATH)
                os.makedirs(os.path.join(self.TEMP_PATH))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("CVT_TEMP 폴더를 못만들었습니다ㅠㅜ")

        try:
            if not (os.path.isdir(self.DESK_PATH)):
                os.makedirs(os.path.join(self.DESK_PATH))
            elif os.path.isdir(self.DESK_PATH):
                shutil.rmtree(self.DESK_PATH)
                os.makedirs(os.path.join(self.DESK_PATH))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("DESK_PATH 폴더를 못만들었습니다ㅠㅜ")

        try:
            if not (os.path.isdir(self.RESULT_PATH)):
                os.makedirs(os.path.join(self.RESULT_PATH))
            elif os.path.isdir(self.RESULT_PATH):
                shutil.rmtree(self.RESULT_PATH)
                os.makedirs(os.path.join(self.RESULT_PATH))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("RESULT_PATH 폴더를 못만들었습니다ㅠㅜ")

        try:
            if not (os.path.isdir(self.CROP_PATH_1)):
                os.makedirs(os.path.join(self.CROP_PATH_1))
            elif os.path.isdir(self.CROP_PATH_1):
                shutil.rmtree(self.CROP_PATH_1)
                os.makedirs(os.path.join(self.CROP_PATH_1))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("CROP_PATH_1 폴더를 못만들었습니다ㅠㅜ")

        try:
            if not (os.path.isdir(self.CROP_PATH_2)):
                os.makedirs(os.path.join(self.CROP_PATH_2))
            elif os.path.isdir(self.CROP_PATH_2):
                shutil.rmtree(self.CROP_PATH_2)
                os.makedirs(os.path.join(self.CROP_PATH_2))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("CROP_PATH_2 폴더를 못만들었습니다ㅠㅜ")

    # path는 시작 위치. output은 전역...

    # 파일 옮기기 (흩어진 폴더들에서 pdf(PDF)만 모아오기)
    # 1. 폴더 안 폴더의 경우
    '''
    def moveFile(self, path, file_type):
        # cnt = len(os.listdir(path))
        for dir_name in os.listdir(path):
            file_name = os.listdir(path + "\\" + dir_name)[0]
            if (file_name[-3:] == file_type.upper()) or (file_name[-3:] == file_type.lower()):
                shutil.copy(path + "\\" + dir_name + "\\" + file_name, self.IMAGE_PATH)
    '''

    # 파일 옮기기
    def moveFile(self, path, file_type):
        # cnt = len(os.listdir(path))
        for file_name in os.listdir(path):
            if (file_name[-3:] == file_type.lower()) or (file_name[-3:] == file_type.upper()):
                # print("경로 확인: ", path)
                shutil.copy(os.path.join(path, file_name), self.IMAGE_PATH)  # path + "\\" + file_name, self.IMAGE_PATH)

    # 파일 변환하기
    # 조건 : 1페이지에 필요한 서류가 있어야 함!!
    def cvtFile(self):
        for file_name in os.listdir(self.IMAGE_PATH):
            pdf2jpg.convert_pdf2jpg(self.IMAGE_PATH + "\\" + file_name, self.TEMP_PATH, dpi=200, pages="0")

    # jpg 파일 한 데 모으기
    def cltImage(self):
        for dir_name in os.listdir(self.TEMP_PATH):
            file_name = os.listdir(self.TEMP_PATH + "\\" + dir_name)[0]
            if file_name[-3:] == "jpg":
                shutil.copy(self.TEMP_PATH + "\\" + dir_name + "\\" + file_name, self.CVT_PATH)

    # 이름 바꾸기
    def chgName(self):
        for file_name in os.listdir(self.CVT_PATH):
            if file_name[-3:] == "jpg":
                os.rename(self.CVT_PATH + "\\" + file_name, self.CVT_PATH + "\\" + file_name[2:11] + ".jpg")

    # 이미지 보정하기
    def deskewImage(self):
        for file_name in os.listdir(self.CVT_PATH):
            with Image(filename=self.CVT_PATH + "\\" + file_name) as img:
                img.deskew(0.4 * img.quantum_range)
                img.type = "grayscale"
                img.save(filename=self.DESK_PATH + "\\" + file_name)

    # 작업에 사용됐던 폴더들 지우기...
    def removeDir(self):
        # 작업에 사용된 폴더들 삭제
        shutil.rmtree(self.TEMP_PATH)
        shutil.rmtree(self.IMAGE_PATH)

    def removeCvtDir(self):
        # 이미지 저장된 폴더 삭제
        shutil.rmtree(self.CVT_PATH)

    def removeDeskDirBuil(self):
        # 보정 이미지 저장된 폴더 삭제
        shutil.rmtree(self.DESK_PATH + "_building")

    def removeDeskDirStand(self):
        # 보정 이미지 저장된 폴더 삭제
        shutil.rmtree(self.DESK_PATH + "_standard")

    def removeCropDir(self):
        # 크롭 이미지 저장된 폴더 삭제
        shutil.rmtree(self.CROP_PATH_1)
        shutil.rmtree(self.CROP_PATH_2)
