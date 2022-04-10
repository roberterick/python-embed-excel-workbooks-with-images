import openpyxl as opx
import os,sys
import shutil

'''
given a directory of excel workbooks
and 2 image files that need to be embedded
scroll through the workbooks and embed the files
'''

WORKBOOKS_DONE=r'C:\path\spreadsheets_done'
WORKBOOKS_DONE_WITH_IMAGES=r'C:\path\spreadsheets_done_with_images'
LOGO_NAME='logo.png'
DISCLOSURE_NAME='conf_horiz10.jpg'


class App:
    def __init__(self):
        self.process_all()

    def process_all(self):
        shortlist=os.listdir(WORKBOOKS_DONE)
        shortlist=filter(lambda x:x.endswith('.xlsx'),shortlist)
        longlist=[os.path.join(WORKBOOKS_DONE,x) for x in shortlist]
        longlist=[os.path.abspath(x) for x in longlist]
        for fpth in longlist:
            self.process_one(fpth)

    def process_one(self,fpth):
        fpth=fpth.replace("\\","/")
        print(fpth)
        basename=os.path.basename(fpth)
        fpthi=os.path.join(WORKBOOKS_DONE_WITH_IMAGES,basename)
        fpthi=fpthi.replace("\\","/")

        shutil.copyfile(fpth,fpthi)
##        f=open(fpthi,'wb')
        wb=opx.load_workbook(fpthi)
        wb.save(fpthi)
        for ws in wb.worksheets:
            img=opx.drawing.image.Image(os.path.join(sys.path[0],LOGO_NAME))
            img.height=74
            img.width=254
            ws.add_image(img,'A1')
            #
            img=opx.drawing.image.Image(os.path.join(sys.path[0],DISCLOSURE_NAME))
            img.height=43
            img.width=655
            ws.add_image(img,'D1')
        wb.save(fpthi)

if __name__=='__main__':
    print sys.path[0]
    app=App()
