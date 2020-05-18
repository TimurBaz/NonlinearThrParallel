import multiprocessing
from joblib import Parallel, delayed
import time
from win32com.client import Dispatch
import os.path


def Test():
    try:
        app = Dispatch('OptiSystem.Application')
        filepath = os.getcwd()
        filename = 'Test.osd'
        filename = os.path.join(filepath, filename)
        app.Open(filename)
        doc = app.GetActiveDocument()
        doc.CalculateProject(True, True)
        doc.Save(filename)
    except ValueError:
        print('Hi world!')


Test()
#num_cores = multiprocessing.cpu_count()
#t = time.time()
#
#
#def myfun(i):
#    xl = Dispatch("Excel.Application")
#    xl.Visible = True  # otherwise excel is hidden
#    # newest excel does not accept forward slash in path
#    if i == 0:
#        xl.Workbooks.Open(r'D:\Education\Python\Projects\exam.xls')
#    elif i == 1:
#        xl.Workbooks.Open(r'D:\Education\Python\Projects\exam1.xls')
#    return 1
#
#
#processed_list = Parallel(2)(delayed(myfun)(i) for i in range(2))
#
## do stuff
#elapsed = time.time() - t
#print(elapsed)


# xl = Dispatch("Excel.Application")
# xl.Visible = True # otherwise excel is hidden

# newest excel does not accept forward slsh in path
# wb = xl.Workbooks.Open(r'D:\Education\Python\Projects\exam.xls')
# wb.Close()
# xl.Quit()
