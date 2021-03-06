from PyQt4.QtGui import *
from view.uiview import UIView
import sys
import traceback
__author__ = 'Project management group 8, Ghent University 2015'
    
def main():
    try:
        app = QApplication(sys.argv)

        window = UIView()
        window.show()
        app.exec_()
    except:
        print("Unhandled Exception occurred in PMConverter of type: {0}\n".format(sys.exc_info()[0]))
        print("Unhandled Exception value = {0}\n".format(sys.exc_info()[1] if sys.exc_info()[1] is not None else "None"))
        print("Unhandled Exception traceback = {0}\n".format(traceback.format_exc()))
        print("\PMConverter CRASHED\n")

if __name__ == "__main__":
    main()