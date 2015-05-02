from PyQt4.QtGui import *
from view.uiview import UIView
import sys
    
try:
    app = QApplication(sys.argv)

    window = UIView()
    window.show()
    app.exec_()
except:
    print("Unhandled Exception occurred in PMConverter of type: {0}\n".format(sys.exc_info()[0]))
    print("Unhandled Exception value = {0}\n".format(sys.exc_info()[1] if sys.exc_info()[1] is not None else "None"))
    print("Unhandled Exception traceback = {0}\n".format(sys.exc_info()[2] if sys.exc_info()[2] is not None else "None"))
    print("\PMConverter CRASHED\n")