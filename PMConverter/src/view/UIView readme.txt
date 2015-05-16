to convert .ui file to .py:
pyuic4 -o ui_UIView.py UIView.ui

to convert .qrc to .py:
pyrcc4 -py3 -o designer_resource_rc.py designer_resource.qrc
-> change import statement in ui_UIView.py to include the path of its folder: view