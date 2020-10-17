This is a Windows application template
---

### Now the program has the following functionality:

1. Empty window when opened
2. Top menu with settings button
3. Settings window with setting the program autorun and the ability to minimize to tray

### **ATTENTION**

For the program to work correctly, you need to create an .exe file using pyinstaller by entering the following command 

`pyinstaller -w -F -i [you icon].ico -n [name program] --add-data="[you icon].ico;." main.py`

and make sure the icon name is specified in the ICON_NAME variable of the main.py file
