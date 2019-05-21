# Parses property information that is copy/pasted from assessor websites
# The loaded data can be merged to a PDF with or without adjustments
#
#  Assessor websites are static and have the following layout:
#
#       Address: 145 Python Ave
#       Owner: Bob
#       Style: 2-Story Condo
#       Size: 900 sf
#       Beds: 2
#       ...
#       Sale Date   Sale Price  Sale Terms
#       2017        $50,000     Warranty Deed
#       ...
#
#  Most data is extracted by finding the substring between two strings
#  Address is equal to the string between "Address:" and "Owner:"
#
#  Some data is extracted by finding the first word that follows a string
#  Bed count is equal to the word that follows the string "Beds:"
#
#  Some data is extracted by finding the nth word that follows a string
#  Sale price is equal to the second word that follows the string "Sale Terms"
#
#  Each website is different, so methods are structured like this:
#   GetBath()
#   if Douglas County: return X
#   if Sarpy County: return Y
#   ...
#
#  The program is broken up into five sections:
#   - GUI is defined
#   - Buttons actions are defined
#   - Methods for extracting property data are defined
#   - Methods for calculating adjustments are defined (superior/inferior)
#   - Data is merged to a PDF
#
#  Search ### to find section breaks


import re  # regular expressions for string parsing
import tkinter as tk  # Gui Framework


def init(top, gui, *args, **kwargs):
    global w, top_level, root
    w = gui
    top_level = top
    root = top


def destroy_window():
    # Function which closes the window.
    global top_level
    top_level.destroy()
    top_level = None


def vp_start_gui():
    # Starting point when module is the main routine.
    global val, w, root
    root = tk.Tk()
    top = Toplevel1(root)
    init(root, top)
    root.mainloop()


def create_Toplevel1(root, *args, **kwargs):
    #  Starting point when module is imported by another program.
    rt = root
    w = tk.Toplevel(root)
    top = Toplevel1(w)
    init(w, top, *args, **kwargs)
    return (w, top)


def destroy_Toplevel1():
    global w
    w.destroy()
    w = None


class Toplevel1:  # Only class in this program

    def __init__(self, top=None):

        # this method only defines GUI and creates a few blank strings.

        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9'  # X11 color: 'gray85'
        _ana1color = '#d9d9d9'  # X11 color: 'gray85'
        _ana2color = '#ececec'  # Closest X11 color: 'gray92'
        font9 = "-family {Open Sans} -size 9"

        top.geometry("903x638+568+271")
        top.title("")
        top.configure(background="#d9d9d9")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")

        self.menubar = tk.Menu(top, font=('Open Sans', 9,), bg=_bgcolor
                               , fg=_fgcolor)
        top.configure(menu=self.menubar)

        self.subjectzip = ""
        self.subjectcounty = ""
        self.subjectstate = ""
        self.subjectcity = ""
        self.subjectassvalue = ""
        self.subjectassyear = ""
        self.selection = ""

        self.Label17 = tk.Label(top)
        self.Label17.place(relx=0.12, rely=0.235, height=23, width=61)
        self.Label17.configure(background="#d9d9d9")
        self.Label17.configure(disabledforeground="#a3a3a3")
        self.Label17.configure(foreground="#000000")
        self.Label17.configure(text='''Sale Price''')

        self.Label2 = tk.Label(top)
        self.Label2.place(relx=0.332, rely=0.016, height=23, width=352)
        self.Label2.configure(activebackground="#f9f9f9")
        self.Label2.configure(activeforeground="black")
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(highlightbackground="#d9d9d9")
        self.Label2.configure(highlightcolor="black")
        self.Label2.configure(text='''Choose What to Fill''')

        self.Label1 = tk.Label(top)
        self.Label1.place(relx=0.11, rely=0.0, height=47, width=180)
        self.Label1.configure(activebackground="#f9f9f9")
        self.Label1.configure(activeforeground="black")
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(highlightbackground="#d9d9d9")
        self.Label1.configure(highlightcolor="black")
        self.Label1.configure(text='''Paste Data''')

        self.Label3 = tk.Label(top)
        self.Label3.place(relx=0.106, rely=0.282, height=23, width=74)
        self.Label3.configure(activebackground="#f9f9f9")
        self.Label3.configure(activeforeground="black")
        self.Label3.configure(background="#d9d9d9")
        self.Label3.configure(disabledforeground="#a3a3a3")
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(highlightbackground="#d9d9d9")
        self.Label3.configure(highlightcolor="black")
        self.Label3.configure(text='''Date of Sale''')

        self.Label4 = tk.Label(top)
        self.Label4.place(relx=0.07, rely=0.329, height=23, width=106)
        self.Label4.configure(activebackground="#f9f9f9")
        self.Label4.configure(activeforeground="black")
        self.Label4.configure(background="#d9d9d9")
        self.Label4.configure(disabledforeground="#a3a3a3")
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(highlightbackground="#d9d9d9")
        self.Label4.configure(highlightcolor="black")
        self.Label4.configure(text='''Conditions of Sale''')

        self.Label5 = tk.Label(top)
        self.Label5.place(relx=0.128, rely=0.376, height=23, width=53)
        self.Label5.configure(activebackground="#f9f9f9")
        self.Label5.configure(activeforeground="black")
        self.Label5.configure(background="#d9d9d9")
        self.Label5.configure(disabledforeground="#a3a3a3")
        self.Label5.configure(foreground="#000000")
        self.Label5.configure(highlightbackground="#d9d9d9")
        self.Label5.configure(highlightcolor="black")
        self.Label5.configure(text='''Location''')

        self.Label6 = tk.Label(top)
        self.Label6.place(relx=0.15, rely=0.423, height=23, width=33)
        self.Label6.configure(activebackground="#f9f9f9")
        self.Label6.configure(activeforeground="black")
        self.Label6.configure(background="#d9d9d9")
        self.Label6.configure(disabledforeground="#a3a3a3")
        self.Label6.configure(foreground="#000000")
        self.Label6.configure(highlightbackground="#d9d9d9")
        self.Label6.configure(highlightcolor="black")
        self.Label6.configure(text='''Style''')

        self.Label7 = tk.Label(top)
        self.Label7.place(relx=0.119, rely=0.47, height=23, width=60)
        self.Label7.configure(activebackground="#f9f9f9")
        self.Label7.configure(activeforeground="black")
        self.Label7.configure(background="#d9d9d9")
        self.Label7.configure(disabledforeground="#a3a3a3")
        self.Label7.configure(foreground="#000000")
        self.Label7.configure(highlightbackground="#d9d9d9")
        self.Label7.configure(highlightcolor="black")
        self.Label7.configure(text='''Year Built''')

        self.Label8 = tk.Label(top)
        self.Label8.place(relx=0.121, rely=0.517, height=23, width=59)
        self.Label8.configure(activebackground="#f9f9f9")
        self.Label8.configure(activeforeground="black")
        self.Label8.configure(background="#d9d9d9")
        self.Label8.configure(disabledforeground="#a3a3a3")
        self.Label8.configure(foreground="#000000")
        self.Label8.configure(highlightbackground="#d9d9d9")
        self.Label8.configure(highlightcolor="black")
        self.Label8.configure(text='''Condition''')

        self.Label9 = tk.Label(top)
        self.Label9.place(relx=0.148, rely=0.564, height=23, width=33)
        self.Label9.configure(activebackground="#f9f9f9")
        self.Label9.configure(activeforeground="black")
        self.Label9.configure(background="#d9d9d9")
        self.Label9.configure(disabledforeground="#a3a3a3")
        self.Label9.configure(foreground="#000000")
        self.Label9.configure(highlightbackground="#d9d9d9")
        self.Label9.configure(highlightcolor="black")
        self.Label9.configure(text='''Area''')

        self.Label10 = tk.Label(top)
        self.Label10.place(relx=0.1129, rely=0.611, height=23, width=64)
        self.Label10.configure(activebackground="#f9f9f9")
        self.Label10.configure(activeforeground="black")
        self.Label10.configure(background="#d9d9d9")
        self.Label10.configure(disabledforeground="#a3a3a3")
        self.Label10.configure(foreground="#000000")
        self.Label10.configure(highlightbackground="#d9d9d9")
        self.Label10.configure(highlightcolor="black")
        self.Label10.configure(text='''Bedrooms''')

        self.Label11 = tk.Label(top)
        self.Label11.place(relx=0.109, rely=0.658, height=23, width=68)
        self.Label11.configure(activebackground="#f9f9f9")
        self.Label11.configure(activeforeground="black")
        self.Label11.configure(background="#d9d9d9")
        self.Label11.configure(disabledforeground="#a3a3a3")
        self.Label11.configure(foreground="#000000")
        self.Label11.configure(highlightbackground="#d9d9d9")
        self.Label11.configure(highlightcolor="black")
        self.Label11.configure(text='''Bathrooms''')

        self.Label12 = tk.Label(top)
        self.Label12.place(relx=0.115, rely=0.705, height=23, width=63)
        self.Label12.configure(activebackground="#f9f9f9")
        self.Label12.configure(activeforeground="black")
        self.Label12.configure(background="#d9d9d9")
        self.Label12.configure(disabledforeground="#a3a3a3")
        self.Label12.configure(foreground="#000000")
        self.Label12.configure(highlightbackground="#d9d9d9")
        self.Label12.configure(highlightcolor="black")
        self.Label12.configure(text='''Basement''')

        self.Label13 = tk.Label(top)
        self.Label13.place(relx=0.13, rely=0.752, height=23, width=48)
        self.Label13.configure(activebackground="#f9f9f9")
        self.Label13.configure(activeforeground="black")
        self.Label13.configure(background="#d9d9d9")
        self.Label13.configure(disabledforeground="#a3a3a3")
        self.Label13.configure(foreground="#000000")
        self.Label13.configure(highlightbackground="#d9d9d9")
        self.Label13.configure(highlightcolor="black")
        self.Label13.configure(text='''Garage''')

        self.Label14 = tk.Label(top)
        self.Label14.place(relx=0.076, rely=0.799, height=23, width=97)
        self.Label14.configure(activebackground="#f9f9f9")
        self.Label14.configure(activeforeground="black")
        self.Label14.configure(background="#d9d9d9")
        self.Label14.configure(disabledforeground="#a3a3a3")
        self.Label14.configure(foreground="#000000")
        self.Label14.configure(highlightbackground="#d9d9d9")
        self.Label14.configure(highlightcolor="black")
        self.Label14.configure(text='''Other Amenities''')

        self.Label15 = tk.Label(top)
        self.Label15.place(relx=0.13, rely=0.188, height=23, width=52)
        self.Label15.configure(activebackground="#f9f9f9")
        self.Label15.configure(activeforeground="black")
        self.Label15.configure(background="#d9d9d9")
        self.Label15.configure(disabledforeground="#a3a3a3")
        self.Label15.configure(foreground="#000000")
        self.Label15.configure(highlightbackground="#d9d9d9")
        self.Label15.configure(highlightcolor="black")
        self.Label15.configure(text='''Address''')

        self.Label16 = tk.Label(top)
        self.Label16.place(relx=0.199, rely=0.149, height=23, width=137)
        self.Label16.configure(activebackground="#f9f9f9")
        self.Label16.configure(activeforeground="black")
        self.Label16.configure(background="#d9d9d9")
        self.Label16.configure(disabledforeground="#a3a3a3")
        self.Label16.configure(foreground="#000000")
        self.Label16.configure(highlightbackground="#d9d9d9")
        self.Label16.configure(highlightcolor="black")
        self.Label16.configure(text='''Subject''')

        self.Label16_11 = tk.Label(top)
        self.Label16_11.place(relx=0.365, rely=0.149, height=23, width=137)
        self.Label16_11.configure(activebackground="#f9f9f9")
        self.Label16_11.configure(activeforeground="black")
        self.Label16_11.configure(background="#d9d9d9")
        self.Label16_11.configure(disabledforeground="#a3a3a3")
        self.Label16_11.configure(foreground="#000000")
        self.Label16_11.configure(highlightbackground="#d9d9d9")
        self.Label16_11.configure(highlightcolor="black")
        self.Label16_11.configure(text='''Comparable 1''')

        self.Label16_12 = tk.Label(top)
        self.Label16_12.place(relx=0.532, rely=0.149, height=23, width=137)
        self.Label16_12.configure(activebackground="#f9f9f9")
        self.Label16_12.configure(activeforeground="black")
        self.Label16_12.configure(background="#d9d9d9")
        self.Label16_12.configure(disabledforeground="#a3a3a3")
        self.Label16_12.configure(foreground="#000000")
        self.Label16_12.configure(highlightbackground="#d9d9d9")
        self.Label16_12.configure(highlightcolor="black")
        self.Label16_12.configure(text='''Comparable 2''')

        self.Label16_13 = tk.Label(top)
        self.Label16_13.place(relx=0.698, rely=0.149, height=23, width=137)
        self.Label16_13.configure(activebackground="#f9f9f9")
        self.Label16_13.configure(activeforeground="black")
        self.Label16_13.configure(background="#d9d9d9")
        self.Label16_13.configure(disabledforeground="#a3a3a3")
        self.Label16_13.configure(foreground="#000000")
        self.Label16_13.configure(highlightbackground="#d9d9d9")
        self.Label16_13.configure(highlightcolor="black")
        self.Label16_13.configure(text='''Comparable 3''')

        self.Subjectaddressbox = tk.Entry(top)
        self.Subjectaddressbox.place(relx=0.199, rely=0.188, height=20
                                     , relwidth=0.148)
        self.Subjectaddressbox.configure(background="white")
        self.Subjectaddressbox.configure(disabledforeground="#a3a3a3")
        self.Subjectaddressbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectaddressbox.configure(foreground="#000000")
        self.Subjectaddressbox.configure(highlightbackground="#d9d9d9")
        self.Subjectaddressbox.configure(highlightcolor="black")
        self.Subjectaddressbox.configure(insertbackground="black")
        self.Subjectaddressbox.configure(justify='center')
        self.Subjectaddressbox.configure(selectbackground="#c4c4c4")
        self.Subjectaddressbox.configure(selectforeground="black")

        self.Comp1addressbox = tk.Entry(top)
        self.Comp1addressbox.place(relx=0.365, rely=0.188, height=20
                                   , relwidth=0.148)
        self.Comp1addressbox.configure(background="white")
        self.Comp1addressbox.configure(disabledforeground="#a3a3a3")
        self.Comp1addressbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1addressbox.configure(foreground="#000000")
        self.Comp1addressbox.configure(highlightbackground="#d9d9d9")
        self.Comp1addressbox.configure(highlightcolor="black")
        self.Comp1addressbox.configure(insertbackground="black")
        self.Comp1addressbox.configure(justify='center')
        self.Comp1addressbox.configure(selectbackground="#c4c4c4")
        self.Comp1addressbox.configure(selectforeground="black")

        self.Comp2addressbox = tk.Entry(top)
        self.Comp2addressbox.place(relx=0.532, rely=0.188, height=20
                                   , relwidth=0.148)
        self.Comp2addressbox.configure(background="white")
        self.Comp2addressbox.configure(disabledforeground="#a3a3a3")
        self.Comp2addressbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2addressbox.configure(foreground="#000000")
        self.Comp2addressbox.configure(highlightbackground="#d9d9d9")
        self.Comp2addressbox.configure(highlightcolor="black")
        self.Comp2addressbox.configure(insertbackground="black")
        self.Comp2addressbox.configure(justify='center')
        self.Comp2addressbox.configure(selectbackground="#c4c4c4")
        self.Comp2addressbox.configure(selectforeground="black")

        self.Comp3addressbox = tk.Entry(top)
        self.Comp3addressbox.place(relx=0.698, rely=0.188, height=20
                                   , relwidth=0.148)
        self.Comp3addressbox.configure(background="white")
        self.Comp3addressbox.configure(disabledforeground="#a3a3a3")
        self.Comp3addressbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3addressbox.configure(foreground="#000000")
        self.Comp3addressbox.configure(highlightbackground="#d9d9d9")
        self.Comp3addressbox.configure(highlightcolor="black")
        self.Comp3addressbox.configure(insertbackground="black")
        self.Comp3addressbox.configure(justify='center')
        self.Comp3addressbox.configure(selectbackground="#c4c4c4")
        self.Comp3addressbox.configure(selectforeground="black")

        self.Subjectpricebox = tk.Entry(top)
        self.Subjectpricebox.place(relx=0.199, rely=0.235, height=20
                                   , relwidth=0.148)
        self.Subjectpricebox.configure(background="white")
        self.Subjectpricebox.configure(disabledforeground="#a3a3a3")
        self.Subjectpricebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectpricebox.configure(foreground="#000000")
        self.Subjectpricebox.configure(highlightbackground="#d9d9d9")
        self.Subjectpricebox.configure(highlightcolor="black")
        self.Subjectpricebox.configure(insertbackground="black")
        self.Subjectpricebox.configure(justify='center')
        self.Subjectpricebox.configure(selectbackground="#c4c4c4")
        self.Subjectpricebox.configure(selectforeground="black")

        self.Comp1pricebox = tk.Entry(top)
        self.Comp1pricebox.place(relx=0.365, rely=0.235, height=20
                                 , relwidth=0.148)
        self.Comp1pricebox.configure(background="white")
        self.Comp1pricebox.configure(disabledforeground="#a3a3a3")
        self.Comp1pricebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1pricebox.configure(foreground="#000000")
        self.Comp1pricebox.configure(highlightbackground="#d9d9d9")
        self.Comp1pricebox.configure(highlightcolor="black")
        self.Comp1pricebox.configure(insertbackground="black")
        self.Comp1pricebox.configure(justify='center')
        self.Comp1pricebox.configure(selectbackground="#c4c4c4")
        self.Comp1pricebox.configure(selectforeground="black")

        self.Comp2pricebox = tk.Entry(top)
        self.Comp2pricebox.place(relx=0.532, rely=0.235, height=20
                                 , relwidth=0.148)
        self.Comp2pricebox.configure(background="white")
        self.Comp2pricebox.configure(disabledforeground="#a3a3a3")
        self.Comp2pricebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2pricebox.configure(foreground="#000000")
        self.Comp2pricebox.configure(highlightbackground="#d9d9d9")
        self.Comp2pricebox.configure(highlightcolor="black")
        self.Comp2pricebox.configure(insertbackground="black")
        self.Comp2pricebox.configure(justify='center')
        self.Comp2pricebox.configure(selectbackground="#c4c4c4")
        self.Comp2pricebox.configure(selectforeground="black")

        self.Comp3pricebox = tk.Entry(top)
        self.Comp3pricebox.place(relx=0.698, rely=0.235, height=20
                                 , relwidth=0.148)
        self.Comp3pricebox.configure(background="white")
        self.Comp3pricebox.configure(disabledforeground="#a3a3a3")
        self.Comp3pricebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3pricebox.configure(foreground="#000000")
        self.Comp3pricebox.configure(highlightbackground="#d9d9d9")
        self.Comp3pricebox.configure(highlightcolor="black")
        self.Comp3pricebox.configure(insertbackground="black")
        self.Comp3pricebox.configure(justify='center')
        self.Comp3pricebox.configure(selectbackground="#c4c4c4")
        self.Comp3pricebox.configure(selectforeground="black")

        self.Subjectdatebox = tk.Entry(top)
        self.Subjectdatebox.place(relx=0.199, rely=0.282, height=20
                                  , relwidth=0.148)
        self.Subjectdatebox.configure(background="white")
        self.Subjectdatebox.configure(disabledforeground="#a3a3a3")
        self.Subjectdatebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectdatebox.configure(foreground="#000000")
        self.Subjectdatebox.configure(highlightbackground="#d9d9d9")
        self.Subjectdatebox.configure(highlightcolor="black")
        self.Subjectdatebox.configure(insertbackground="black")
        self.Subjectdatebox.configure(justify='center')
        self.Subjectdatebox.configure(selectbackground="#c4c4c4")
        self.Subjectdatebox.configure(selectforeground="black")

        self.Comp1datebox = tk.Entry(top)
        self.Comp1datebox.place(relx=0.365, rely=0.282, height=20
                                , relwidth=0.148)
        self.Comp1datebox.configure(background="white")
        self.Comp1datebox.configure(disabledforeground="#a3a3a3")
        self.Comp1datebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1datebox.configure(foreground="#000000")
        self.Comp1datebox.configure(highlightbackground="#d9d9d9")
        self.Comp1datebox.configure(highlightcolor="black")
        self.Comp1datebox.configure(insertbackground="black")
        self.Comp1datebox.configure(justify='center')
        self.Comp1datebox.configure(selectbackground="#c4c4c4")
        self.Comp1datebox.configure(selectforeground="black")

        self.Comp2datebox = tk.Entry(top)
        self.Comp2datebox.place(relx=0.532, rely=0.282, height=20
                                , relwidth=0.148)
        self.Comp2datebox.configure(background="white")
        self.Comp2datebox.configure(disabledforeground="#a3a3a3")
        self.Comp2datebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2datebox.configure(foreground="#000000")
        self.Comp2datebox.configure(highlightbackground="#d9d9d9")
        self.Comp2datebox.configure(highlightcolor="black")
        self.Comp2datebox.configure(insertbackground="black")
        self.Comp2datebox.configure(justify='center')
        self.Comp2datebox.configure(selectbackground="#c4c4c4")
        self.Comp2datebox.configure(selectforeground="black")

        self.Comp3datebox = tk.Entry(top)
        self.Comp3datebox.place(relx=0.698, rely=0.282, height=20
                                , relwidth=0.148)
        self.Comp3datebox.configure(background="white")
        self.Comp3datebox.configure(disabledforeground="#a3a3a3")
        self.Comp3datebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3datebox.configure(foreground="#000000")
        self.Comp3datebox.configure(highlightbackground="#d9d9d9")
        self.Comp3datebox.configure(highlightcolor="black")
        self.Comp3datebox.configure(insertbackground="black")
        self.Comp3datebox.configure(justify='center')
        self.Comp3datebox.configure(selectbackground="#c4c4c4")
        self.Comp3datebox.configure(selectforeground="black")

        self.Subjectsaleconbox = tk.Entry(top)
        self.Subjectsaleconbox.place(relx=0.199, rely=0.329, height=20
                                     , relwidth=0.148)
        self.Subjectsaleconbox.configure(background="white")
        self.Subjectsaleconbox.configure(disabledforeground="#a3a3a3")
        self.Subjectsaleconbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectsaleconbox.configure(foreground="#000000")
        self.Subjectsaleconbox.configure(highlightbackground="#d9d9d9")
        self.Subjectsaleconbox.configure(highlightcolor="black")
        self.Subjectsaleconbox.configure(insertbackground="black")
        self.Subjectsaleconbox.configure(justify='center')
        self.Subjectsaleconbox.configure(selectbackground="#c4c4c4")
        self.Subjectsaleconbox.configure(selectforeground="black")

        self.Comp1saleconbox = tk.Entry(top)
        self.Comp1saleconbox.place(relx=0.365, rely=0.329, height=20
                                   , relwidth=0.148)
        self.Comp1saleconbox.configure(background="white")
        self.Comp1saleconbox.configure(disabledforeground="#a3a3a3")
        self.Comp1saleconbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1saleconbox.configure(foreground="#000000")
        self.Comp1saleconbox.configure(highlightbackground="#d9d9d9")
        self.Comp1saleconbox.configure(highlightcolor="black")
        self.Comp1saleconbox.configure(insertbackground="black")
        self.Comp1saleconbox.configure(justify='center')
        self.Comp1saleconbox.configure(selectbackground="#c4c4c4")
        self.Comp1saleconbox.configure(selectforeground="black")

        self.Comp2saleconbox = tk.Entry(top)
        self.Comp2saleconbox.place(relx=0.532, rely=0.329, height=20
                                   , relwidth=0.148)
        self.Comp2saleconbox.configure(background="white")
        self.Comp2saleconbox.configure(disabledforeground="#a3a3a3")
        self.Comp2saleconbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2saleconbox.configure(foreground="#000000")
        self.Comp2saleconbox.configure(highlightbackground="#d9d9d9")
        self.Comp2saleconbox.configure(highlightcolor="black")
        self.Comp2saleconbox.configure(insertbackground="black")
        self.Comp2saleconbox.configure(justify='center')
        self.Comp2saleconbox.configure(selectbackground="#c4c4c4")
        self.Comp2saleconbox.configure(selectforeground="black")

        self.Comp3saleconbox = tk.Entry(top)
        self.Comp3saleconbox.place(relx=0.698, rely=0.329, height=20
                                   , relwidth=0.148)
        self.Comp3saleconbox.configure(background="white")
        self.Comp3saleconbox.configure(disabledforeground="#a3a3a3")
        self.Comp3saleconbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3saleconbox.configure(foreground="#000000")
        self.Comp3saleconbox.configure(highlightbackground="#d9d9d9")
        self.Comp3saleconbox.configure(highlightcolor="black")
        self.Comp3saleconbox.configure(insertbackground="black")
        self.Comp3saleconbox.configure(justify='center')
        self.Comp3saleconbox.configure(selectbackground="#c4c4c4")
        self.Comp3saleconbox.configure(selectforeground="black")

        self.Subjectlocationbox = tk.Entry(top)
        self.Subjectlocationbox.place(relx=0.199, rely=0.376, height=20
                                      , relwidth=0.148)
        self.Subjectlocationbox.configure(background="white")
        self.Subjectlocationbox.configure(disabledforeground="#a3a3a3")
        self.Subjectlocationbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectlocationbox.configure(foreground="#000000")
        self.Subjectlocationbox.configure(highlightbackground="#d9d9d9")
        self.Subjectlocationbox.configure(highlightcolor="black")
        self.Subjectlocationbox.configure(insertbackground="black")
        self.Subjectlocationbox.configure(justify='center')
        self.Subjectlocationbox.configure(selectbackground="#c4c4c4")
        self.Subjectlocationbox.configure(selectforeground="black")

        self.Comp1locationbox = tk.Entry(top)
        self.Comp1locationbox.place(relx=0.365, rely=0.376, height=20
                                    , relwidth=0.148)
        self.Comp1locationbox.configure(background="white")
        self.Comp1locationbox.configure(disabledforeground="#a3a3a3")
        self.Comp1locationbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1locationbox.configure(foreground="#000000")
        self.Comp1locationbox.configure(highlightbackground="#d9d9d9")
        self.Comp1locationbox.configure(highlightcolor="black")
        self.Comp1locationbox.configure(insertbackground="black")
        self.Comp1locationbox.configure(justify='center')
        self.Comp1locationbox.configure(selectbackground="#c4c4c4")
        self.Comp1locationbox.configure(selectforeground="black")

        self.Comp2locationbox = tk.Entry(top)
        self.Comp2locationbox.place(relx=0.532, rely=0.376, height=20
                                    , relwidth=0.148)
        self.Comp2locationbox.configure(background="white")
        self.Comp2locationbox.configure(disabledforeground="#a3a3a3")
        self.Comp2locationbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2locationbox.configure(foreground="#000000")
        self.Comp2locationbox.configure(highlightbackground="#d9d9d9")
        self.Comp2locationbox.configure(highlightcolor="black")
        self.Comp2locationbox.configure(insertbackground="black")
        self.Comp2locationbox.configure(justify='center')
        self.Comp2locationbox.configure(selectbackground="#c4c4c4")
        self.Comp2locationbox.configure(selectforeground="black")

        self.Comp3locationbox = tk.Entry(top)
        self.Comp3locationbox.place(relx=0.698, rely=0.376, height=20
                                    , relwidth=0.148)
        self.Comp3locationbox.configure(background="white")
        self.Comp3locationbox.configure(disabledforeground="#a3a3a3")
        self.Comp3locationbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3locationbox.configure(foreground="#000000")
        self.Comp3locationbox.configure(highlightbackground="#d9d9d9")
        self.Comp3locationbox.configure(highlightcolor="black")
        self.Comp3locationbox.configure(insertbackground="black")
        self.Comp3locationbox.configure(justify='center')
        self.Comp3locationbox.configure(selectbackground="#c4c4c4")
        self.Comp3locationbox.configure(selectforeground="black")

        self.Subjectstylebox = tk.Entry(top)
        self.Subjectstylebox.place(relx=0.199, rely=0.423, height=20
                                   , relwidth=0.148)
        self.Subjectstylebox.configure(background="white")
        self.Subjectstylebox.configure(disabledforeground="#a3a3a3")
        self.Subjectstylebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectstylebox.configure(foreground="#000000")
        self.Subjectstylebox.configure(highlightbackground="#d9d9d9")
        self.Subjectstylebox.configure(highlightcolor="black")
        self.Subjectstylebox.configure(insertbackground="black")
        self.Subjectstylebox.configure(justify='center')
        self.Subjectstylebox.configure(selectbackground="#c4c4c4")
        self.Subjectstylebox.configure(selectforeground="black")

        self.Comp1stylebox = tk.Entry(top)
        self.Comp1stylebox.place(relx=0.365, rely=0.423, height=20
                                 , relwidth=0.148)
        self.Comp1stylebox.configure(background="white")
        self.Comp1stylebox.configure(disabledforeground="#a3a3a3")
        self.Comp1stylebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1stylebox.configure(foreground="#000000")
        self.Comp1stylebox.configure(highlightbackground="#d9d9d9")
        self.Comp1stylebox.configure(highlightcolor="black")
        self.Comp1stylebox.configure(insertbackground="black")
        self.Comp1stylebox.configure(justify='center')
        self.Comp1stylebox.configure(selectbackground="#c4c4c4")
        self.Comp1stylebox.configure(selectforeground="black")

        self.Comp2stylebox = tk.Entry(top)
        self.Comp2stylebox.place(relx=0.532, rely=0.423, height=20
                                 , relwidth=0.148)
        self.Comp2stylebox.configure(background="white")
        self.Comp2stylebox.configure(disabledforeground="#a3a3a3")
        self.Comp2stylebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2stylebox.configure(foreground="#000000")
        self.Comp2stylebox.configure(highlightbackground="#d9d9d9")
        self.Comp2stylebox.configure(highlightcolor="black")
        self.Comp2stylebox.configure(insertbackground="black")
        self.Comp2stylebox.configure(justify='center')
        self.Comp2stylebox.configure(selectbackground="#c4c4c4")
        self.Comp2stylebox.configure(selectforeground="black")

        self.Comp3stylebox = tk.Entry(top)
        self.Comp3stylebox.place(relx=0.698, rely=0.423, height=20
                                 , relwidth=0.148)
        self.Comp3stylebox.configure(background="white")
        self.Comp3stylebox.configure(disabledforeground="#a3a3a3")
        self.Comp3stylebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3stylebox.configure(foreground="#000000")
        self.Comp3stylebox.configure(highlightbackground="#d9d9d9")
        self.Comp3stylebox.configure(highlightcolor="black")
        self.Comp3stylebox.configure(insertbackground="black")
        self.Comp3stylebox.configure(justify='center')
        self.Comp3stylebox.configure(selectbackground="#c4c4c4")
        self.Comp3stylebox.configure(selectforeground="black")

        self.Subjectyearbox = tk.Entry(top)
        self.Subjectyearbox.place(relx=0.199, rely=0.47, height=20
                                  , relwidth=0.148)
        self.Subjectyearbox.configure(background="white")
        self.Subjectyearbox.configure(disabledforeground="#a3a3a3")
        self.Subjectyearbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectyearbox.configure(foreground="#000000")
        self.Subjectyearbox.configure(highlightbackground="#d9d9d9")
        self.Subjectyearbox.configure(highlightcolor="black")
        self.Subjectyearbox.configure(insertbackground="black")
        self.Subjectyearbox.configure(justify='center')
        self.Subjectyearbox.configure(selectbackground="#c4c4c4")
        self.Subjectyearbox.configure(selectforeground="black")

        self.Comp1yearbox = tk.Entry(top)
        self.Comp1yearbox.place(relx=0.365, rely=0.47, height=20, relwidth=0.148)

        self.Comp1yearbox.configure(background="white")
        self.Comp1yearbox.configure(disabledforeground="#a3a3a3")
        self.Comp1yearbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1yearbox.configure(foreground="#000000")
        self.Comp1yearbox.configure(highlightbackground="#d9d9d9")
        self.Comp1yearbox.configure(highlightcolor="black")
        self.Comp1yearbox.configure(insertbackground="black")
        self.Comp1yearbox.configure(justify='center')
        self.Comp1yearbox.configure(selectbackground="#c4c4c4")
        self.Comp1yearbox.configure(selectforeground="black")

        self.Comp2yearbox = tk.Entry(top)
        self.Comp2yearbox.place(relx=0.532, rely=0.47, height=20, relwidth=0.148)

        self.Comp2yearbox.configure(background="white")
        self.Comp2yearbox.configure(disabledforeground="#a3a3a3")
        self.Comp2yearbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2yearbox.configure(foreground="#000000")
        self.Comp2yearbox.configure(highlightbackground="#d9d9d9")
        self.Comp2yearbox.configure(highlightcolor="black")
        self.Comp2yearbox.configure(insertbackground="black")
        self.Comp2yearbox.configure(justify='center')
        self.Comp2yearbox.configure(selectbackground="#c4c4c4")
        self.Comp2yearbox.configure(selectforeground="black")

        self.Comp3yearbox = tk.Entry(top)
        self.Comp3yearbox.place(relx=0.698, rely=0.47, height=20, relwidth=0.148)

        self.Comp3yearbox.configure(background="white")
        self.Comp3yearbox.configure(disabledforeground="#a3a3a3")
        self.Comp3yearbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3yearbox.configure(foreground="#000000")
        self.Comp3yearbox.configure(highlightbackground="#d9d9d9")
        self.Comp3yearbox.configure(highlightcolor="black")
        self.Comp3yearbox.configure(insertbackground="black")
        self.Comp3yearbox.configure(justify='center')
        self.Comp3yearbox.configure(selectbackground="#c4c4c4")
        self.Comp3yearbox.configure(selectforeground="black")

        self.Subjectconditionbox = tk.Entry(top)
        self.Subjectconditionbox.place(relx=0.199, rely=0.517, height=20
                                       , relwidth=0.148)
        self.Subjectconditionbox.configure(background="white")
        self.Subjectconditionbox.configure(disabledforeground="#a3a3a3")
        self.Subjectconditionbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectconditionbox.configure(foreground="#000000")
        self.Subjectconditionbox.configure(highlightbackground="#d9d9d9")
        self.Subjectconditionbox.configure(highlightcolor="black")
        self.Subjectconditionbox.configure(insertbackground="black")
        self.Subjectconditionbox.configure(justify='center')
        self.Subjectconditionbox.configure(selectbackground="#c4c4c4")
        self.Subjectconditionbox.configure(selectforeground="black")

        self.Comp1conditionbox = tk.Entry(top)
        self.Comp1conditionbox.place(relx=0.365, rely=0.517, height=20
                                     , relwidth=0.148)
        self.Comp1conditionbox.configure(background="white")
        self.Comp1conditionbox.configure(disabledforeground="#a3a3a3")
        self.Comp1conditionbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1conditionbox.configure(foreground="#000000")
        self.Comp1conditionbox.configure(highlightbackground="#d9d9d9")
        self.Comp1conditionbox.configure(highlightcolor="black")
        self.Comp1conditionbox.configure(insertbackground="black")
        self.Comp1conditionbox.configure(justify='center')
        self.Comp1conditionbox.configure(selectbackground="#c4c4c4")
        self.Comp1conditionbox.configure(selectforeground="black")

        self.Comp2conditionbox = tk.Entry(top)
        self.Comp2conditionbox.place(relx=0.532, rely=0.517, height=20
                                     , relwidth=0.148)
        self.Comp2conditionbox.configure(background="white")
        self.Comp2conditionbox.configure(disabledforeground="#a3a3a3")
        self.Comp2conditionbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2conditionbox.configure(foreground="#000000")
        self.Comp2conditionbox.configure(highlightbackground="#d9d9d9")
        self.Comp2conditionbox.configure(highlightcolor="black")
        self.Comp2conditionbox.configure(insertbackground="black")
        self.Comp2conditionbox.configure(justify='center')
        self.Comp2conditionbox.configure(selectbackground="#c4c4c4")
        self.Comp2conditionbox.configure(selectforeground="black")

        self.Comp3conditionbox = tk.Entry(top)
        self.Comp3conditionbox.place(relx=0.698, rely=0.517, height=20
                                     , relwidth=0.148)
        self.Comp3conditionbox.configure(background="white")
        self.Comp3conditionbox.configure(disabledforeground="#a3a3a3")
        self.Comp3conditionbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3conditionbox.configure(foreground="#000000")
        self.Comp3conditionbox.configure(highlightbackground="#d9d9d9")
        self.Comp3conditionbox.configure(highlightcolor="black")
        self.Comp3conditionbox.configure(insertbackground="black")
        self.Comp3conditionbox.configure(justify='center')
        self.Comp3conditionbox.configure(selectbackground="#c4c4c4")
        self.Comp3conditionbox.configure(selectforeground="black")

        self.Subjectsizebox = tk.Entry(top)
        self.Subjectsizebox.place(relx=0.199, rely=0.564, height=20
                                  , relwidth=0.148)
        self.Subjectsizebox.configure(background="white")
        self.Subjectsizebox.configure(disabledforeground="#a3a3a3")
        self.Subjectsizebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectsizebox.configure(foreground="#000000")
        self.Subjectsizebox.configure(highlightbackground="#d9d9d9")
        self.Subjectsizebox.configure(highlightcolor="black")
        self.Subjectsizebox.configure(insertbackground="black")
        self.Subjectsizebox.configure(justify='center')
        self.Subjectsizebox.configure(selectbackground="#c4c4c4")
        self.Subjectsizebox.configure(selectforeground="black")

        self.Comp1sizebox = tk.Entry(top)
        self.Comp1sizebox.place(relx=0.365, rely=0.564, height=20
                                , relwidth=0.148)
        self.Comp1sizebox.configure(background="white")
        self.Comp1sizebox.configure(disabledforeground="#a3a3a3")
        self.Comp1sizebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1sizebox.configure(foreground="#000000")
        self.Comp1sizebox.configure(highlightbackground="#d9d9d9")
        self.Comp1sizebox.configure(highlightcolor="black")
        self.Comp1sizebox.configure(insertbackground="black")
        self.Comp1sizebox.configure(justify='center')
        self.Comp1sizebox.configure(selectbackground="#c4c4c4")
        self.Comp1sizebox.configure(selectforeground="black")

        self.Comp2sizebox = tk.Entry(top)
        self.Comp2sizebox.place(relx=0.532, rely=0.564, height=20
                                , relwidth=0.148)
        self.Comp2sizebox.configure(background="white")
        self.Comp2sizebox.configure(disabledforeground="#a3a3a3")
        self.Comp2sizebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2sizebox.configure(foreground="#000000")
        self.Comp2sizebox.configure(highlightbackground="#d9d9d9")
        self.Comp2sizebox.configure(highlightcolor="black")
        self.Comp2sizebox.configure(insertbackground="black")
        self.Comp2sizebox.configure(justify='center')
        self.Comp2sizebox.configure(selectbackground="#c4c4c4")
        self.Comp2sizebox.configure(selectforeground="black")

        self.Comp3sizebox = tk.Entry(top)
        self.Comp3sizebox.place(relx=0.698, rely=0.564, height=20
                                , relwidth=0.148)
        self.Comp3sizebox.configure(background="white")
        self.Comp3sizebox.configure(disabledforeground="#a3a3a3")
        self.Comp3sizebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3sizebox.configure(foreground="#000000")
        self.Comp3sizebox.configure(highlightbackground="#d9d9d9")
        self.Comp3sizebox.configure(highlightcolor="black")
        self.Comp3sizebox.configure(insertbackground="black")
        self.Comp3sizebox.configure(justify='center')
        self.Comp3sizebox.configure(selectbackground="#c4c4c4")
        self.Comp3sizebox.configure(selectforeground="black")

        self.Subjectbedbox = tk.Entry(top)
        self.Subjectbedbox.place(relx=0.199, rely=0.611, height=20
                                 , relwidth=0.148)
        self.Subjectbedbox.configure(background="white")
        self.Subjectbedbox.configure(disabledforeground="#a3a3a3")
        self.Subjectbedbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectbedbox.configure(foreground="#000000")
        self.Subjectbedbox.configure(highlightbackground="#d9d9d9")
        self.Subjectbedbox.configure(highlightcolor="black")
        self.Subjectbedbox.configure(insertbackground="black")
        self.Subjectbedbox.configure(justify='center')
        self.Subjectbedbox.configure(selectbackground="#c4c4c4")
        self.Subjectbedbox.configure(selectforeground="black")

        self.Comp1bedbox = tk.Entry(top)
        self.Comp1bedbox.place(relx=0.365, rely=0.611, height=20, relwidth=0.148)

        self.Comp1bedbox.configure(background="white")
        self.Comp1bedbox.configure(disabledforeground="#a3a3a3")
        self.Comp1bedbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1bedbox.configure(foreground="#000000")
        self.Comp1bedbox.configure(highlightbackground="#d9d9d9")
        self.Comp1bedbox.configure(highlightcolor="black")
        self.Comp1bedbox.configure(insertbackground="black")
        self.Comp1bedbox.configure(justify='center')
        self.Comp1bedbox.configure(selectbackground="#c4c4c4")
        self.Comp1bedbox.configure(selectforeground="black")

        self.Comp2bedbox = tk.Entry(top)
        self.Comp2bedbox.place(relx=0.532, rely=0.611, height=20, relwidth=0.148)

        self.Comp2bedbox.configure(background="white")
        self.Comp2bedbox.configure(disabledforeground="#a3a3a3")
        self.Comp2bedbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2bedbox.configure(foreground="#000000")
        self.Comp2bedbox.configure(highlightbackground="#d9d9d9")
        self.Comp2bedbox.configure(highlightcolor="black")
        self.Comp2bedbox.configure(insertbackground="black")
        self.Comp2bedbox.configure(justify='center')
        self.Comp2bedbox.configure(selectbackground="#c4c4c4")
        self.Comp2bedbox.configure(selectforeground="black")

        self.Comp3bedbox = tk.Entry(top)
        self.Comp3bedbox.place(relx=0.698, rely=0.611, height=20, relwidth=0.148)

        self.Comp3bedbox.configure(background="white")
        self.Comp3bedbox.configure(disabledforeground="#a3a3a3")
        self.Comp3bedbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3bedbox.configure(foreground="#000000")
        self.Comp3bedbox.configure(highlightbackground="#d9d9d9")
        self.Comp3bedbox.configure(highlightcolor="black")
        self.Comp3bedbox.configure(insertbackground="black")
        self.Comp3bedbox.configure(justify='center')
        self.Comp3bedbox.configure(selectbackground="#c4c4c4")
        self.Comp3bedbox.configure(selectforeground="black")

        self.Subjectbathbox = tk.Entry(top)
        self.Subjectbathbox.place(relx=0.199, rely=0.658, height=20
                                  , relwidth=0.148)
        self.Subjectbathbox.configure(background="white")
        self.Subjectbathbox.configure(disabledforeground="#a3a3a3")
        self.Subjectbathbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectbathbox.configure(foreground="#000000")
        self.Subjectbathbox.configure(highlightbackground="#d9d9d9")
        self.Subjectbathbox.configure(highlightcolor="black")
        self.Subjectbathbox.configure(insertbackground="black")
        self.Subjectbathbox.configure(justify='center')
        self.Subjectbathbox.configure(selectbackground="#c4c4c4")
        self.Subjectbathbox.configure(selectforeground="black")

        self.Comp1bathbox = tk.Entry(top)
        self.Comp1bathbox.place(relx=0.365, rely=0.658, height=20
                                , relwidth=0.148)
        self.Comp1bathbox.configure(background="white")
        self.Comp1bathbox.configure(disabledforeground="#a3a3a3")
        self.Comp1bathbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1bathbox.configure(foreground="#000000")
        self.Comp1bathbox.configure(highlightbackground="#d9d9d9")
        self.Comp1bathbox.configure(highlightcolor="black")
        self.Comp1bathbox.configure(insertbackground="black")
        self.Comp1bathbox.configure(justify='center')
        self.Comp1bathbox.configure(selectbackground="#c4c4c4")
        self.Comp1bathbox.configure(selectforeground="black")

        self.Comp2bathbox = tk.Entry(top)
        self.Comp2bathbox.place(relx=0.532, rely=0.658, height=20
                                , relwidth=0.148)
        self.Comp2bathbox.configure(background="white")
        self.Comp2bathbox.configure(disabledforeground="#a3a3a3")
        self.Comp2bathbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2bathbox.configure(foreground="#000000")
        self.Comp2bathbox.configure(highlightbackground="#d9d9d9")
        self.Comp2bathbox.configure(highlightcolor="black")
        self.Comp2bathbox.configure(insertbackground="black")
        self.Comp2bathbox.configure(justify='center')
        self.Comp2bathbox.configure(selectbackground="#c4c4c4")
        self.Comp2bathbox.configure(selectforeground="black")

        self.Comp3bathbox = tk.Entry(top)
        self.Comp3bathbox.place(relx=0.698, rely=0.658, height=20
                                , relwidth=0.148)
        self.Comp3bathbox.configure(background="white")
        self.Comp3bathbox.configure(disabledforeground="#a3a3a3")
        self.Comp3bathbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3bathbox.configure(foreground="#000000")
        self.Comp3bathbox.configure(highlightbackground="#d9d9d9")
        self.Comp3bathbox.configure(highlightcolor="black")
        self.Comp3bathbox.configure(insertbackground="black")
        self.Comp3bathbox.configure(justify='center')
        self.Comp3bathbox.configure(selectbackground="#c4c4c4")
        self.Comp3bathbox.configure(selectforeground="black")

        self.Subjectbasebox = tk.Entry(top)
        self.Subjectbasebox.place(relx=0.199, rely=0.705, height=20
                                  , relwidth=0.148)
        self.Subjectbasebox.configure(background="white")
        self.Subjectbasebox.configure(disabledforeground="#a3a3a3")
        self.Subjectbasebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectbasebox.configure(foreground="#000000")
        self.Subjectbasebox.configure(highlightbackground="#d9d9d9")
        self.Subjectbasebox.configure(highlightcolor="black")
        self.Subjectbasebox.configure(insertbackground="black")
        self.Subjectbasebox.configure(justify='center')
        self.Subjectbasebox.configure(selectbackground="#c4c4c4")
        self.Subjectbasebox.configure(selectforeground="black")

        self.Comp1basebox = tk.Entry(top)
        self.Comp1basebox.place(relx=0.365, rely=0.705, height=20
                                , relwidth=0.148)
        self.Comp1basebox.configure(background="white")
        self.Comp1basebox.configure(disabledforeground="#a3a3a3")
        self.Comp1basebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1basebox.configure(foreground="#000000")
        self.Comp1basebox.configure(highlightbackground="#d9d9d9")
        self.Comp1basebox.configure(highlightcolor="black")
        self.Comp1basebox.configure(insertbackground="black")
        self.Comp1basebox.configure(justify='center')
        self.Comp1basebox.configure(selectbackground="#c4c4c4")
        self.Comp1basebox.configure(selectforeground="black")

        self.Comp2basebox = tk.Entry(top)
        self.Comp2basebox.place(relx=0.532, rely=0.705, height=20
                                , relwidth=0.148)
        self.Comp2basebox.configure(background="white")
        self.Comp2basebox.configure(disabledforeground="#a3a3a3")
        self.Comp2basebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2basebox.configure(foreground="#000000")
        self.Comp2basebox.configure(highlightbackground="#d9d9d9")
        self.Comp2basebox.configure(highlightcolor="black")
        self.Comp2basebox.configure(insertbackground="black")
        self.Comp2basebox.configure(justify='center')
        self.Comp2basebox.configure(selectbackground="#c4c4c4")
        self.Comp2basebox.configure(selectforeground="black")

        self.Comp3basebox = tk.Entry(top)
        self.Comp3basebox.place(relx=0.698, rely=0.705, height=20
                                , relwidth=0.148)
        self.Comp3basebox.configure(background="white")
        self.Comp3basebox.configure(disabledforeground="#a3a3a3")
        self.Comp3basebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3basebox.configure(foreground="#000000")
        self.Comp3basebox.configure(highlightbackground="#d9d9d9")
        self.Comp3basebox.configure(highlightcolor="black")
        self.Comp3basebox.configure(insertbackground="black")
        self.Comp3basebox.configure(justify='center')
        self.Comp3basebox.configure(selectbackground="#c4c4c4")
        self.Comp3basebox.configure(selectforeground="black")

        self.Subjectgaragebox = tk.Entry(top)
        self.Subjectgaragebox.place(relx=0.199, rely=0.752, height=20
                                    , relwidth=0.148)
        self.Subjectgaragebox.configure(background="white")
        self.Subjectgaragebox.configure(disabledforeground="#a3a3a3")
        self.Subjectgaragebox.configure(font="-family {Helvetica} -size 8")
        self.Subjectgaragebox.configure(foreground="#000000")
        self.Subjectgaragebox.configure(highlightbackground="#d9d9d9")
        self.Subjectgaragebox.configure(highlightcolor="black")
        self.Subjectgaragebox.configure(insertbackground="black")
        self.Subjectgaragebox.configure(justify='center')
        self.Subjectgaragebox.configure(selectbackground="#c4c4c4")
        self.Subjectgaragebox.configure(selectforeground="black")

        self.Comp1garagebox = tk.Entry(top)
        self.Comp1garagebox.place(relx=0.365, rely=0.752, height=20
                                  , relwidth=0.148)
        self.Comp1garagebox.configure(background="white")
        self.Comp1garagebox.configure(disabledforeground="#a3a3a3")
        self.Comp1garagebox.configure(font="-family {Helvetica} -size 8")
        self.Comp1garagebox.configure(foreground="#000000")
        self.Comp1garagebox.configure(highlightbackground="#d9d9d9")
        self.Comp1garagebox.configure(highlightcolor="black")
        self.Comp1garagebox.configure(insertbackground="black")
        self.Comp1garagebox.configure(justify='center')
        self.Comp1garagebox.configure(selectbackground="#c4c4c4")
        self.Comp1garagebox.configure(selectforeground="black")

        self.Comp2garagebox = tk.Entry(top)
        self.Comp2garagebox.place(relx=0.532, rely=0.752, height=20
                                  , relwidth=0.148)
        self.Comp2garagebox.configure(background="white")
        self.Comp2garagebox.configure(disabledforeground="#a3a3a3")
        self.Comp2garagebox.configure(font="-family {Helvetica} -size 8")
        self.Comp2garagebox.configure(foreground="#000000")
        self.Comp2garagebox.configure(highlightbackground="#d9d9d9")
        self.Comp2garagebox.configure(highlightcolor="black")
        self.Comp2garagebox.configure(insertbackground="black")
        self.Comp2garagebox.configure(justify='center')
        self.Comp2garagebox.configure(selectbackground="#c4c4c4")
        self.Comp2garagebox.configure(selectforeground="black")

        self.Comp3garagebox = tk.Entry(top)
        self.Comp3garagebox.place(relx=0.698, rely=0.752, height=20
                                  , relwidth=0.148)
        self.Comp3garagebox.configure(background="white")
        self.Comp3garagebox.configure(disabledforeground="#a3a3a3")
        self.Comp3garagebox.configure(font="-family {Helvetica} -size 8")
        self.Comp3garagebox.configure(foreground="#000000")
        self.Comp3garagebox.configure(highlightbackground="#d9d9d9")
        self.Comp3garagebox.configure(highlightcolor="black")
        self.Comp3garagebox.configure(insertbackground="black")
        self.Comp3garagebox.configure(justify='center')
        self.Comp3garagebox.configure(selectbackground="#c4c4c4")
        self.Comp3garagebox.configure(selectforeground="black")

        self.Subjectotherbox = tk.Entry(top)
        self.Subjectotherbox.place(relx=0.199, rely=0.799, height=20
                                   , relwidth=0.148)
        self.Subjectotherbox.configure(background="white")
        self.Subjectotherbox.configure(disabledforeground="#a3a3a3")
        self.Subjectotherbox.configure(font="-family {Helvetica} -size 8")
        self.Subjectotherbox.configure(foreground="#000000")
        self.Subjectotherbox.configure(highlightbackground="#d9d9d9")
        self.Subjectotherbox.configure(highlightcolor="black")
        self.Subjectotherbox.configure(insertbackground="black")
        self.Subjectotherbox.configure(justify='center')
        self.Subjectotherbox.configure(selectbackground="#c4c4c4")
        self.Subjectotherbox.configure(selectforeground="black")

        self.Comp1otherbox = tk.Entry(top)
        self.Comp1otherbox.place(relx=0.365, rely=0.799, height=20
                                 , relwidth=0.148)
        self.Comp1otherbox.configure(background="white")
        self.Comp1otherbox.configure(disabledforeground="#a3a3a3")
        self.Comp1otherbox.configure(font="-family {Helvetica} -size 8")
        self.Comp1otherbox.configure(foreground="#000000")
        self.Comp1otherbox.configure(highlightbackground="#d9d9d9")
        self.Comp1otherbox.configure(highlightcolor="black")
        self.Comp1otherbox.configure(insertbackground="black")
        self.Comp1otherbox.configure(justify='center')
        self.Comp1otherbox.configure(selectbackground="#c4c4c4")
        self.Comp1otherbox.configure(selectforeground="black")

        self.Comp2otherbox = tk.Entry(top)
        self.Comp2otherbox.place(relx=0.532, rely=0.799, height=20
                                 , relwidth=0.148)
        self.Comp2otherbox.configure(background="white")
        self.Comp2otherbox.configure(disabledforeground="#a3a3a3")
        self.Comp2otherbox.configure(font="-family {Helvetica} -size 8")
        self.Comp2otherbox.configure(foreground="#000000")
        self.Comp2otherbox.configure(highlightbackground="#d9d9d9")
        self.Comp2otherbox.configure(highlightcolor="black")
        self.Comp2otherbox.configure(insertbackground="black")
        self.Comp2otherbox.configure(justify='center')
        self.Comp2otherbox.configure(selectbackground="#c4c4c4")
        self.Comp2otherbox.configure(selectforeground="black")

        self.Comp3otherbox = tk.Entry(top)
        self.Comp3otherbox.place(relx=0.698, rely=0.799, height=20
                                 , relwidth=0.148)
        self.Comp3otherbox.configure(background="white")
        self.Comp3otherbox.configure(disabledforeground="#a3a3a3")
        self.Comp3otherbox.configure(font="-family {Helvetica} -size 8")
        self.Comp3otherbox.configure(foreground="#000000")
        self.Comp3otherbox.configure(highlightbackground="#d9d9d9")
        self.Comp3otherbox.configure(highlightcolor="black")
        self.Comp3otherbox.configure(insertbackground="black")
        self.Comp3otherbox.configure(justify='center')
        self.Comp3otherbox.configure(selectbackground="#c4c4c4")
        self.Comp3otherbox.configure(selectforeground="black")

        self.adjustmentcheck = tk.Checkbutton(top)
        self.adjustmentcheck.configure(text="Adjustments")
        self.adjustmentcheck.configure(background="#d9d9d9")
        self.adjustmentcheck.place(relx=.0875, rely=.89)
        self.adjustmentchecker = tk.IntVar()
        self.adjustmentchecker.set(1)
        self.adjustmentcheck.configure(variable=self.adjustmentchecker)

        self.GLAlab = tk.Label(top)
        self.GLAlab.place(relx=0.01, rely=0.94, height=23, width=42)
        self.GLAlab.configure(background="#d9d9d9")
        self.GLAlab.configure(disabledforeground="#a3a3a3")
        self.GLAlab.configure(foreground="#000000")
        self.GLAlab.configure(text='''GLA %''')

        self.Lotlab = tk.Label(top)
        self.Lotlab.place(relx=0.09, rely=0.94, height=23, width=36)
        self.Lotlab.configure(background="#d9d9d9")
        self.Lotlab.configure(disabledforeground="#a3a3a3")
        self.Lotlab.configure(foreground="#000000")
        self.Lotlab.configure(text='''Lot %''')

        self.GLApercent = tk.Entry(top)
        self.GLApercent.place(relx=0.06, rely=0.94, height=20, width=25)
        self.GLApercent.configure(background="white")
        self.GLApercent.configure(disabledforeground="#a3a3a3")
        self.GLApercent.configure(font="-family {Courier New} -size 10")
        self.GLApercent.configure(foreground="#000000")
        self.GLApercent.configure(insertbackground="black")
        self.GLApercent.configure(width=44)
        self.GLApercent.insert('end', '10')

        self.Lotpercent = tk.Entry(top)
        self.Lotpercent.place(relx=0.13, rely=0.94, height=20, width=25)
        self.Lotpercent.configure(background="white")
        self.Lotpercent.configure(disabledforeground="#a3a3a3")
        self.Lotpercent.configure(font="-family {Courier New} -size 10")
        self.Lotpercent.configure(foreground="#000000")
        self.Lotpercent.configure(highlightbackground="#d9d9d9")
        self.Lotpercent.configure(highlightcolor="black")
        self.Lotpercent.configure(insertbackground="black")
        self.Lotpercent.configure(selectbackground="#c4c4c4")
        self.Lotpercent.configure(selectforeground="black")
        self.Lotpercent.configure(width=44)
        self.Lotpercent.insert('end', '30')

        self.Yearlabb = tk.Label(top)
        self.Yearlabb.place(relx=0.16, rely=0.94, height=23, width=36)
        self.Yearlabb.configure(background="#d9d9d9")
        self.Yearlabb.configure(disabledforeground="#a3a3a3")
        self.Yearlabb.configure(foreground="#000000")
        self.Yearlabb.configure(text='''Year''')

        self.Yeardif = tk.Entry(top)
        self.Yeardif.place(relx=0.2, rely=0.94, height=20, width=25)
        self.Yeardif.configure(background="white")
        self.Yeardif.configure(disabledforeground="#a3a3a3")
        self.Yeardif.configure(font="-family {Courier New} -size 10")
        self.Yeardif.configure(foreground="#000000")
        self.Yeardif.configure(insertbackground="black")
        self.Yeardif.configure(width=44)
        self.Yeardif.insert('end', '10')

        self.Bedcheck = tk.Checkbutton(top)
        self.Bedcheck.configure(text="Beds")
        self.Bedcheck.configure(background="#d9d9d9")
        self.Bedcheck.place(relx=.23, rely=.935)
        self.Bedchecker = tk.IntVar()
        self.Bedchecker.set(1)
        self.Bedcheck.configure(variable=self.Bedchecker)

        self.Createpdfbutton = tk.Button(top)
        self.Createpdfbutton.place(relx=0.748, rely=0.039, height=50, width=90)
        self.Createpdfbutton.configure(activebackground="#ececec")
        self.Createpdfbutton.configure(activeforeground="#000000")
        self.Createpdfbutton.configure(background="#d9d9d9")
        self.Createpdfbutton.configure(disabledforeground="#a3a3a3")
        self.Createpdfbutton.configure(foreground="#000000")
        self.Createpdfbutton.configure(highlightbackground="#d9d9d9")
        self.Createpdfbutton.configure(highlightcolor="black")
        self.Createpdfbutton.configure(pady="0")
        self.Createpdfbutton.configure(text='''Create PDF''')
        self.Createpdfbutton.configure(command=self.CreatePDF)

        self.Clearsubjectbutton = tk.Button(top)
        self.Clearsubjectbutton.place(relx=0.249, rely=0.854, height=28
                                      , width=41)
        self.Clearsubjectbutton.configure(activebackground="#ececec")
        self.Clearsubjectbutton.configure(activeforeground="#000000")
        self.Clearsubjectbutton.configure(background="#d9d9d9")
        self.Clearsubjectbutton.configure(disabledforeground="#a3a3a3")
        self.Clearsubjectbutton.configure(foreground="#000000")
        self.Clearsubjectbutton.configure(highlightbackground="#d9d9d9")
        self.Clearsubjectbutton.configure(highlightcolor="black")
        self.Clearsubjectbutton.configure(pady="0")
        self.Clearsubjectbutton.configure(text='''Clear''')
        self.Clearsubjectbutton.configure(command=self.Clearsubject)

        self.Clearcomp1button = tk.Button(top)
        self.Clearcomp1button.place(relx=0.41, rely=0.854, height=28, width=41)
        self.Clearcomp1button.configure(activebackground="#ececec")
        self.Clearcomp1button.configure(activeforeground="#000000")
        self.Clearcomp1button.configure(background="#d9d9d9")
        self.Clearcomp1button.configure(disabledforeground="#a3a3a3")
        self.Clearcomp1button.configure(foreground="#000000")
        self.Clearcomp1button.configure(highlightbackground="#d9d9d9")
        self.Clearcomp1button.configure(highlightcolor="black")
        self.Clearcomp1button.configure(pady="0")
        self.Clearcomp1button.configure(text='''Clear''')
        self.Clearcomp1button.configure(command=self.Clearcomp1)

        self.Clearcomp2button = tk.Button(top)
        self.Clearcomp2button.place(relx=0.581, rely=0.854, height=28, width=41)
        self.Clearcomp2button.configure(activebackground="#ececec")
        self.Clearcomp2button.configure(activeforeground="#000000")
        self.Clearcomp2button.configure(background="#d9d9d9")
        self.Clearcomp2button.configure(disabledforeground="#a3a3a3")
        self.Clearcomp2button.configure(foreground="#000000")
        self.Clearcomp2button.configure(highlightbackground="#d9d9d9")
        self.Clearcomp2button.configure(highlightcolor="black")
        self.Clearcomp2button.configure(pady="0")
        self.Clearcomp2button.configure(text='''Clear''')
        self.Clearcomp2button.configure(command=self.Clearcomp2)

        self.Clearcomp3button = tk.Button(top)
        self.Clearcomp3button.place(relx=0.748, rely=0.854, height=28, width=41)
        self.Clearcomp3button.configure(activebackground="#ececec")
        self.Clearcomp3button.configure(activeforeground="#000000")
        self.Clearcomp3button.configure(background="#d9d9d9")
        self.Clearcomp3button.configure(disabledforeground="#a3a3a3")
        self.Clearcomp3button.configure(foreground="#000000")
        self.Clearcomp3button.configure(highlightbackground="#d9d9d9")
        self.Clearcomp3button.configure(highlightcolor="black")
        self.Clearcomp3button.configure(pady="0")
        self.Clearcomp3button.configure(text='''Clear''')
        self.Clearcomp3button.configure(command=self.Clearcomp3)

        self.Assessordata = tk.Text(top)
        self.Assessordata.place(relx=0.138, rely=0.047, relheight=0.066
                                , relwidth=0.148)
        self.Assessordata.configure(background="white")
        self.Assessordata.configure(font="-family {Open Sans} -size 9")
        self.Assessordata.configure(foreground="black")
        self.Assessordata.configure(highlightbackground="#d9d9d9")
        self.Assessordata.configure(highlightcolor="black")
        self.Assessordata.configure(insertbackground="black")
        self.Assessordata.configure(selectbackground="#c4c4c4")
        self.Assessordata.configure(selectforeground="black")
        self.Assessordata.configure(width=134)
        self.Assessordata.configure(wrap="word")

        self.Fillcomp1button = tk.Button(top)
        self.Fillcomp1button.place(relx=0.432, rely=0.063, height=28, width=78)
        self.Fillcomp1button.configure(activebackground="#ececec")
        self.Fillcomp1button.configure(activeforeground="#000000")
        self.Fillcomp1button.configure(background="#d9d9d9")
        self.Fillcomp1button.configure(disabledforeground="#a3a3a3")
        self.Fillcomp1button.configure(foreground="#000000")
        self.Fillcomp1button.configure(highlightbackground="#d9d9d9")
        self.Fillcomp1button.configure(highlightcolor="black")
        self.Fillcomp1button.configure(pady="0")
        self.Fillcomp1button.configure(text='''Comp 1''')
        self.Fillcomp1button.configure(command=self.FillComp1)

        self.Clearassessor = tk.Button(top)
        self.Clearassessor.place(relx=0.20, rely=0.111, height=10, width=20)
        self.Clearassessor.configure(font=('arial', 8))
        self.Clearassessor.configure(activebackground="#ececec")
        self.Clearassessor.configure(activeforeground="#000000")
        self.Clearassessor.configure(background="#d9d9d9")
        self.Clearassessor.configure(disabledforeground="#a3a3a3")
        self.Clearassessor.configure(foreground="#000000")
        self.Clearassessor.configure(highlightbackground="#d9d9d9")
        self.Clearassessor.configure(highlightcolor="black")
        self.Clearassessor.configure(pady="0")
        self.Clearassessor.configure(text='''X''')
        self.Clearassessor.configure(command=self.Clearassdata)

        self.Fillcomp2button = tk.Button(top)
        self.Fillcomp2button.place(relx=0.532, rely=0.063, height=28, width=78)
        self.Fillcomp2button.configure(activebackground="#ececec")
        self.Fillcomp2button.configure(activeforeground="#000000")
        self.Fillcomp2button.configure(background="#d9d9d9")
        self.Fillcomp2button.configure(disabledforeground="#a3a3a3")
        self.Fillcomp2button.configure(foreground="#000000")
        self.Fillcomp2button.configure(highlightbackground="#d9d9d9")
        self.Fillcomp2button.configure(highlightcolor="black")
        self.Fillcomp2button.configure(pady="0")
        self.Fillcomp2button.configure(text='''Comp 2''')
        self.Fillcomp2button.configure(command=self.FillComp2)

        self.Fillcomp3button = tk.Button(top)
        self.Fillcomp3button.place(relx=0.631, rely=0.063, height=28, width=78)
        self.Fillcomp3button.configure(activebackground="#ececec")
        self.Fillcomp3button.configure(activeforeground="#000000")
        self.Fillcomp3button.configure(background="#d9d9d9")
        self.Fillcomp3button.configure(disabledforeground="#a3a3a3")
        self.Fillcomp3button.configure(foreground="#000000")
        self.Fillcomp3button.configure(highlightbackground="#d9d9d9")
        self.Fillcomp3button.configure(highlightcolor="black")
        self.Fillcomp3button.configure(pady="0")
        self.Fillcomp3button.configure(text='''Comp 3''')
        self.Fillcomp3button.configure(command=self.FillComp3)

        self.Fillsubjectbutton = tk.Button(top)
        self.Fillsubjectbutton.place(relx=0.332, rely=0.063, height=28, width=78)

        self.Fillsubjectbutton.configure(activebackground="#ececec")
        self.Fillsubjectbutton.configure(activeforeground="#000000")
        self.Fillsubjectbutton.configure(background="#d9d9d9")
        self.Fillsubjectbutton.configure(disabledforeground="#a3a3a3")
        self.Fillsubjectbutton.configure(foreground="#000000")
        self.Fillsubjectbutton.configure(highlightbackground="#d9d9d9")
        self.Fillsubjectbutton.configure(highlightcolor="black")
        self.Fillsubjectbutton.configure(pady="0")
        self.Fillsubjectbutton.configure(text='''Subject''')
        self.Fillsubjectbutton.configure(command=self.Fillsubject)

    ### Beginning of button section

    def Fillsubject(self):

        # fills the subject column using methods defined in a following section.
        # The other fill methods do the same thing for different columns.

        if len(self.Assessordata.get("1.0", 'end-1c')) > 0:
            self.Clearsubject()
            self.Getselection()
            self.GetSubAddress()
            self.Subjectaddressbox.insert(0, self.GetAddress())
            self.Subjectlocationbox.insert(0, self.GetLocation())
            self.Subjectpricebox.insert(0, self.GetSaleprice())
            self.Subjectdatebox.insert(0, self.GetSaledate())
            self.Subjectsaleconbox.insert(0, self.GetSalecon())
            self.Subjectstylebox.insert(0, self.GetStyle())
            self.Subjectyearbox.insert(0, self.GetYear())
            self.Subjectsizebox.insert(0, self.GetSize())
            self.Subjectbedbox.insert(0, self.GetBeds())
            self.Subjectbathbox.insert(0, self.GetBaths())
            self.Subjectbasebox.insert(0, self.GetBase())
            self.Subjectgaragebox.insert(0, self.GetGarage())
            self.Subjectotherbox.insert(0, self.GetOther())

            self.Assessordata.delete("1.0", 'end')

    def FillComp1(self):
        if len(self.Assessordata.get("1.0", 'end-1c')) > 0:
            self.Clearcomp1()
            self.Getselection()
            self.Comp1addressbox.insert(0, self.GetAddress())
            self.Comp1locationbox.insert(0, self.GetLocation())
            self.Comp1pricebox.insert(0, self.GetSaleprice())
            self.Comp1datebox.insert(0, self.GetSaledate())
            self.Comp1saleconbox.insert(0, self.GetSalecon())
            self.Comp1stylebox.insert(0, self.GetStyle())
            self.Comp1yearbox.insert(0, self.GetYear())
            self.Comp1sizebox.insert(0, self.GetSize())
            self.Comp1bedbox.insert(0, self.GetBeds())
            self.Comp1bathbox.insert(0, self.GetBaths())
            self.Comp1basebox.insert(0, self.GetBase())
            self.Comp1garagebox.insert(0, self.GetGarage())
            self.Comp1otherbox.insert(0, self.GetOther())
            self.Assessordata.delete("1.0", 'end')

    def FillComp2(self):
        if len(self.Assessordata.get("1.0", 'end-1c')) > 0:
            self.Clearcomp2()
            self.Getselection()
            self.Comp2addressbox.insert(0, self.GetAddress())
            self.Comp2locationbox.insert(0, self.GetLocation())
            self.Comp2pricebox.insert(0, self.GetSaleprice())
            self.Comp2datebox.insert(0, self.GetSaledate())
            self.Comp2saleconbox.insert(0, self.GetSalecon())
            self.Comp2stylebox.insert(0, self.GetStyle())
            self.Comp2yearbox.insert(0, self.GetYear())
            self.Comp2sizebox.insert(0, self.GetSize())
            self.Comp2bedbox.insert(0, self.GetBeds())
            self.Comp2bathbox.insert(0, self.GetBaths())
            self.Comp2basebox.insert(0, self.GetBase())
            self.Comp2garagebox.insert(0, self.GetGarage())
            self.Comp2otherbox.insert(0, self.GetOther())
            self.Assessordata.delete("1.0", 'end')

    def FillComp3(self):
        if len(self.Assessordata.get("1.0", 'end-1c')) > 0:
            self.Clearcomp3()
            self.Getselection()
            self.Comp3addressbox.insert(0, self.GetAddress())
            self.Comp3locationbox.insert(0, self.GetLocation())
            self.Comp3pricebox.insert(0, self.GetSaleprice())
            self.Comp3datebox.insert(0, self.GetSaledate())
            self.Comp3saleconbox.insert(0, self.GetSalecon())
            self.Comp3stylebox.insert(0, self.GetStyle())
            self.Comp3yearbox.insert(0, self.GetYear())
            self.Comp3sizebox.insert(0, self.GetSize())
            self.Comp3bedbox.insert(0, self.GetBeds())
            self.Comp3bathbox.insert(0, self.GetBaths())
            self.Comp3basebox.insert(0, self.GetBase())
            self.Comp3garagebox.insert(0, self.GetGarage())
            self.Comp3otherbox.insert(0, self.GetOther())
            self.Assessordata.delete("1.0", 'end')

    def Clearassdata(self):  # clears the assessor data box
        self.Assessordata.delete("1.0", 'end')

    def Clearsubject(self):  # clears the subject column
        self.Subjectaddressbox.delete("0", 'end')
        self.Subjectlocationbox.delete("0", 'end')
        self.Subjectpricebox.delete("0", 'end')
        self.Subjectdatebox.delete("0", 'end')
        self.Subjectsaleconbox.delete("0", 'end')
        self.Subjectconditionbox.delete("0", 'end')
        self.Subjectstylebox.delete("0", 'end')
        self.Subjectyearbox.delete("0", 'end')
        self.Subjectsizebox.delete("0", 'end')
        self.Subjectbedbox.delete("0", 'end')
        self.Subjectbathbox.delete("0", 'end')
        self.Subjectbasebox.delete("0", 'end')
        self.Subjectgaragebox.delete("0", 'end')
        self.Subjectotherbox.delete("0", 'end')

    def Clearcomp1(self):  # clears comp 1 column
        self.Comp1addressbox.delete("0", 'end')
        self.Comp1locationbox.delete("0", 'end')
        self.Comp1pricebox.delete("0", 'end')
        self.Comp1datebox.delete("0", 'end')
        self.Comp1conditionbox.delete("0", 'end')
        self.Comp1saleconbox.delete("0", 'end')
        self.Comp1stylebox.delete("0", 'end')
        self.Comp1yearbox.delete("0", 'end')
        self.Comp1sizebox.delete("0", 'end')
        self.Comp1bedbox.delete("0", 'end')
        self.Comp1bathbox.delete("0", 'end')
        self.Comp1basebox.delete("0", 'end')
        self.Comp1garagebox.delete("0", 'end')
        self.Comp1otherbox.delete("0", 'end')

    def Clearcomp2(self):  # clears comp 2 column
        self.Comp2addressbox.delete("0", 'end')
        self.Comp2locationbox.delete("0", 'end')
        self.Comp2pricebox.delete("0", 'end')
        self.Comp2datebox.delete("0", 'end')
        self.Comp2saleconbox.delete("0", 'end')
        self.Comp2conditionbox.delete("0", 'end')
        self.Comp2stylebox.delete("0", 'end')
        self.Comp2yearbox.delete("0", 'end')
        self.Comp2sizebox.delete("0", 'end')
        self.Comp2bedbox.delete("0", 'end')
        self.Comp2bathbox.delete("0", 'end')
        self.Comp2basebox.delete("0", 'end')
        self.Comp2garagebox.delete("0", 'end')
        self.Comp2otherbox.delete("0", 'end')

    def Clearcomp3(self):  # clears comp 3 column
        self.Comp3addressbox.delete("0", 'end')
        self.Comp3locationbox.delete("0", 'end')
        self.Comp3pricebox.delete("0", 'end')
        self.Comp3datebox.delete("0", 'end')
        self.Comp3saleconbox.delete("0", 'end')
        self.Comp3stylebox.delete("0", 'end')
        self.Comp3conditionbox.delete("0", 'end')
        self.Comp3yearbox.delete("0", 'end')
        self.Comp3sizebox.delete("0", 'end')
        self.Comp3bedbox.delete("0", 'end')
        self.Comp3bathbox.delete("0", 'end')
        self.Comp3basebox.delete("0", 'end')
        self.Comp3garagebox.delete("0", 'end')
        self.Comp3otherbox.delete("0", 'end')

    ### The following section parses the assessor data

    def Getselection(self):

        # determines the source of the pasted data.

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if "Nebraska Property Record" in data:
            self.selection = "douglas"
        if "Red Bell Real Estate" in data:
            self.selection = "redbell"
        if "KS Uniform Parcel Num" in data:
            self.selection = 'johnson'
        if "Valuation Information Valuation" in data:
            self.selection = "sarpy"
        if "Weld County" in data:
            self.selection = "weld"
        if "Larimer County" in data:
            self.selection = "larimer"  # must use legacy site http://legacy.larimer.org/assessor/query/search.cfm
        if "============" in data:
            self.selection = "cb"
        if "Great Plains Regional MLS" in data:
            self.selection = "mls"
        if "Property Report for Account " in data:
            self.selection = 'boulder'

    def GetSubAddress(self):

        # gets more specific address info for the subject property and grabs assessed value

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data1 = (data.split("Parcel Address:  "))[1].split("-")[0].title()
            addressdat = data1.split()
            self.subjectzip = addressdat[-1]
            self.subjectstate = "NE"
            if "Douglas County" in data1:
                self.subjectcity = "Omaha"
            else:
                self.subjectcity = addressdat[-3]
            self.subjectcounty = "Douglas"
            self.subjectassyear = data.split("Improvement	Total")[1].split()[0].strip()
            temp = (data.split("Improvement	Total"))[1].split("Sales Information")[0].title()
            self.subjectassvalue = temp.split()[3][:-3]

        if self.selection == 'redbell':
            self.subjectzip = data.split("Zip Code:")[1].split()[0].strip()
            self.subjectcounty = data.split("County:")[1].split()[0].strip()
            self.subjectstate = data.split("State:")[1].split()[0].strip()
            self.subjectcity = (data.split("City:"))[1].split("County:")[0].title().strip()

        if self.selection == 'johnson':
            addressdat = (data.split("Site Address:"))[1].split("Legal Description:")[0].title().strip().split()

            self.subjectzip = addressdat[-1]
            self.subjectstate = "KS"
            self.subjectcity = (data.split("City/Township:"))[1].split("Quarter Section:")[0].title().strip()
            self.subjectcounty = "Johnson"
            temp = data.split("Value	Change:")[1].split()[0].strip()
            self.subjectassyear = data.split(temp)[1].split()[0].strip()
            temp = (data.split("Value	Change:"))[1].split("Main Dwelling")[0].title()
            self.subjectassvalue = temp.split()[2]

        if self.selection == 'sarpy':
            self.subjectstate = "NE"
            self.subjectcounty = "Sarpy"
            self.subjectcity = ""
            self.subjectassyear = data.split("Form191")[1].split()[0].strip()
            temp = (data.split("Form191"))[1].split("Residential Information for")[0].title()
            self.subjectassvalue = temp.split()[4]

        if self.selection == 'weld':
            self.subjectstate = "CO"
            self.subjectcounty = "Weld"
            self.subjectcity = ""

            temp = (data.split("Assessed Value"))[1].split("Account	Owner")[0].title()
            self.subjectassvalue = "$" + temp.split()[5]
            self.subjectassyear = temp.split()[3]

        if self.selection == 'larimer':
            self.subjectstate = "CO"
            self.subjectcounty = "Larimer"
            self.subjectcity = ""
            tt = (data.split("Property Address"))[1].split("-")[0].strip()
            tt = tt.split()
            self.subjectzip = tt[-1]

            temp = (data.split("Assessed Value"))[1].split("Account	Owner")[0].title()
            self.subjectassvalue = data.split("Totals:")[1].split()[0].strip()
            self.subjectassyear = data.split("Property Tax Year:")[1].split()[0].strip()

        if self.selection == 'cb':
            self.subjectcity = ""

            temp = self.Assessordata.get("1.0", 'end-1c')

            self.subjectcity = self.Assessordata.get("1.0", 'end-1c').splitlines()[5].split(",")[0].title()

            self.subjectstate = "IA"
            self.subjectcounty = "Pottawattamie"

            self.subjectzip = data.split(", IA")[1].split()[0].strip()

            temp = (data.split("class*"))[1].split("R")[0].title()
            temp = temp.split()
            self.subjectassvalue = temp[-2]
            self.subjectassyear = temp[-1]

        if self.selection == 'mls':
            self.subjectstate = "NE"
            self.subjectcounty = (data.split("County"))[1].split(" County")[0].strip()
            self.subjectcity = (data.split("City"))[1].split("Zip")[0].strip()
            self.subjectzip = (data.split("Zip"))[1].split("State")[0].strip()

            self.subjectassvalue = ""
            self.subjectassyear = ""

        if self.selection == 'boulder':
            self.subjectcity = (data.split("City:"))[1].split("Owner:")[0].strip().title()
            self.subjectstate = "CO"
            self.subjectcounty = "Boulder"

            temp = (data.split("City, State, Zip:"))[1].split("Sec-Town-Range:")[0].strip().title()
            temp = temp.split()
            self.subjectzip = temp[-1]

            temp = (data.split("Actual	Assessed"))[1].split("X-Features:")[0].strip().split()

            self.subjectassvalue = "$" + str((format(int(temp[1]), ',d')))
            self.subjectassyear = ""

    def GetAddress(self):

        # gets address info.  Some websites have formatting issues that are corrected

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data = (data.split("Parcel Address:  "))[1].split("-")[0].strip()

            if "DOUGLAS COUNTY" in data:
                data = data.replace("DOUGLAS COUNTY", '').strip()
                data = ((data.rsplit(' ', 1)[0]).rsplit(' ', 1)[0]).strip()
            else:
                data = ((data.rsplit(' ', 1)[0]).rsplit(' ', 1)[0]).rsplit(' ', 1)[0].strip()

            test = data.split()[-1]

            if test == "AV":
                data = data + "e"
            if test == "CR":
                data = data.rsplit(' ', 1)[0]
                data = data + " Cir"
            if test == "PA":
                data = data.rsplit(' ', 1)[0]
                data = data + " Plz"

            def GetOrdinal(number):  # Douglas county addresses don't include 'th' 'nd' 'st'. This adds them back
                if type(number) != type(1):
                    try:
                        number = int(number)
                    except:
                        raise ValueError("This number is not an Int!")

                lastdigit = int(str(number)[len(str(number)) - 1])
                last2 = int(str(number)[len(str(number)) - 2:])
                if last2 > 10 and last2 < 13:
                    return str(number) + "th"
                if lastdigit == 1:
                    return str(number) + "st"
                if lastdigit == 2:
                    return str(number) + "nd"
                if lastdigit == 3:
                    return str(number) + "rd"
                return str(number) + "th"

            data = data.split()

            if data[1].isdigit():
                data[1] = GetOrdinal(data[1])

            if len(data) > 2:
                if data[2].isdigit():
                    data[2] = GetOrdinal(data[2])

            return " ".join([word.capitalize() for word in data])

        if self.selection == 'redbell':
            return (data.split("Address:"))[1].split("Unit#:")[0].title().strip()

        if self.selection == 'johnson':
            data = (data.split("Site Address:"))[1].split(",")[0].strip()
            if ("BONNER SPRINGS" in data) or ("DE SOTO" in data) or ("LAKE QUIVIRA" in data) or (
                    "MISSION HILLS" in data) or ("MISSION WOODS" in data) or ("OVERLAND PARK" in data) or (
                    "PRAIRIE VILLAGE" in data) or ("ROELAND PARK" in data) or ("SPRING HILL" in data) or (
                    "WESTWOOD HILLS" in data):
                data = ((data.rsplit(' ', 1)[0]).rsplit(' ', 1)[0]).strip()
            else:
                data = ((data.rsplit(' ', 1)[0])).strip()

            return " ".join([word.capitalize() for word in data.split()])

        if self.selection == 'sarpy':
            data = (data.split("Location:"))[1].split("Owner:")[0].title().strip()
            data = data.replace("\\", "")
            while str(data)[0] == "0":
                data = data[1:]
            return " ".join([word.capitalize() for word in data.split()])

        if self.selection == 'weld':
            data = (data.split("Township	Range"))[1].split("Close Section")[0].strip()
            data = data[:-10].strip()
            data = data.replace("\t", " ")
            if "GARDEN CITY" in data:
                data = data.replace("GARDEN CITY", "")
            elif "FORT LUPTON" in data:
                data = data.replace("FORT LUPTON", "")
            elif "WELD" in data:
                data = data.replace("WELD", "")
            else:
                data = ' '.join(data.split(' ')[:-1])

            return " ".join([word.capitalize() for word in data.split()])

        if self.selection == 'larimer':
            data = (data.split("Property Address"))[1].split("-")[0].strip()
            if "FORT COLLINS" in data:
                data = data.replace("FORT COLLINS", "FORT")
            if "ESTES PARK" in data:
                data = data.replace("ESTES PARK", "ESTES")
            data = data.rsplit(' ', 1)[0]
            data = data.rsplit(' ', 1)[0]
            return " ".join([word.capitalize() for word in data.split()])

        if self.selection == 'cb':
            data = self.Assessordata.get("1.0", 'end-1c').splitlines()[4].split("      ")[0].title()
            data = data.split()
            return " ".join([word.capitalize() for word in data])

        if self.selection == 'mls':
            return (data.split("Address"))[1].split("Unit #")[0].strip().title()

        if self.selection == 'boulder':
            return (data.split("Property Address:"))[1].split("City:")[0].strip().title()

    def GetSaleprice(self):

        # some websites don't share sale price/date

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':

            if "Price:" in data:
                return data.split("Price:")[1].split()[0].strip()[:-3]
            else:
                return ""

        if self.selection == 'redbell':
            return data.split("Sold Price:")[1].split()[0].strip()

        if self.selection == 'johnson':
            return ""
        if self.selection == 'sarpy':
            data = (data.split("Sales Information"))[1].split("GIS Information")[0]
            data = data.split()
            for word in data:
                if word[0] == "$":
                    return word
            return ""

        if self.selection == 'weld':
            data = (data.split("Document History"))[1].split("Close Section")[0].title().strip()
            data = data.split()
            return "$" + data[-1]

        if self.selection == 'larimer':
            data = (data.split("view the document."))[1].split("Value Information")[0].strip()
            data = data.split()
            return data[10]

        if self.selection == 'cb':
            if "Book/Page" not in data:
                return ""
            else:
                data = (data.split("Book/Page"))[1].split("Interior Listing:")[0].strip()
                data = data.split()
                return "$" + data[1]

        if self.selection == 'mls':
            data = data.replace("Sold Price Per", "")
            return (data.split("Sold Price"))[1].split("Selling CommentsSelling")[0].strip()

        if self.selection == 'boulder':
            data = (data.split("Recorded	Sale Price"))[1].split("Zoning Report")[0].strip().split()
            temp = data[3]
            return temp[:-3]

    def GetSaledate(self):

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            if "Sales Information Sales Date:" in data:
                temp = data.split("Sales Information Sales Date:")[1].split()[0].strip()
                temp = temp.split("-")
                return temp[1] + "/" + temp[2] + "/" + temp[0]
            else:
                return ""

        if self.selection == 'redbell':
            return data.split("Sold Date:")[1].split()[0].strip() + " - " + data.split("DOM:")[1].split()[
                0].strip() + " DOM"

        if self.selection == 'johnson':
            return ""

        if self.selection == 'sarpy':
            return data.split("Adjusted Sale Price")[1].split()[0].strip()
        if self.selection == 'weld':
            data = (data.split("Document History"))[1].split("Close Section")[0].title().strip()
            data = data.split()
            return data[-2].replace("-", "/")

        if self.selection == 'larimer':
            data = (data.split("view the document."))[1].split("Value Information")[0].strip()
            data = data.split()
            return data[8]

        if self.selection == 'cb':
            if "Book/Page" not in data:
                return ""
            else:
                data = (data.split("Book/Page"))[1].split("Interior Listing:")[0].strip()
                data = data.split()
                return data[0]

        if self.selection == 'mls':
            return (data.split("Closing Date"))[1].split("Sold Price")[0].strip()

        if self.selection == 'boulder':
            data = (data.split("Recorded	Sale Price"))[1].split("Zoning Report")[0].strip().split()
            return data[1]

    def GetSalecon(self):

        # Assumed to be normal.  Only a few websites report.

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'redbell':
            if "REO: Yes" in data:
                return "REO Sale"
            if "Short Sale: Yes" in data:
                return "Short Sale"
            else:
                return "Arms-Length"

        return "Arms-length"

    def GetLocation(self):

        # returns 0.25 ac.  Some websites share other data like noisy street

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data2 = data.split("Width	Vacant ")[1].split()[0]
            data2 = format(float(data2), '.2f')

            if "Negative Influence" in data and "Traffic" in data:
                return str(data2) + " ac" + " Traffic"
            else:
                return str(data2) + " ac" + " Res"

        if self.selection == 'redbell':
            return data.split("Lot Size:")[1].split()[0].strip() + " ac"

        if self.selection == 'johnson':
            return (data.split("Property Area:"))[1].split("Addresses:")[0].strip()[:-3]

        if self.selection == 'sarpy':
            lot1 = float(data.split("Lot Depth:")[1].split()[0].strip())
            lot2 = float(data.split("Lot Width:")[1].split()[0].strip())
            return str(round((lot1 * lot2) / 43560, 2)) + " ac"

        if self.selection == 'weld':
            if "Condominium unit" in data:
                return "Condo"
            data = (data.split("Totals"))[1].split("Comparable sales")[0].strip()[:-3]
            data = data.split()
            return str(round(float(data[-2]), 2)) + " ac"

        if self.selection == 'larimer':
            if "Townhouse/Condo" in data:
                return "Condo"
            data = (data.split("Totals:"))[1].split("Building Improvements")[0].strip()
            data = data.split()
            return data[2] + " ac"

        if self.selection == 'cb':
            if (data.split("sqFt"))[1].split("acres")[0].strip() == "":
                return "Condo"
            else:
                data = (data.split("sqFt"))[1].split("acres")[0].strip()
                return str(round(float(data), 2)) + " ac"

        if self.selection == 'mls':
            return ""

        if self.selection == 'boulder':
            return data.split("Acres:")[1].split()[0].strip() + " ac"

    def GetStyle(self):

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data = (data.split("Built As: "))[1].split("Condo Square")[0].title().strip()
            if data == "1 1/2 Story Fin":
                return "1.5 Story"
            if data == "2 1/2 Story Fin":
                return "2.5 Story"
            if data == "Townhouse 1 1/2 Story":
                return "Townhouse 1.5 Story"
            return data

        if self.selection == 'redbell':
            return (data.split("Style:"))[1].split("Bath:")[0].title().strip()

        if self.selection == 'johnson':
            temp = (data.split("Style:"))[1].split("Total Rooms:")[0].title().strip()
            if temp == "Reverse One-And-One Half":
                return "Ranch"
            if temp == "Conventional":
                return "2-Story"
            else:
                return temp

        if self.selection == 'sarpy':
            return (data.split("Style:"))[1].split("Click Picture")[0].title().strip()

        if self.selection == 'weld':
            data = (data.split("Width 1.00"))[1].split("Additional Details")[0].title().strip()
            if "Condo <= 3 Stories" in data:
                return "Condo"
            data = data.split()
            return " ".join([word.capitalize() for word in data[:-5]])

        if self.selection == 'larimer':
            return (data.split("Built As:"))[1].split("Occupancy:")[0].title().strip()

        if self.selection == 'cb':
            data = (data.split("BUILDING....."))[1].split("/")[0].title().strip()
            return data.rsplit(' ', 1)[0]

        if self.selection == 'mls':
            return (data.split("Style"))[1].split("Agreement Type")[0].strip()

        if self.selection == 'boulder':
            return (data.split("Design:"))[1].split("Number of rooms:")[0].strip()

    def GetYear(self):  # year built

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data = data.split("Physical Age")[1].split()[0].strip()
            return data

        if self.selection == 'redbell':
            return data.split("Yr Built:")[1].split()[0].strip()

        if self.selection == 'sarpy' or self.selection == 'larimer' or self.selection == 'johnson':
            return data.split("Year Built:")[1].split()[0].strip()

        if self.selection == 'weld':
            data = (data.split("Width 1.00"))[1].split("Additional Details")[0].title().strip()
            data = data.split()
            return data[-4]

        if self.selection == 'cb':
            data = (data.split("Built:"))[1].split("Bsmnt")[0].title().strip()
            data = data.split()
            return data[0]

        if self.selection == 'mls':
            return (data.split("Year Built"))[1].split("New Construction")[0].strip()

        if self.selection == 'boulder':
            return (data.split("Built:"))[1].split("Design:")[0].strip()

    def GetSize(self):

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data = data.split("Footage:")[1].split()[0].strip()[:-2]
            return data

        if self.selection == 'redbell':
            return data.split("Above Grade SqFt:")[1].split()[0].strip()

        if self.selection == 'johnson':
            return data.split("Total SFLA:")[1].split()[0].strip().replace(",", "")

        if self.selection == 'sarpy':
            return data.split("Total Sqft:")[1].split()[0].strip()

        if self.selection == 'weld':
            data = (data.split("Width 1.00"))[1].split("Additional Details")[0].title().strip()
            data = data.split()
            return data[-5]

        if self.selection == 'larimer':
            return data.split("Total Sq Ft:")[1].split()[0].strip()

        if self.selection == 'cb':
            if "Attic Finish: None" not in data:
                attic = int(data.split("Attic Finish:")[1].split()[0].strip())
                return int((data.split("Bedrooms Above/Below"))[1].split("SF")[0].title().strip()) + attic
            else:
                return (data.split("Bedrooms Above/Below"))[1].split("SF")[0].title().strip()

        if self.selection == 'mls':
            return (data.split("Above Grade SQFT"))[1].split("Total Finished SqFt")[0].strip()

        if self.selection == 'boulder':
            firstfloor = int(data.split("FIRST FLOOR (ABOVE GROUND) FINISHED AREA")[1].split()[0].strip())
            if "2ND FLOOR AND HIGHER FINISHED AREA" in data:
                secondfloor = int(data.split("2ND FLOOR AND HIGHER FINISHED AREA")[1].split()[0].strip())
            else:
                secondfloor = 0
            return firstfloor + secondfloor

    def GetBeds(self):  # Some websites split up above/below grade beds.  They are combined.

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            data = data.split("Bedrooms:")[1].split()[0].strip()[:-2]
            return data

        if self.selection == 'redbell':
            return data.split("Bed:")[1].split()[0].strip()

        if self.selection == 'johnson':
            return data.split("Bedrooms:")[1].split()[0].strip()

        if self.selection == 'sarpy':
            return data.split("#Bedrooms above Grade:")[1].split()[0].strip()

        if self.selection == 'weld':
            data = (data.split("Baths	Rooms"))[1].split("ID")[0].title().strip()
            data = data.split()
            return data[-3]

        if self.selection == 'larimer':
            return data.split("Bedrooms:")[1].split()[0].strip()

        if self.selection == 'cb':
            data = (data.split("Rooms Above/Below"))[1].split("Bedrooms Above/Below")[0].title().strip()
            data = data.split('/')
            return int(data[0]) + int(data[1])

        if self.selection == 'mls':
            return (data.split("Bedrooms"))[1].split("Bathrooms")[0].strip()

        if self.selection == 'boulder':
            return (data.split("Bedrooms:"))[1].split("Full Bath:")[0].strip()

    def GetBaths(self):

        # Attempts to return number of full and half baths. formatted as 2.1 or 2.5
        # Some sites are too difficult to parse and some only report full baths.

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if (self.selection == 'douglas'):
            fullb = float(data.split("Bath Full")[1].split()[0].strip())
            halfb = 0
            if "Bath Half" in data:
                halfb = float(data.split("Bath Half")[1].split()[0].strip())
            return float(fullb + .1 * halfb)

        if self.selection == 'redbell':
            temp = data.split("Bath:")[1].split()[0].strip()
            if temp == "0.75":
                return "1.0"
            if temp == "1.25":
                return "1.1"
            if temp == "1.75":
                return "2.0"
            if temp == "2.25":
                return "2.1"
            if temp == "2.75":
                return "3.0"
            if temp == "3.25":
                return "3.1"
            if temp == "3.75":
                return "4.0"
            if temp == "4.1":
                return "4.5"
            if temp == "4.75":
                return "5.0"
            if temp == "5.25":
                return "5.1"
            if temp == "0":
                return "1.0"
            if temp == "1.00":
                return "1.0"
            if temp == "2.00":
                return "2.0"
            if temp == "3.00":
                return "3.0"
            return temp

        if self.selection == 'johnson':
            if data.split("Half Baths:")[1].split()[0].strip() == "Finish":
                return data.split("Full Baths:")[1].split()[0].strip() + ".0"
            else:
                return data.split("Full Baths:")[1].split()[0].strip() + "." + data.split("Half Baths:")[1].split()[
                    0].strip()

        if self.selection == 'sarpy':
            return data.split("#Bathrooms Above Grade:")[1].split()[0].strip()

        if self.selection == 'weld':
            fullb = 0
            halfb = 0
            if "Fixture	Bath 2" in data:
                halfb += float(data.split("Fixture	Bath 2")[1].split()[0].strip())

            if "Fixture	Bath 3" in data:
                fullb += float(data.split("Fixture	Bath 3")[1].split()[0].strip())

            if "Fixture	Bath 4" in data:
                fullb += float(data.split("Fixture	Bath 4")[1].split()[0].strip())

            if "Fixture	Bath 5" in data:
                fullb += float(data.split("Fixture	Bath 5")[1].split()[0].strip())

            if "Fixture	Bath 6" in data:
                fullb += float(data.split("Fixture	Bath 6")[1].split()[0].strip())

            return float(fullb + .1 * halfb)

        if self.selection == 'larimer':
            temp = data.split("Baths:")[1].split()[0].strip()
            if temp == "0.75":
                return "1.0"
            if temp == "1.25":
                return "1.1"
            if temp == "1.50":
                return "1.1"
            if temp == "1.75":
                return "2.0"
            if temp == "2.25":
                return "2.1"
            if temp == "2.50":
                return "2.1"
            if temp == "2.75":
                return "3.0"
            if temp == "3.25":
                return "3.1"
            if temp == "3.50":
                return "3.1"
            if temp == "3.75":
                return "4.0"
            if temp == "4.25":
                return "4.1"
            if temp == "4.50":
                return "4.1"
            if temp == "4.75":
                return "5.0"
            if temp == "5.25":
                return "5.1"
            if temp == "5.50":
                return "5.1"
            if temp == "1.00":
                return "1.0"
            if temp == "2.00":
                return "2.0"
            if temp == "3.00":
                return "3.0"
            return temp

        if self.selection == 'cb':  # too difficult
            return ""

        if self.selection == 'mls':
            return (data.split("Bathrooms"))[1].split("# of Fireplaces")[0].strip()

        if self.selection == 'boulder':
            fullbath = int((data.split("Full Bath:"))[1].split("3/4 Bath:")[0])
            threebath = int((data.split("3/4 Bath:"))[1].split("Half Bath:")[0])
            halfbath = int((data.split("Half Bath:"))[1].split("Areas of levels")[0])

            return float(fullbath + threebath + .1 * halfbath)

    def GetBase(self):

        # return "None" if there is no basement
        # return "[size] Unfinished" if basement is unfinished
        # else return both size and finished amount

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            if "Basement	Finished" in data:
                basefin = data.split("Basement	Finished")[1].split()[0].strip()[:-2]
            else:
                basefin = 0
            if "Bsmnt" in data:
                list_of_words = data.split()
                base = list_of_words[list_of_words.index("Bsmnt") + 4][:-2]
            else:
                base = "None"

            if base == "None":
                return base
            elif basefin == 0:
                return base + " Unfinished"
            else:
                return base + " - " + basefin + " Fin"

        if self.selection == 'redbell':
            basefin = ""
            base = ""
            if len((data.split("Below Grade SqFt:"))[1].split("Below Grade Finished SqFt:")[0].title().strip()) > 2:
                base = data.split("Below Grade SqFt:")[1].split()[0].strip()
            if len((data.split("Basement Finished:"))[1].split("Basement Finished %:")[0].title().strip()) > 2:
                basefin = data.split("Basement Finished %:")[1].split()[0].strip() + "% Fin"
            if len(base) > 1:
                if len(basefin) > 1:
                    return base + " sf - " + basefin
                return base + " sf Unfinished"
            else:
                return "None"

        if self.selection == 'johnson':
            basefin = ""
            base = ""
            if "Total Basement Area (SF)" in data:
                base = data.split("Total Basement Area (SF)")[1].split()[0].strip().replace(",", "")
                basefin = data.split("Finish Bsmt:")[1].split()[0].strip().replace(",", "")
                if len(basefin) > 2:
                    return base + " - " + basefin + " Fin"
                return base + " sf Unfinished"
            else:
                return "None"

        if self.selection == 'sarpy':
            if (data.split("Bsmt Total Sqft:"))[1].split("Garage Type:")[0].title().strip() == "":
                return "None"
            if (data.split("Bsmt Total Sqft:"))[1].split("Garage Type:")[0].title().strip() == "0":
                return "None"
            base = int((data.split("Bsmt Total Sqft:"))[1].split("Garage Type:")[0].title().strip())

            if (data.split("Total Bsmt Finish Sqft:"))[1].split("Bsmt Total Sqft:")[0].title().strip() == "":
                return str(base) + " sf Unfinished"

            if (data.split("Total Bsmt Finish Sqft:"))[1].split("Bsmt Total Sqft:")[0].title().strip() == "0":
                return str(base) + " sf Unfinished"

            basefin = int((data.split("Total Bsmt Finish Sqft:"))[1].split("Bsmt Total Sqft:")[0].title().strip())

            return str(base) + " - " + str(basefin) + " Fin"

        if self.selection == 'weld':
            data = (data.split("Porch SF"))[1].split("Built As")[0].title().strip()
            data = data.split()
            base = data[-6]
            basefin = data[-5]
            if base == "0":
                return "None"
            elif basefin == '0':
                return base + " sf Unfinished"
            else:
                return base + " - " + basefin + " Fin"

        if self.selection == 'larimer':
            if "Bsmt. Sq Ft:" not in data:
                return "None"
            elif "Bsmt. Fin. Sq Ft:" not in data:
                return data.split("Bsmt. Sq Ft:")[1].split()[0].strip() + " sf Unfinished"
            else:
                return data.split("Bsmt. Sq Ft:")[1].split()[0].strip() + " - " + \
                       data.split("Bsmt. Fin. Sq Ft:")[1].split()[0].strip() + " Fin"

        if self.selection == 'cb':
            if "Bsmt: None" in data:
                return "None"
            elif "Bsmt Finish: None" in data:
                return data.split("Bsmt:")[1].split()[0].strip() + " Unfinished"
            else:
                return data.split("Bsmt:")[1].split()[0].strip() + " - " + data.split("Bsmt Finish:")[1].split()[
                    0].strip() + " Fin"

        if self.selection == 'mls':
            if "BasementYes" in data:
                if (data.split("Finished Below Grade"))[1].split("Above Grade SQFT")[0].strip() == "0":
                    return "Unfinished"
                else:
                    return (data.split("Finished Below Grade"))[1].split("Above Grade SQFT")[0].strip() + " Fin"
            else:
                return "None"

        if self.selection == 'boulder':
            if "LOWER LVL GARDEN FINISHED (BI-SPLIT LVL)" in data:
                return data.split("LOWER LVL GARDEN FINISHED (BI-SPLIT LVL)")[1].split()[0].strip() + " - " + \
                       data.split("LOWER LVL GARDEN FINISHED (BI-SPLIT LVL)")[1].split()[0].strip() + " Fin"
            elif ("LOWER LVL GARDEN FINISHED (BI-SPLIT LVL)" not in data) and (
                    "SUBTERRANEAN BASEMENT UNFINISHED AREA" not in data):
                return "None"
            elif "SUBTERRANEAN BASEMENT FINISHED AREA" not in data:
                return data.split("SUBTERRANEAN BASEMENT UNFINISHED AREA")[1].split()[0].strip() + " sf Unfinished"
            else:
                fin = int(data.split("SUBTERRANEAN BASEMENT FINISHED AREA")[1].split()[0].strip())
                unfin = int(data.split("SUBTERRANEAN BASEMENT UNFINISHED AREA")[1].split()[0].strip())
                return str(unfin + fin) + " - " + str(fin) + " Fin"

    def GetGarage(self):

        # returns car capacity.  Makes assumptions based on square footage

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            if "Basement Double" in data:
                return "2-Car"
            elif "Basement Single" in data:
                return "1-Car"
            elif "Garage	Built In" in data:
                size = data.split("Garage	Built In")[1].split()[0].strip()[:-2]
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 700:
                    return "2-Car"
                else:
                    return "3-Car"

            elif "Garage	Attached" in data:
                size = data.split("Garage	Attached")[1].split()[0].strip()[:-2]
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"

            elif "Garage	Detached" in data:
                size = data.split("Garage	Detached")[1].split()[0].strip()[:-2]
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            else:
                return "None"

        if self.selection == 'redbell':
            if "Garage/Carport: None" in data:
                return "None"
            if data.split("Garage Spaces:")[1].split()[0].strip() == "1.00":
                return "1-Car"
            if data.split("Garage Spaces:")[1].split()[0].strip() == "2.00":
                return "2-Car"
            if data.split("Garage Spaces:")[1].split()[0].strip() == "3.00":
                return "3-Car"

        if self.selection == 'johnson':
            if "Basement Garage, Double (#)" in data:
                return "2-Car"

            if "Basement Garage, Single (#)" in data:
                return "1-Car"

            elif "Attached Garage (SF)" in data:
                size = data.split("Attached Garage (SF)")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"

            elif "Detached Garage (SF)" in data:
                size = data.split("Detached Garage (SF)")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            else:
                return "None"

        if self.selection == 'sarpy':
            if (data.split("Garage Sqft:"))[1].split("Lot Depth:")[0].title().strip() == '':
                return "None"
            if (data.split("Garage Sqft:"))[1].split("Lot Depth:")[0].title().strip() == '0':
                return "None"
            size = int((data.split("Garage Sqft:"))[1].split("Lot Depth:")[0].title().strip())
            if 0 <= size <= 360:
                return "1-Car"
            elif 361 <= size <= 625:
                return "2-Car"
            else:
                return "3-Car"

        if self.selection == 'weld':
            data = (data.split("Porch SF"))[1].split("Built As")[0].title().strip()
            data = data.split()
            size = int(str(data[-4]).replace(",", ""))
            if 0 <= size <= 360:
                return "1-Car"
            elif 361 <= size <= 625:
                return "2-Car"
            else:
                return "3-Car"
        if self.selection == 'larimer':
            if "Garage	Built In" in data:
                size = data.split("Garage	Built In")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            elif "Garage	Attached" in data:
                size = data.split("Garage	Attached")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            elif "Garage	Detached" in data:
                size = data.split("Garage	Detached")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            else:
                return "None"

        if self.selection == 'cb':
            if "2 Bsmt Stalls" in data:
                return "2-Car"
            if "1 Bsmt Stall" in data:
                return "1-Car"
            if "Garage 1:" in data:
                size = data.split("Garage 1:")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            else:
                return "None"

        if self.selection == 'mls':
            if (data.split("Garage Spaces"))[1].split("3rd Floor SqFt")[0].title().strip() == '0':
                return "None"
            else:
                return (data.split("Garage Spaces"))[1].split("3rd Floor SqFt")[0].title().strip() + "-Car"

        if self.selection == 'boulder':
            if "ATTACHED GARAGE AREA" in data:
                size = data.split("ATTACHED GARAGE AREA")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"

            elif "DETACHED GARAGE" in data:
                size = data.split("DETACHED GARAGE")[1].split()[0].strip()
                size = int(size)
                if 0 <= size <= 360:
                    return "1-Car"
                elif 361 <= size <= 625:
                    return "2-Car"
                else:
                    return "3-Car"
            else:
                return "None"

    def GetOther(self):

        # checks for walkout basement, 2nd garage, outbuilding, etc.
        # if no extra features are found, return "Typical"

        data = self.Assessordata.get("1.0", 'end-1c')  # creates string based on pasted data
        data = data.replace('\n', ' ')

        if self.selection == 'douglas':
            Amen = ""

            if "Basement	Walkout" in data:
                Amen = Amen + "Walkout Basement "
            if "Garage	Detached" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Detached" in data and "Garage	Built In" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Built In" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'redbell':
            return "Typical"

        if self.selection == 'johnson':
            Amen = ""

            if "Basement Type:   Walkout" in data:
                Amen = Amen + "Walkout Basement "
            if "Pool," in data:
                Amen = Amen + "Pool "
            if "Attached Garage (SF)" in data and "Detached Garage (SF)" in data:
                Amen = Amen + " Det Garage "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'sarpy':
            Amen = ""

            if "BSMT OUTSIDE ENTRY" in data:
                Amen = Amen + "Walkout Basement"
            if "BSMT OUTSIDE~ENTRY" in data:
                Amen = Amen + "Walkout Basement"
            if "BLDG,POLE UTILITY" in data:
                Amen = Amen + " - Outbuilding"
            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'weld':
            Amen = ""

            if "Basement	Walkout" in data:
                Amen = Amen + "Walkout Basement "
            if "Out Building" in data:
                Amen = Amen + "Building "
            if "Garage	Detached" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Detached" in data and "Garage	Built In" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Built In" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'larimer':
            Amen = ""

            if "Basement	Outside Entrance" in data:
                Amen = Amen + "Walkout Basement "
            if "Out Building" in data:
                Amen = Amen + "Building "
            if "Garage	Detached" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Detached" in data and "Garage	Built In" in data:
                Amen = Amen + " Det Garage "
            if "Garage	Built In" in data and "Garage	Attached" in data:
                Amen = Amen + " Det Garage "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'cb':
            Amen = ""

            if "Garage 2:" in data:
                Amen = Amen + "Det Garage "

            if "Utility Building" in data:
                Amen = Amen + "Building "

            if " Barn " in data:
                Amen = Amen + "Building "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'mls':
            Amen = ""

            if "Out Building" in data:
                Amen = Amen + "Building "

            if "Walk-Out BasementYes" in data:
                Amen = Amen + "Walkout Basement "

            if len(Amen) > 1:
                return Amen
            else:
                return "Typical"

        if self.selection == 'boulder':
            Amen = ""

            if "ATTACHED GARAGE AREA" in data and "DETACHED GARAGE" in data:
                return "Detached Garage"
            else:
                return "Typical"

    ### The following section calculates inferior/superior/similar adjustments

    def GLAAdjuster(self, comp):

        # compares subject and comp.
        # similar is returned if size varies by less than the user defined percentage

        if comp.get() == "":
            return ""
        c = float(re.sub('[^0-9]', '', comp.get()))
        s = float(re.sub('[^0-9]', '', self.Subjectsizebox.get()))
        per = s * float(self.GLApercent.get()) / 100

        if s == c:
            return ""
        elif s - per <= c <= s + per:
            return "Similar"
        elif s > c:
            return "Inferior"
        else:
            return "Superior"

    def YearAdjuster(self, comp):

        # similar is returned if the age difference is less what the user sets

        if comp.get() == "":
            return ""
        c = int(re.sub('[^0-9]', '', comp.get()))
        s = int(re.sub('[^0-9]', '', self.Subjectyearbox.get()))

        if s == c:
            return ""
        elif (s - int(self.Yeardif.get()) <= c <= s + int(self.Yeardif.get())):
            return "Similar"
        elif (s > c):
            return "Inferior"
        else:
            return "Superior"

    def LotAdjuster(self, comp):

        # similar is returned if lot size difference is less than what user sets

        if comp.get() == "":
            return ""
        c = comp.get().split()
        s = self.Subjectlocationbox.get().split()
        c = float(c[0])
        s = float(s[0])

        per = s * float(self.Lotpercent.get()) / 100

        if s == c:
            return ""
        elif (s - per <= c <= s + per):
            return "Similar"
        elif (s > c):
            return "Inferior"
        else:
            return "Superior"

    def BedAdjuster(self, comp):

        # if field empty return nothing
        # if bed adjustments disabled: return "Similar"

        if comp.get() == "":
            return ""

        if self.Bedchecker.get() == 0:
            return "Similar"

        c = comp.get()
        s = self.Subjectbedbox.get()

        if int(s) == int(c):
            return ""
        elif int(s) > int(c):
            return "Inferior"
        else:
            return "Superior"

    def StyleAdjuster(self, comp):

        # Ranch is superior to all styles
        # All non ranches are similar

        if comp.get() == "":
            return ""
        c = comp.get()
        s = self.Subjectstylebox.get()
        if c == s:
            return ""
        elif c == "Ranch" and s != "Ranch":
            return "Superior"
        elif s == "Ranch" and c != "Ranch":
            return "Inferior"
        elif "Townhouse" in s and "Townhouse" not in c:
            return "Superior"
        elif "Townhouse" in c and "Townhouse" not in s:
            return "Inferior"
        else:
            return "Similar"

    def BathAdjuster(self, comp):
        if comp.get() == "":
            return ""
        c = comp.get()
        s = self.Subjectbathbox.get()

        if c == "": return ""

        if float(s) == float(c):
            return ""
        elif float(s) > float(c):
            return "Inferior"
        else:
            return "Superior"

    def GarAdjuster(self, comp):
        if comp.get() == "":
            return ""
        c = comp.get()
        s = self.Subjectgaragebox.get()

        if s == c:
            return ""
        elif "None" in s and "None" not in c:
            return "Superior"
        elif "None" in c and "None" not in s:
            return "Inferior"
        elif int(s[0]) > int(c[0]):
            return "Inferior"
        else:
            return "Superior"

    def BaseAdjuster(self, comp):  # Compares finished and unfinished area.
        # Gives more weight to finished space.  Many different rules.
        if comp.get() == "":
            return ""
        c = [int(s) for s in str(comp.get()).replace(",", "").split() if s.isdigit()]
        s = [int(s) for s in str(self.Subjectbasebox.get()).replace(",", "").split() if s.isdigit()]
        if c == s:
            return ""
        if not c and not s:
            return ""
        elif not c:
            return "Inferior"
        elif not s:
            return "Superior"
        elif len(c) == 1 and len(s) == 1:
            per = s[0] / 10
            if s[0] - per <= c[0] <= s[0] + per:
                return "Similar"
            elif s[0] > c[0]:
                return "Inferior"
            else:
                return "Superior"
        elif len(c) == 1 and len(s) == 2:
            if c[0] > 3 * s[0]:
                return "Superior"
            if s[0] + (s[1] * 2) < c[0]:
                return "Similar"
            else:
                return "Inferior"
        elif len(c) == 2 and len(s) == 1:
            if s[0] > 3 * c[0]:
                return "Inferior"
            if c[0] + (c[1] * 2) < c[0]:
                return "Similar"
            else:
                return "Superior"
        else:
            per = s[0] * .1
            per1 = s[1] * .1
            per2 = (s[0] + (2 * s[1])) * .1
            if s[0] - per <= c[0] <= s[0] + per:
                if s[1] - per1 <= c[1] <= s[1] + per1:
                    return "Similar"
                if s[1] > c[1]:
                    return "Inferior"
                else:
                    return "Superior"
            if abs(s[0] + (s[1] * 2) - c[0] + (c[1] * 2)) <= per2:
                return "Similar"
            elif s[1] > c[1]:
                return "Inferior"
            else:
                return "Superior"

    def OtherAdjuster(self, comp):

        # the largest string is considered superior

        if comp.get() == "":
            return ""
        c = comp.get()
        s = self.Subjectotherbox.get()

        if s == c:
            return ""
        elif len(s) > len(c):
            return "Inferior"
        else:
            return "Superior"

    def ConditionAdjuster(self, comp):

        # checks for various condition ratings and makes comparisons

        if comp.get() == "":
            return ""
        c = str(comp.get()).lower()
        s = str(self.Subjectconditionbox.get()).lower()

        if s == c:
            return ""
        elif ("above" in c or "update" in c or "good" in c or "+" in c or "higher" in c or "better" in c) and (
                "above" not in s or "update" not in s or "good" not in s or "+" not in s or "higher" not in s or "better" not in s):
            return "Superior"
        elif ("below" in c or "fair" in c or "poor" in c or "work" in c or "-" in c or "dated" in c) and (
                "below" not in s or "fair" not in s or "poor" not in s or "work" not in s or "-" not in s or "dated" not in s):
            return "Inferior"
        else:
            return "Similar"

    ### The following section merges the data to a PDF of the operators choice

    def CreatePDF(self):
        from PyPDF2 import PdfFileWriter, PdfFileReader  # to merge data to PDF
        from PyPDF2.generic import BooleanObject, NameObject, IndirectObject

        def set_need_appearances_writer(writer: PdfFileWriter):  # makes form fields readable after adding data.
            # without, all fields need to be manually refreshed
            try:
                catalog = writer._root_object
                # get the AcroForm tree
                if "/AcroForm" not in catalog:
                    writer._root_object.update({
                        NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)})

                need_appearances = NameObject("/NeedAppearances")
                writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
                return writer

            except Exception as e:
                print('set_need_appearances_writer() catch : ', repr(e))
                return writer

        from tkinter.filedialog import askopenfilename  # ask user for PDF to fill
        filename = askopenfilename(title="Select Eval Template",
                                   filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))

        infile = filename
        outfile = "%s eval.pdf" % (self.Subjectaddressbox.get(),)  # output filename based on address

        pdf = PdfFileReader(open(infile, "rb"), strict=False)
        if "/AcroForm" in pdf.trailer["/Root"]:
            pdf.trailer["/Root"]["/AcroForm"].update(
                {NameObject("/NeedAppearances"): BooleanObject(True)})

        pdf2 = PdfFileWriter()
        set_need_appearances_writer(pdf2)
        if "/AcroForm" in pdf2._root_object:
            pdf2._root_object["/AcroForm"].update(
                {NameObject("/NeedAppearances"): BooleanObject(True)})

        if self.adjustmentchecker.get() == 1:  # if adjustments enabled, also fill adjustment fields
            # field names must match user's PDF
            field_dictionary = {"Property Address": self.Subjectaddressbox.get(),
                                "County": self.subjectcounty,
                                "State": self.subjectstate,
                                "Zip": self.subjectzip,
                                "Assessment": self.subjectassvalue + " (%s)" % (self.subjectassyear,),
                                "City": self.subjectcity,
                                "Area Sq Feet or Acres": self.Subjectlocationbox.get().rsplit(' ', 1)[0],
                                "History sales listings offers": self.Subjectpricebox.get() + " " + self.Subjectdatebox.get(),
                                "DescriptionSp Financing": self.Subjectlocationbox.get(),
                                "DescriptionLocation": self.Subjectstylebox.get(),
                                "DescriptionYear Built": self.Subjectyearbox.get(),
                                "DescriptionCondition": self.Subjectconditionbox.get(),
                                "DescriptionArea_Sq_Feet": self.Subjectsizebox.get(),
                                "DescriptionBsmt Sq Feet": self.Subjectbasebox.get(),
                                "DescriptionBed": self.Subjectbedbox.get(),
                                "DescriptionBath": self.Subjectbathbox.get(),
                                "DescriptionGarage": self.Subjectgaragebox.get(),
                                "DescriptionOther Amenities": self.Subjectotherbox.get(),
                                "DescriptionDate of Sale": "-",
                                "DescriptionCond of Sale": "-",

                                "Comparable 1Address": self.Comp1addressbox.get(),
                                "Comp1_SalesPrice": re.sub('[^0-9]', '', self.Comp1pricebox.get()),
                                "DescriptionDate of Sale_2": self.Comp1datebox.get(),
                                "DescriptionCond of Sale_2": self.Comp1saleconbox.get(),
                                "DescriptionSp Financing_2": self.Comp1locationbox.get(),
                                "DescriptionLocation_2": self.Comp1stylebox.get(),
                                "DescriptionYear Built_2": self.Comp1yearbox.get(),
                                "DescriptionCondition_2": self.Comp1conditionbox.get(),
                                "DescriptionArea_Sq_Feet_2": self.Comp1sizebox.get(),
                                "DescriptionBsmt Sq Feet_2": self.Comp1basebox.get(),
                                "DescriptionBed_2": self.Comp1bedbox.get(),
                                "DescriptionBath_2": self.Comp1bathbox.get(),
                                "DescriptionGarage_2": self.Comp1garagebox.get(),
                                "DescriptionOther Amenities_2": self.Comp1otherbox.get(),

                                "Comp1_SpFinancing": self.LotAdjuster(self.Comp1locationbox),
                                "Comp1_Location": self.StyleAdjuster(self.Comp1stylebox),
                                "Comp1_YearBuilt": self.YearAdjuster(self.Comp1yearbox),
                                "Comp1_Area": self.GLAAdjuster(self.Comp1sizebox),
                                "Comp1_Bed": self.BedAdjuster(self.Comp1bedbox),
                                "Comp1_Bath": self.BathAdjuster(self.Comp1bathbox),
                                "Comp1_Basement": self.BaseAdjuster(self.Comp1basebox),
                                "Comp1_Garage": self.GarAdjuster(self.Comp1garagebox),
                                "Comp1_Other": self.OtherAdjuster(self.Comp1otherbox),
                                "Comp1_Condition": self.ConditionAdjuster(self.Comp1conditionbox),

                                "Comparable 2Address": self.Comp2addressbox.get(),
                                "Comp2_SalesPrice": re.sub('[^0-9]', '', self.Comp2pricebox.get()),
                                "DescriptionDate of Sale_3": self.Comp2datebox.get(),
                                "DescriptionCond of Sale_3": self.Comp2saleconbox.get(),
                                "DescriptionSp Financing_3": self.Comp2locationbox.get(),
                                "DescriptionLocation_3": self.Comp2stylebox.get(),
                                "DescriptionYear Built_3": self.Comp2yearbox.get(),
                                "DescriptionCondition_3": self.Comp2conditionbox.get(),
                                "DescriptionArea_Sq_Feet_3": self.Comp2sizebox.get(),
                                "DescriptionBsmt Sq Feet_3": self.Comp2basebox.get(),
                                "DescriptionBed_3": self.Comp2bedbox.get(),
                                "DescriptionBath_3": self.Comp2bathbox.get(),
                                "DescriptionGarage_3": self.Comp2garagebox.get(),
                                "DescriptionOther Amenities_3": self.Comp2otherbox.get(),

                                "Comp2_SpFinancing": self.LotAdjuster(self.Comp2locationbox),
                                "Comp2_Location": self.StyleAdjuster(self.Comp2stylebox),
                                "Comp2_YearBuilt": self.YearAdjuster(self.Comp2yearbox),
                                "Comp2_Area": self.GLAAdjuster(self.Comp2sizebox),
                                "Comp2_Bed": self.BedAdjuster(self.Comp2bedbox),
                                "Comp2_Bath": self.BathAdjuster(self.Comp2bathbox),
                                "Comp2_Basement": self.BaseAdjuster(self.Comp2basebox),
                                "Comp2_Garage": self.GarAdjuster(self.Comp2garagebox),
                                "Comp2_Other": self.OtherAdjuster(self.Comp2otherbox),
                                "Comp2_Condition": self.ConditionAdjuster(self.Comp2conditionbox),

                                "Comparable 3Address": self.Comp3addressbox.get(),
                                "Comp3_SalesPrice": re.sub('[^0-9]', '', self.Comp3pricebox.get()),
                                "DescriptionDate of Sale_4": self.Comp3datebox.get(),
                                "DescriptionCond of Sale_4": self.Comp3saleconbox.get(),
                                "DescriptionSp Financing_4": self.Comp3locationbox.get(),
                                "DescriptionLocation_4": self.Comp3stylebox.get(),
                                "DescriptionYear Built_4": self.Comp3yearbox.get(),
                                "DescriptionCondition_4": self.Comp3conditionbox.get(),
                                "DescriptionArea_Sq_Feet_4": self.Comp3sizebox.get(),
                                "DescriptionBsmt Sq Feet_4": self.Comp3basebox.get(),
                                "DescriptionBed_4": self.Comp3bedbox.get(),
                                "DescriptionBath_4": self.Comp3bathbox.get(),
                                "DescriptionGarage_4": self.Comp3garagebox.get(),
                                "DescriptionOther Amenities_4": self.Comp3otherbox.get(),

                                "Comp3_SpFinancing": self.LotAdjuster(self.Comp3locationbox),
                                "Comp3_Location": self.StyleAdjuster(self.Comp3stylebox),
                                "Comp3_YearBuilt": self.YearAdjuster(self.Comp3yearbox),
                                "Comp3_Area": self.GLAAdjuster(self.Comp3sizebox),
                                "Comp3_Bed": self.BedAdjuster(self.Comp3bedbox),
                                "Comp3_Bath": self.BathAdjuster(self.Comp3bathbox),
                                "Comp3_Basement": self.BaseAdjuster(self.Comp3basebox),
                                "Comp3_Garage": self.GarAdjuster(self.Comp3garagebox),
                                "Comp3_Other": self.OtherAdjuster(self.Comp3otherbox),
                                "Comp3_Condition": self.ConditionAdjuster(self.Comp3conditionbox), }
        else:
            field_dictionary = {"Property Address": self.Subjectaddressbox.get(),
                                "County": self.subjectcounty,
                                "State": self.subjectstate,
                                "Zip": self.subjectzip,
                                "Assessment": self.subjectassvalue + " (%s)" % (self.subjectassyear,),
                                "City": self.subjectcity,
                                "Area Sq Feet or Acres": self.Subjectlocationbox.get().rsplit(' ', 1)[0],
                                "History sales listings offers": self.Subjectpricebox.get() + " " + self.Subjectdatebox.get(),
                                "DescriptionSp Financing": self.Subjectlocationbox.get(),
                                "DescriptionLocation": self.Subjectstylebox.get(),
                                "DescriptionYear Built": self.Subjectyearbox.get(),
                                "DescriptionCondition": self.Subjectconditionbox.get(),
                                "DescriptionArea_Sq_Feet": self.Subjectsizebox.get().replace(",", ""),
                                "DescriptionBsmt Sq Feet": self.Subjectbasebox.get(),
                                "DescriptionBed": self.Subjectbedbox.get(),
                                "DescriptionBath": self.Subjectbathbox.get(),
                                "DescriptionGarage": self.Subjectgaragebox.get(),
                                "DescriptionOther Amenities": self.Subjectotherbox.get(),
                                "DescriptionDate of Sale": "-",
                                "DescriptionCond of Sale": "-",

                                "Comparable 1Address": self.Comp1addressbox.get(),
                                "Comp1_SalesPrice": re.sub('[^0-9]', '', self.Comp1pricebox.get()),
                                "DescriptionDate of Sale_2": self.Comp1datebox.get(),
                                "DescriptionCond of Sale_2": self.Comp1saleconbox.get(),
                                "DescriptionSp Financing_2": self.Comp1locationbox.get(),
                                "DescriptionLocation_2": self.Comp1stylebox.get(),
                                "DescriptionYear Built_2": self.Comp1yearbox.get(),
                                "DescriptionCondition_2": self.Comp1conditionbox.get(),
                                "DescriptionArea_Sq_Feet_2": self.Comp1sizebox.get(),
                                "DescriptionBsmt Sq Feet_2": self.Comp1basebox.get(),
                                "DescriptionBed_2": self.Comp1bedbox.get(),
                                "DescriptionBath_2": self.Comp1bathbox.get(),
                                "DescriptionGarage_2": self.Comp1garagebox.get(),
                                "DescriptionOther Amenities_2": self.Comp1otherbox.get(),

                                "Comparable 2Address": self.Comp2addressbox.get(),
                                "Comp2_SalesPrice": re.sub('[^0-9]', '', self.Comp2pricebox.get()),
                                "DescriptionDate of Sale_3": self.Comp2datebox.get(),
                                "DescriptionCond of Sale_3": self.Comp2saleconbox.get(),
                                "DescriptionSp Financing_3": self.Comp2locationbox.get(),
                                "DescriptionLocation_3": self.Comp2stylebox.get(),
                                "DescriptionYear Built_3": self.Comp2yearbox.get(),
                                "DescriptionCondition_3": self.Comp2conditionbox.get(),
                                "DescriptionArea_Sq_Feet_3": self.Comp2sizebox.get(),
                                "DescriptionBsmt Sq Feet_3": self.Comp2basebox.get(),
                                "DescriptionBed_3": self.Comp2bedbox.get(),
                                "DescriptionBath_3": self.Comp2bathbox.get(),
                                "DescriptionGarage_3": self.Comp2garagebox.get(),
                                "DescriptionOther Amenities_3": self.Comp2otherbox.get(),

                                "Comparable 3Address": self.Comp3addressbox.get(),
                                "Comp3_SalesPrice": re.sub('[^0-9]', '', self.Comp3pricebox.get()),
                                "DescriptionDate of Sale_4": self.Comp3datebox.get(),
                                "DescriptionCond of Sale_4": self.Comp3saleconbox.get(),
                                "DescriptionSp Financing_4": self.Comp3locationbox.get(),
                                "DescriptionLocation_4": self.Comp3stylebox.get(),
                                "DescriptionYear Built_4": self.Comp3yearbox.get(),
                                "DescriptionCondition_4": self.Comp3conditionbox.get(),
                                "DescriptionArea_Sq_Feet_4": self.Comp3sizebox.get(),
                                "DescriptionBsmt Sq Feet_4": self.Comp3basebox.get(),
                                "DescriptionBed_4": self.Comp3bedbox.get(),
                                "DescriptionBath_4": self.Comp3bathbox.get(),
                                "DescriptionGarage_4": self.Comp3garagebox.get(),
                                "DescriptionOther Amenities_4": self.Comp3otherbox.get(), }

        pdf2.addPage(pdf.getPage(0))  # copies the first page of the user's PDF
        pdf2.updatePageFormFieldValues(pdf2.getPage(0), field_dictionary)  # merges all the files on the first page
        pdf2.addPage(pdf.getPage(1))

        outputStream = open(outfile, "wb")
        pdf2.write(outputStream)  # creates PDF


if __name__ == '__main__':
    vp_start_gui()
