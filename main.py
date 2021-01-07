from tkinter import Tk, Entry, Label, Button, filedialog, StringVar, Text, messagebox, DISABLED, NORMAL

from PIL import Image
from pandas import DataFrame, ExcelWriter
from random import randint


def make_plan():
    global size_page, nbr_page, excel_name, shape_name, submit_btn, query_btn, root
    submit_btn['state'] = DISABLED
    query_btn['state'] = DISABLED

    size_page = int(size_page.get()) if type(size_page) != int else size_page
    nbr_page = int(nbr_page.get()) if type(nbr_page) != int else size_page

    shape = Image.open(shape_name)
    shape = shape.convert('1')
    width, height = shape.size
    shape.show()

    smallest_r = []
    for i in range(0, 35):
        smallest_r.append(width % (nbr_page - i))
    empty_page = smallest_r.index(min(smallest_r))
    leaf_size = width // (nbr_page - empty_page)

    smallest_r = []
    for i in range(20, 80):
        smallest_r.append(height % (size_page - i))
    empty_cm = 20 + smallest_r.index(min(smallest_r))
    mm_size = height // (size_page - empty_cm)

    plan = {}
    plan_exact = {}

    x_count = 0
    for page in range(0, nbr_page):
        page_final = []

        if page <= empty_page/2:
            page_plan = [0] * (size_page+1)
        elif page >= nbr_page - (empty_page / 2):
            page_plan = [0] * (size_page+1)
        else:
            page_plan = []

            i = 0
            y_count = 0
            while i <= size_page and y_count <= height:
                block = []
                if i <= empty_cm/2:
                    block = [1] * mm_size
                elif i >= size_page - (empty_cm/2):
                    block = [1] * mm_size
                else:
                    for x in range(0, leaf_size):
                        x_pix = x_count + x if (x_count + x) < (width - 2) else width - 2
                        for y in range(0, mm_size):
                            y_pix = y_count + y if (y_count + y) < (height - 2) else height - 2
                            block.append(shape.getpixel((x_pix, y_pix)))
                        y_count = y_count + mm_size
                i = i + 1
                page_plan.append(i if 0 in block else 0)
            x_count = x_count + leaf_size
        print(len(page_plan))
        i = 1
        while i < len(page_plan):
            if page_plan[i] != 0:
                if i == 1:
                    page_final.append((page_plan[i] - 1) / 10)
                elif page_plan[i - 1] == 0:
                    page_final.append((page_plan[i] - 1) / 10)
                elif i == len(page_plan) - 1:
                    page_final.append((page_plan[i] - 1) / 10)
            elif page_plan[i] == 0:
                if page_plan[i - 1] != 0:
                    page_final.append((page_plan[i - 1] - 1) / 10)
            i = i + 1

        plan_exact.update({page + 1: page_plan})
        plan.update({page + 1: page_final})
    plan = DataFrame.from_dict(plan, orient='index')
    plan_exact = DataFrame.from_dict(plan_exact)

    with ExcelWriter(str(excel_name.get())) as writer:
        plan.to_excel(excel_writer=writer, sheet_name='Plan')
        plan_exact.to_excel(excel_writer=writer, sheet_name='Exact_measurement')

    if messagebox.askyesno(title="Plan généré avec succès", message="L'excel a été créé avec succès ! "
                                                                    "Réaliser un autre plan ?"):
        submit_btn['state'] = NORMAL
        query_btn['state'] = NORMAL
        excel_name.delete(0, 'end')
    else:
        root.destroy()


def callback():
    global shape_name
    shape_name = filedialog.askopenfilename()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    root = Tk()
    root.title("Plans app")
    root.geometry("420x150")

    # create text boxes
    x1 = StringVar()
    x1.set('plan.xlsx')

    size_page = Entry(root, width=10)
    size_page.grid(row=0, column=1)
    nbr_page = Entry(root, width=10)
    nbr_page.grid(row=1, column=1)
    excel_name = Entry(root, width=10, textvariable=x1)
    excel_name.grid(row=2, column=1)

    # create label
    size_label = Label(root, text='Taille page du livre en mm')
    size_label.grid(row=0, column=0)

    nbr_label = Label(root, text='Nombre de feuille composant le livre')
    nbr_label.grid(row=1, column=0)

    section_label = Label(root, text='Nom du fichier excel de sortie (format : nom.xlsx)')
    section_label.grid(row=2, column=0)

    # create submit button
    warning = Label(root, text='Pas de fond transparent !', fg='red')
    warning.grid(row=3, column=0)

    shape_name = ""
    submit_btn = Button(root, text='Choisir le modèle de départ', command=callback)
    submit_btn.grid(row=4, column=0, columnspan=2)

    # create a query button
    query_btn = Button(root, text='Build plan', command=make_plan)
    query_btn.grid(row=4, column=1, columnspan=2)

    root.mainloop()
