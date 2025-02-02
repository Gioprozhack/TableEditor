import csv
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.headings = []
        self.data = []

        self.editable = False
        self.reverse_sort = False

        self.main_menu = tk.Menu()
        self.file_menu = tk.Menu(tearoff=0)
        self.edit_menu = tk.Menu(tearoff=0)

        self.file_menu.add_command(label="New", command=self.new_file)
        self.file_menu.add_command(label="Open", command=self.open_file)
        self.file_menu.add_command(label="Save", command=self.save_file)

        self.edit_menu.add_command(label="Add column", command=self.add_column)
        self.edit_menu.add_command(label="Add row", command=self.add_row)
        self.edit_menu.add_command(label="Delete column", command=self.delete_column)
        self.edit_menu.add_command(label="Delete row", command=self.delete_row)
        self.edit_menu.add_command(label="Edit contents", command=self.edit_contents)

        self.main_menu.add_cascade(label="File", menu=self.file_menu)
        self.main_menu.add_cascade(label="Edit", menu=self.edit_menu)
        self.config(menu=self.main_menu)

        self.title("Ttk Treeview")
        self.geometry("600x400")

        self.tree = ttk.Treeview(self)
        ysb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.tree.bind("<Double-1>", lambda event: self.edit_cell(event))

        self.tree.grid(row=1, column=0, columnspan=2, sticky="nsew")
        ysb.grid(row=1, column=1, sticky="ns")
        xsb.grid(row=2, column=0, sticky="ew")

        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

    def update_table(self):

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = self.headings
        if self.editable:
            self.tree["show"] = ""
        else:
            self.tree["show"] = "headings"

        for h in self.headings:
            self.tree.heading(h, text=h, command=lambda i=self.headings.index(h): self.sort_column(i))
            self.tree.column(h, width=100, anchor="center")

        for d in self.data:
            self.tree.insert("", tk.END, values=d)

    def new_file(self):

        self.headings.clear()
        self.data.clear()
        self.update_table()

    def open_file(self):

        file_path = tk.filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        file_contents = []
        try:
            with open(file_path, newline="", encoding="utf-8") as f:
                for line in csv.reader(f, delimiter=";"):
                    file_contents.append(line)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при чтении файла:\n{e}")
            return

        if file_contents:
            self.headings = file_contents[0]
            self.data = file_contents[1:] if len(file_contents) > 1 else []
            self.update_table()

    def save_file(self):

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not file_path:
            return

        try:
            with open(file_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=";")
                if self.headings:
                    writer.writerow(self.headings)
                writer.writerows(self.data)
            messagebox.showinfo("Сохранение", "Файл успешно сохранён!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении файла:\n{e}")

    def add_column(self):

        if not self.editable:
            column_name = f"Column {len(self.headings) + 1}"
            self.headings.append(column_name)
            if self.data:
                for row in self.data:
                    row.append("")
            else:
                self.add_row()
            self.update_table()

    def add_row(self):

        if not self.editable:
            if self.headings:
                self.data.append(["" for h in self.headings])
            else:
                self.add_column()
            self.update_table()

    def delete_column(self):

        if not self.editable:
            if self.headings:
                self.headings.pop()
                for row in self.data:
                    if row:
                        row.pop()
            self.update_table()

    def delete_row(self):

        if not self.editable:
            if self.data:
                self.data.pop()
            self.update_table()

    def edit_cell(self, event):

        if self.editable:
            selected_item = self.tree.focus()
            if not selected_item:
                return

            col_id = self.tree.identify_column(event.x)
            col_index = int(col_id[1:]) - 1
            row_index = self.tree.index(selected_item)

            if col_index < 0 or row_index < 0:
                return

            current_value = self.data[row_index][col_index]

            x, y, width, height = self.tree.bbox(selected_item, col_id)

            entry = tk.Entry(self.tree, justify="center")
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, current_value)
            entry.focus()

            def save_edit(event=None):
                new_value = entry.get()
                self.data[row_index][col_index] = new_value
                self.update_table()
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)
        else:
            return

    def edit_contents(self):

        if not self.editable:
            self.data = [self.headings] + self.data
        else:
            self.headings = self.data[0]
            self.data = self.data[1:]

        self.editable = not self.editable
        self.update_table()

    def sort_column(self, col):

        sorted_data = {}

        for row in self.data:
            sorted_data[row[col]] = row
        sorted_data = dict(sorted(sorted_data.items(), reverse=self.reverse_sort))

        self.reverse_sort = not self.reverse_sort
        self.data.clear()

        for key in sorted_data.keys():
            self.data.append(sorted_data[key])

        self.update_table()


if __name__ == "__main__":
    app = App()
    app.mainloop()
