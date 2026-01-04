import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

CHUNK_SIZE = 5000

class CSVExcelViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV / Excel Viewer")
        self.root.geometry("1250x720")

        
        top = tk.Frame(root)
        top.pack(fill="x", pady=4)

        tk.Button(top, text="Open CSV/Excel", command=self.open_file).pack(side="left", padx=4)
        tk.Button(top, text="◀ Prev", command=self.prev_chunk).pack(side="left")
        tk.Button(top, text="Next ▶", command=self.next_chunk).pack(side="left", padx=4)

        tk.Label(top, text="Search:").pack(side="left", padx=10)
        self.search_entry = tk.Entry(top, width=25)
        self.search_entry.pack(side="left")
        tk.Button(top, text="Find", command=self.search).pack(side="left", padx=4)

        self.info = tk.Label(top, text="No file loaded")
        self.info.pack(side="right", padx=10)

        
        frame = tk.Frame(root)
        frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, show="headings", selectmode="extended")
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.tag_configure("odd", background="#f0f0f0")
        self.tree.tag_configure("even", background="#ffffff")
        self.tree.tag_configure("found", background="#ffeb3b")

        
        self.root.bind_all("<KeyPress>", self.global_key_handler)

        self.reader = None
        self.chunks = []
        self.page = 0
        self.file = None
        self.filetype = "csv"
        self.search_results = []

        
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="Copy", command=self.copy_selection)
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        try:
            self.tree_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_menu.grab_release()

    
    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV / Excel files", "*.csv *.xlsx")])
        if not path:
            return

        self.file = path
        if path.endswith(".xlsx"):
            self.filetype = "excel"
            self.reader = pd.read_excel(path, sheet_name=None)
            self.sheet_names = list(self.reader.keys())
            self.current_sheet_index = 0
            self.load_sheet(self.sheet_names[self.current_sheet_index])
        else:
            self.filetype = "csv"
            self.reader = pd.read_csv(path, chunksize=CHUNK_SIZE, iterator=True, low_memory=False)
            self.chunks = []
            self.page = 0
            self.load_chunk()

    def load_sheet(self, sheet):
        df = self.reader[sheet]
        self.chunks = [df]
        self.page = 0
        self.show_table(df)
        self.info.config(text=f"Sheet: {sheet}, Rows 1 – {len(df)}")

    def load_chunk(self):
        try:
            while len(self.chunks) <= self.page:
                self.chunks.append(next(self.reader))
            self.show_table(self.chunks[self.page])
            self.info.config(text=f"Rows {self.page*CHUNK_SIZE + 1} – {(self.page+1)*CHUNK_SIZE}")
        except StopIteration:
            messagebox.showinfo("End", "End of file")

    def show_table(self, df):
        self.tree.delete(*self.tree.get_children())

        cols = ["#"] + list(df.columns)
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")

        start_num = self.page * CHUNK_SIZE + 1

        for i, (_, row) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            values = [start_num + i] + list(row)
            self.tree.insert("", "end", values=values, tags=(tag,))

   
    def next_chunk(self):
        if self.filetype == "csv" and self.reader:
            self.page += 1
            try:
                if self.page < len(self.chunks):
                    self.show_table(self.chunks[self.page])
                else:
                    self.chunks.append(next(self.reader))
                    self.show_table(self.chunks[self.page])
                self.info.config(text=f"Rows {self.page*CHUNK_SIZE + 1} – {(self.page+1)*CHUNK_SIZE}")
            except StopIteration:
                self.page -= 1
                messagebox.showinfo("End", "End of file")

    def prev_chunk(self):
        if self.page > 0:
            self.page -= 1
            self.show_table(self.chunks[self.page])
            self.info.config(text=f"Rows {self.page*CHUNK_SIZE + 1} – {(self.page+1)*CHUNK_SIZE}")

    
    def search(self):
        query = self.search_entry.get().lower()
        if not query or not self.file:
            return

        self.search_results = []

        if self.filetype == "csv":
            reader = pd.read_csv(self.file, chunksize=CHUNK_SIZE, low_memory=False)
            for chunk in reader:
                mask = chunk.astype(str).apply(lambda r: r.str.lower().str.contains(query, na=False)).any(axis=1)
                if mask.any():
                    df_res = chunk[mask].copy()
                    df_res["_row_num"] = df_res.index + 1
                    self.search_results.append(df_res)
        else:
            for sheet, df in self.reader.items():
                mask = df.astype(str).apply(lambda r: r.str.lower().str.contains(query, na=False)).any(axis=1)
                if mask.any():
                    df_res = df[mask].copy()
                    df_res["_row_num"] = df_res.index + 1
                    df_res["_sheet"] = sheet
                    self.search_results.append(df_res)

        if not self.search_results:
            messagebox.showinfo("Search", "Nothing found")
            return

        self.show_search_results_window(query)
        self.highlight_all(query)

    def highlight_all(self, query):
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            if any(query in str(v).lower() for v in values):
                self.tree.item(item, tags=("found",))

    
    def show_search_results_window(self, query):
        win = tk.Toplevel(self.root)
        win.title(f"Search results for '{query}'")
        win.geometry("1000x500")

        frame = tk.Frame(win)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(frame, show="headings", selectmode="extended")
        tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        tree.configure(yscrollcommand=vsb.set)

        sample_df = self.search_results[0]
        cols = ["Row #"] + [c for c in sample_df.columns if c not in ["_row_num", "_sheet"]]
        tree["columns"] = cols

        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=130, anchor="center")

        for df in self.search_results:
            for _, row in df.iterrows():
                rownum = row["_row_num"]
                values = [rownum] + [row[c] for c in cols if c != "Row #"]
                tree.insert("", "end", values=values)

       
        menu = tk.Menu(win, tearoff=0)
        menu.add_command(label="Copy", command=lambda: self.copy_from_tree(tree, win))
        tree.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

    
        tree.bind("<Control-KeyPress-c>", lambda e: self.copy_from_tree(tree, win))
        tree.bind("<Control-KeyPress-C>", lambda e: self.copy_from_tree(tree, win))

    def copy_from_tree(self, tree, win):
        items = tree.selection()
        if not items:
            return
        data = []
        for item in items:
            row = tree.item(item, "values")
            data.append("\t".join(map(str, row)))
        win.clipboard_clear()
        win.clipboard_append("\n".join(data))
        win.update_idletasks()

    
    def global_key_handler(self, event):
        if (event.state & 0x4) and event.keysym.lower() in ["c", "с"]:
            self.copy_selection()
            return "break"

   
    def copy_selection(self):
        try:
            items = self.tree.selection()
            if not items:
                return
            data = []
            for item in items:
                row = self.tree.item(item, "values")
                data.append("\t".join(map(str, row)))
            text = "\n".join(data)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update_idletasks()
        except:
            pass


if __name__ == "__main__":
    root = tk.Tk()
    CSVExcelViewer(root)
    root.mainloop()
