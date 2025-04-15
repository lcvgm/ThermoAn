import os
import re
import time
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
from tkinter import filedialog, Tk, Button, Label, Toplevel, BooleanVar, Checkbutton
from tkinter import ttk
from tkinter.messagebox import showinfo, askyesno
import win32com.client  # Origin COM API

# Для отображения графиков в Tkinter с панелью навигации
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

# Постоянная Больцмана (в эВ/К)
k_B = 8.617e-5

# Глобальная переменная для смещения интерактивных окон
interactive_counter = 0


def get_folder_from_longname(longname):
    folder = os.path.dirname(longname)
    if not folder or folder.strip() == "":
        folder = "<Root>"
    return folder


def get_folder_info(origin):
    folders_info = {}
    total_pages = int(origin.WorksheetPages.Count)
    for i in range(total_pages):
        page = origin.WorksheetPages(i)
        if page.Layers.Count == 0:
            continue
        book_fullname = page.LongName if hasattr(page, "LongName") and page.LongName else page.Name
        folder_part = os.path.dirname(book_fullname)
        if not folder_part or folder_part.strip() == "":
            folder_part = get_folder_from_longname(book_fullname)
        if not folder_part or folder_part.strip() == "":
            folder_part = "<Root>"
        base_book_name = os.path.basename(book_fullname)
        folders_info.setdefault(folder_part, []).append((i, base_book_name))
    return folders_info


def select_books_and_options(folders_info):
    top = Toplevel(root)
    top.title("Select Books and Options")
    top.geometry("400x350")
    top.transient(root)
    top.grab_set()

    process_all_var = BooleanVar(value=False)
    chk_all = Checkbutton(top, text="Process All Books", variable=process_all_var)
    chk_all.pack(anchor="w", padx=10, pady=5)

    interactive_var = BooleanVar(value=False)
    chk_interactive = Checkbutton(top, text="Enable Interactive Exclusion", variable=interactive_var)
    chk_interactive.pack(anchor="w", padx=10, pady=5)

    tree = ttk.Treeview(top, selectmode="extended")
    tree["columns"] = ("BookIndex",)
    tree.column("#0", width=300, anchor="w")
    tree.heading("#0", text="Folder / Book (LongName)\n(Hold Ctrl for multiple)")
    tree.column("BookIndex", width=80, anchor="center")
    tree.heading("BookIndex", text="Index")
    tree.pack(fill="both", expand=True, padx=10, pady=5)

    for folder_path, books in folders_info.items():
        parent_id = tree.insert("", "end", text=folder_path, open=True)
        for (book_index, book_name) in books:
            tree.insert(parent_id, "end", text=book_name, values=(book_index,))

    def confirm():
        selected_books = []
        if process_all_var.get():
            for parent in tree.get_children():
                for child in tree.get_children(parent):
                    val = tree.item(child, "values")
                    if val:
                        selected_books.append(int(val[0]))
        else:
            for item in tree.selection():
                children = tree.get_children(item)
                if children:
                    for child in children:
                        val = tree.item(child, "values")
                        if val:
                            selected_books.append(int(val[0]))
                else:
                    val = tree.item(item, "values")
                    if val:
                        selected_books.append(int(val[0]))
        top.selected_books = selected_books
        top.interactive = interactive_var.get()
        top.destroy()

    Button(top, text="Confirm Selection", command=confirm).pack(pady=10)
    top.wait_window()
    return (top.selected_books if hasattr(top, "selected_books") else [],
            top.interactive if hasattr(top, "interactive") else False)


def update_axis(ax, x, y, primary_mask, dual_mask, label, dual_enabled):
    for ln in ax.lines[:]:
        ln.remove()
    if dual_enabled:
        primary_line = None
        dual_line = None
        if np.sum(primary_mask) >= 2:
            reg = linregress(x[primary_mask], y[primary_mask])
            Ea_primary = reg.slope * k_B
            x_sorted = np.sort(x[primary_mask])
            primary_line, = ax.plot(x_sorted, reg.intercept + reg.slope * x_sorted, 'b-',
                                    label=f"{label} Primary: slope={reg.slope:.3f}, intercept={reg.intercept:.3f}, r²={reg.rvalue ** 2:.3f}, Ea={Ea_primary:.3f} eV")
            coeff_primary = (reg.slope, reg.intercept, reg.rvalue ** 2)
        else:
            ax.set_title(f"{label} Primary: Not enough points")
            coeff_primary = (None, None, None)
        if np.sum(dual_mask) >= 2:
            reg2 = linregress(x[dual_mask], y[dual_mask])
            Ea_dual = reg2.slope * k_B
            x_sorted2 = np.sort(x[dual_mask])
            dual_line, = ax.plot(x_sorted2, reg2.intercept + reg2.slope * x_sorted2, 'r--',
                                 label=f"{label} Dual: slope={reg2.slope:.3f}, intercept={reg2.intercept:.3f}, r²={reg2.rvalue ** 2:.3f}, Ea={Ea_dual:.3f} eV")
        else:
            ax.set_title(f"{label} Dual: Not enough points")
        return coeff_primary + (primary_line,)
    else:
        if np.sum(primary_mask) >= 2:
            reg = linregress(x[primary_mask], y[primary_mask])
            Ea = reg.slope * k_B
            x_sorted = np.sort(x[primary_mask])
            line, = ax.plot(x_sorted, reg.intercept + reg.slope * x_sorted, 'b-',
                            label=f"{label}: slope={reg.slope:.3f}, intercept={reg.intercept:.3f}, r²={reg.rvalue ** 2:.3f}, Ea={Ea:.3f} eV")
            return (reg.slope, reg.intercept, reg.rvalue ** 2, line)
        else:
            ax.set_title(f"{label}: Not enough points")
            return (None, None, None, None)


def interactive_edit_all_models(df_reg, longname):
    global interactive_counter
    win = Toplevel(root)
    win.title("Interactive Editing for All Models")
    x_offset = 100 + interactive_counter * 50
    y_offset = 100 + interactive_counter * 50
    win.geometry(f"+{x_offset}+{y_offset}")
    interactive_counter += 1

    sync_exclusion_var = BooleanVar(value=True)
    sync_chk = Checkbutton(win, text="Sync Exclusion Across Graphs", variable=sync_exclusion_var)
    sync_chk.pack(pady=5)

    dual_trend_var = BooleanVar(value=False)
    dual_trend_chk = Checkbutton(win, text="Enable Dual Trend Mode", variable=dual_trend_var)
    dual_trend_chk.pack(pady=5)

    fig, axes = plt.subplots(3, 1, figsize=(8, 16))

    x_arr = df_reg["1/T (1/K)"].values
    x_half = df_reg["invT_half"].values
    x_quarter = df_reg["invT_quarter"].values
    y_data = df_reg["ln(R)"].values
    N = len(x_arr)
    # Всегда создаём оба набора масок
    mask_primary_arr = np.ones(N, dtype=bool)
    mask_dual_arr = np.zeros(N, dtype=bool)
    mask_primary_half = np.ones(N, dtype=bool)
    mask_dual_half = np.zeros(N, dtype=bool)
    mask_primary_quarter = np.ones(N, dtype=bool)
    mask_dual_quarter = np.zeros(N, dtype=bool)

    # Начальные цвета scatter-плотов в зависимости от состояния dual режима
    if dual_trend_var.get():
        colors_arr = ['red' if v else 'green' for v in mask_dual_arr]
        colors_half = ['red' if v else 'green' for v in mask_dual_half]
        colors_quarter = ['red' if v else 'green' for v in mask_dual_quarter]
    else:
        colors_arr = ['green' if v else 'red' for v in mask_primary_arr]
        colors_half = ['green' if v else 'red' for v in mask_primary_half]
        colors_quarter = ['green' if v else 'red' for v in mask_primary_quarter]

    sc_arr = axes[0].scatter(x_arr, y_data, c=colors_arr, picker=5)
    sc_half = axes[1].scatter(x_half, y_data, c=colors_half, picker=5)
    sc_quarter = axes[2].scatter(x_quarter, y_data, c=colors_quarter, picker=5)

    def update_current_axis(ax, x, y, primary_mask, dual_mask, label):
        return update_axis(ax, x, y, primary_mask, dual_mask, label, dual_trend_var.get())

    slope_arr, intercept_arr, r2_arr, primary_line_arr = update_current_axis(axes[0], x_arr, y_data, mask_primary_arr,
                                                                             mask_dual_arr, "Arrhenius Model")
    slope_half, intercept_half, r2_half, primary_line_half = update_current_axis(axes[1], x_half, y_data,
                                                                                 mask_primary_half, mask_dual_half,
                                                                                 "(1/T)^0.5 Model")
    slope_quarter, intercept_quarter, r2_quarter, primary_line_quarter = update_current_axis(axes[2], x_quarter, y_data,
                                                                                             mask_primary_quarter,
                                                                                             mask_dual_quarter,
                                                                                             "(1/T)^0.25 Model")

    fig.tight_layout()

    def onpick(event):
        nonlocal slope_arr, intercept_arr, r2_arr, slope_half, intercept_half, r2_half, slope_quarter, intercept_quarter, r2_quarter
        for idx, ax in enumerate(axes):
            if event.artist in ax.collections:
                ind = event.ind[0]
                if dual_trend_var.get():
                    if sync_exclusion_var.get():
                        new_val = not mask_dual_arr[ind]
                        mask_dual_arr[ind] = new_val
                        mask_dual_half[ind] = new_val
                        mask_dual_quarter[ind] = new_val
                    else:
                        if idx == 0:
                            mask_dual_arr[ind] = not mask_dual_arr[ind]
                        elif idx == 1:
                            mask_dual_half[ind] = not mask_dual_half[ind]
                        elif idx == 2:
                            mask_dual_quarter[ind] = not mask_dual_quarter[ind]
                else:
                    if sync_exclusion_var.get():
                        new_val = not mask_primary_arr[ind]
                        mask_primary_arr[ind] = new_val
                        mask_primary_half[ind] = new_val
                        mask_primary_quarter[ind] = new_val
                    else:
                        if idx == 0:
                            mask_primary_arr[ind] = not mask_primary_arr[ind]
                        elif idx == 1:
                            mask_primary_half[ind] = not mask_primary_half[ind]
                        elif idx == 2:
                            mask_primary_quarter[ind] = not mask_primary_quarter[ind]
                break

        if dual_trend_var.get():
            sc_arr.set_color(['red' if v else 'green' for v in mask_dual_arr])
            sc_half.set_color(['red' if v else 'green' for v in mask_dual_half])
            sc_quarter.set_color(['red' if v else 'green' for v in mask_dual_quarter])
            res0 = update_current_axis(axes[0], x_arr, y_data, mask_primary_arr, mask_dual_arr, "Arrhenius Model")
            res1 = update_current_axis(axes[1], x_half, y_data, mask_primary_half, mask_dual_half, "(1/T)^0.5 Model")
            res2 = update_current_axis(axes[2], x_quarter, y_data, mask_primary_quarter, mask_dual_quarter,
                                       "(1/T)^0.25 Model")
            if res0[0] is not None:
                slope_arr, intercept_arr, r2_arr, _ = res0
            if res1[0] is not None:
                slope_half, intercept_half, r2_half, _ = res1
            if res2[0] is not None:
                slope_quarter, intercept_quarter, r2_quarter, _ = res2
        else:
            sc_arr.set_color(['green' if v else 'red' for v in mask_primary_arr])
            sc_half.set_color(['green' if v else 'red' for v in mask_primary_half])
            sc_quarter.set_color(['green' if v else 'red' for v in mask_primary_quarter])
            res0 = update_current_axis(axes[0], x_arr, y_data, mask_primary_arr, None, "Arrhenius Model")
            res1 = update_current_axis(axes[1], x_half, y_data, mask_primary_half, None, "(1/T)^0.5 Model")
            res2 = update_current_axis(axes[2], x_quarter, y_data, mask_primary_quarter, None, "(1/T)^0.25 Model")
            if res0[0] is not None:
                slope_arr, intercept_arr, r2_arr, _ = res0
            if res1[0] is not None:
                slope_half, intercept_half, r2_half, _ = res1
            if res2[0] is not None:
                slope_quarter, intercept_quarter, r2_quarter, _ = res2

        # Обновляем легенды для каждого субплота с заголовком longname
        for ax in axes:
            ax.legend(title=longname)
        fig.canvas.draw()

    fig.canvas.mpl_connect('pick_event', onpick)

    for ax in axes:
        ax.legend(title=longname)

    canvas = FigureCanvasTkAgg(fig, master=win)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)
    toolbar = NavigationToolbar2Tk(canvas, win)
    toolbar.update()
    Button(win, text="Done", command=win.destroy).pack(pady=10)
    win.wait_window()

    plt.close(fig)

    return (slope_arr, intercept_arr, r2_arr,
            slope_half, intercept_half, r2_half,
            slope_quarter, intercept_quarter, r2_quarter)


def perform_direct_regression(df_reg):
    reg_arr = linregress(df_reg["1/T (1/K)"], df_reg["ln(R)"])
    slope_arr = reg_arr.slope
    intercept_arr = reg_arr.intercept
    r2_arr = reg_arr.rvalue ** 2

    reg_half = linregress(df_reg["invT_half"], df_reg["ln(R)"])
    slope_half = reg_half.slope
    intercept_half = reg_half.intercept
    r2_half = reg_half.rvalue ** 2

    reg_quarter = linregress(df_reg["invT_quarter"], df_reg["ln(R)"])
    slope_quarter = reg_quarter.slope
    intercept_quarter = reg_quarter.intercept
    r2_quarter = reg_quarter.rvalue ** 2

    return (slope_arr, intercept_arr, r2_arr,
            slope_half, intercept_half, r2_half,
            slope_quarter, intercept_quarter, r2_quarter)


def process_file():
    filepath = filedialog.askopenfilename(
        title="Select CSV, TXT or Origin File",
        filetypes=[("Supported files", "*.csv *.txt *.dat *.opj *.opju")]
    )
    if not filepath:
        return
    ext = os.path.splitext(filepath)[1].lower()
    result_prefix = os.path.splitext(filepath)[0]
    results = []
    composite_data = {"Arrhenius": [], "(1/T)^0.5": [], "(1/T)^0.25": []}

    if ext in [".opj", ".opju"]:
        try:
            try:
                origin = win32com.client.GetActiveObject("Origin.ApplicationSI")
                print("[INFO] Using active Origin instance.")
            except Exception:
                origin = win32com.client.Dispatch("Origin.ApplicationSI")
                print("[INFO] Started new Origin instance.")
            origin.Visible = True
            if origin.WorksheetPages.Count == 0:
                origin.Execute(f'doc -o "{filepath}"')
                time.sleep(3)
            total_pages = int(origin.WorksheetPages.Count)
            print(f"[INFO] Found {total_pages} worksheet pages in Origin file.")

            folders_info = get_folder_info(origin)
            if not folders_info:
                showinfo("Error", "No valid books with worksheets found.")
                return

            selected_books, interactive_flag = select_books_and_options(folders_info)
            if not selected_books:
                showinfo("Error", "No books selected.")
                return

            for book_index in selected_books:
                page = origin.WorksheetPages(book_index)
                if page.Layers.Count == 0:
                    continue
                wks = page.Layers(0)
                book_longname = page.LongName if (hasattr(page, "LongName") and page.LongName) else page.Name
                print(f"[INFO] Processing book {book_index}: {book_longname}")
                columns_data = []
                col_names = []
                for j in range(wks.Columns.Count):
                    col = wks.Columns(j)
                    name = col.LongName if (
                                col is not None and hasattr(col, "LongName") and col.LongName) else f"Col{j + 1}"
                    try:
                        if col is None:
                            continue
                        raw = col.GetData(5, 0)
                        values = raw
                        cleaned = []
                        valid_values = 0
                        for v in values:
                            try:
                                val = float(str(v).strip())
                                cleaned.append(val)
                                valid_values += 1
                            except (ValueError, TypeError):
                                cleaned.append(np.nan)
                        if valid_values > 0:
                            col_names.append(name)
                            columns_data.append(cleaned)
                    except Exception as e:
                        print(f"Warning: Could not read column {name}: {e}")
                if not columns_data:
                    print(f"[WARN] No valid data in book {book_index}. Skipping.")
                    continue
                max_len = max(len(col) for col in columns_data)
                for i in range(len(columns_data)):
                    while len(columns_data[i]) < max_len:
                        columns_data[i].append(np.nan)
                df = pd.DataFrame({name: data for name, data in zip(col_names, columns_data)})
                df = df.replace({"": np.nan})
                required = ["T, K", "Resistivity, Ohm*cm", "Bulk Con, cm^-3"]
                if any(req not in df.columns for req in required):
                    print(f"[WARN] Book {book_index} missing required columns. Skipping.")
                    continue

                df_reg = df.dropna(subset=["T, K", "Resistivity, Ohm*cm"]).copy()
                df_reg["1/T (1/K)"] = 1 / df_reg["T, K"]
                df_reg["ln(R)"] = np.log(df_reg["Resistivity, Ohm*cm"])
                df_reg["invT_half"] = df_reg["1/T (1/K)"] ** 0.5
                df_reg["invT_quarter"] = df_reg["1/T (1/K)"] ** 0.25

                if interactive_flag:
                    (slope_arr, intercept_arr, r2_arr,
                     slope_half, intercept_half, r2_half,
                     slope_quarter, intercept_quarter, r2_quarter) = interactive_edit_all_models(df_reg, book_longname)
                else:
                    (slope_arr, intercept_arr, r2_arr,
                     slope_half, intercept_half, r2_half,
                     slope_quarter, intercept_quarter, r2_quarter) = perform_direct_regression(df_reg)

                activation_energy_arr = slope_arr * k_B
                activation_energy_half = slope_half * k_B
                activation_energy_quarter = slope_quarter * k_B

                df_reg["Fit_Arrhenius"] = intercept_arr + slope_arr * (1 / df_reg["T, K"])
                df_reg["Fit_Half"] = intercept_half + slope_half * ((1 / df_reg["T, K"]) ** 0.5)
                df_reg["Fit_Quarter"] = intercept_quarter + slope_quarter * ((1 / df_reg["T, K"]) ** 0.25)

                summary = {
                    "Book_Index": book_index,
                    "Book_Name": book_longname,
                    "Slope_Arr": slope_arr,
                    "Intercept_Arr": intercept_arr,
                    "R2_Arr": r2_arr,
                    "Ea_Arr": activation_energy_arr,
                    "Slope_Half": slope_half,
                    "Intercept_Half": intercept_half,
                    "R2_Half": r2_half,
                    "Ea_Half": activation_energy_half,
                    "Slope_Quarter": slope_quarter,
                    "Intercept_Quarter": intercept_quarter,
                    "R2_Quarter": r2_quarter,
                    "Ea_Quarter": activation_energy_quarter
                }
                sheet_name = f"Book{book_index}_{book_longname}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                results.append({
                    "sheet_name": sheet_name,
                    "df": df,
                    "df_reg": df_reg,
                    "summary": summary,
                    "longname": book_longname
                })
                composite_data["Arrhenius"].append(
                    (df_reg["T, K"].values, df_reg["Fit_Arrhenius"].values, activation_energy_arr, book_longname))
                composite_data["(1/T)^0.5"].append(
                    (df_reg["T, K"].values, df_reg["Fit_Half"].values, activation_energy_half, book_longname))
                composite_data["(1/T)^0.25"].append(
                    (df_reg["T, K"].values, df_reg["Fit_Quarter"].values, activation_energy_quarter, book_longname))

            if len(selected_books) > 1:
                comp_win = Toplevel(root)
                comp_win.title("Composite Fit Charts")
                comp_fig, comp_axes = plt.subplots(3, 1, figsize=(8, 12))
                for label, comp_data in composite_data.items():
                    for t_vals, fit_vals, ea, ln in comp_data:
                        comp_label = f"Ea={ea:.3f} eV: {ln}"
                        if label == "Arrhenius":
                            comp_axes[0].plot(t_vals, fit_vals, '-', label=comp_label)
                        elif label == "(1/T)^0.5":
                            comp_axes[1].plot(t_vals, fit_vals, '-', label=comp_label)
                        elif label == "(1/T)^0.25":
                            comp_axes[2].plot(t_vals, fit_vals, '-', label=comp_label)
                comp_axes[0].set_title("Composite Arrhenius Dependency")
                comp_axes[1].set_title("Composite (1/T)^0.5 Dependency")
                comp_axes[2].set_title("Composite (1/T)^0.25 Dependency")
                for ax in comp_axes:
                    ax.legend()
                    ax.grid(True)
                comp_canvas = FigureCanvasTkAgg(comp_fig, master=comp_win)
                comp_canvas.draw()
                comp_canvas.get_tk_widget().pack(fill="both", expand=True)
                comp_toolbar = NavigationToolbar2Tk(comp_canvas, comp_win)
                comp_toolbar.update()
                comp_canvas.get_tk_widget().pack(fill="both", expand=True)
        except Exception as e:
            showinfo("Error", f"Error processing Origin file: {e}")
            return
    else:
        try:
            if ext == ".csv":
                df = pd.read_csv(filepath)
            elif ext in [".txt", ".dat"]:
                df = pd.read_csv(filepath, delimiter='\t', engine='python')
            else:
                showinfo("Error", "Unsupported file format.")
                return
            result_prefix = os.path.splitext(filepath)[0]
            results = [{
                "sheet_name": "Data",
                "df": df,
                "df_reg": df.dropna(subset=["T, K", "Resistivity, Ohm*cm"]).copy(),
                "longname": ""
            }]
            df_reg = results[0]["df_reg"]
            df_reg["1/T (1/K)"] = 1 / df_reg["T, K"]
            df_reg["ln(R)"] = np.log(df_reg["Resistivity, Ohm*cm"])
            df_reg["invT_half"] = df_reg["1/T (1/K)"] ** 0.5
            df_reg["invT_quarter"] = df_reg["1/T (1/K)"] ** 0.25

            (slope_arr, intercept_arr, r2_arr,
             slope_half, intercept_half, r2_half,
             slope_quarter, intercept_quarter, r2_quarter) = perform_direct_regression(df_reg)

            df_reg["Fit_Arrhenius"] = intercept_arr + slope_arr * (1 / df_reg["T, K"])
            df_reg["Fit_Half"] = intercept_half + slope_half * ((1 / df_reg["T, K"]) ** 0.5)
            df_reg["Fit_Quarter"] = intercept_quarter + slope_quarter * ((1 / df_reg["T, K"]) ** 0.25)

            summary = {
                "Slope_Arr": slope_arr,
                "Intercept_Arr": intercept_arr,
                "R2_Arr": r2_arr,
                "Ea_Arr": slope_arr * k_B,
                "Slope_Half": slope_half,
                "Intercept_Half": intercept_half,
                "R2_Half": r2_half,
                "Ea_Half": slope_half * k_B,
                "Slope_Quarter": slope_quarter,
                "Intercept_Quarter": intercept_quarter,
                "R2_Quarter": r2_quarter,
                "Ea_Quarter": slope_quarter * k_B
            }
            results[0]["summary"] = summary
        except Exception as e:
            showinfo("Error", f"Unable to read file: {e}")
            return

    excel_output = result_prefix + "_processed.xlsx"
    txt_output = result_prefix + "_processed.txt"

    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        for res in results:
            sheet_name = res["sheet_name"]
            header_str = f"Book: {res['longname']}" if res["longname"] else sheet_name
            worksheet = writer.book.add_worksheet(sheet_name)
            worksheet.merge_range('A1:E1', header_str)
            res["df"].to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)
            n_data = len(res["df"])
            start_reg = n_data + 4
            res["df_reg"].to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_reg)
            n_reg = len(res["df_reg"])
            start_sum = start_reg + n_reg + 2

            worksheet.write(start_sum, 0, "Model")
            worksheet.write(start_sum, 1, "Slope")
            worksheet.write(start_sum, 2, "Intercept")
            worksheet.write(start_sum, 3, "R-squared")
            worksheet.write(start_sum, 4, "Activation Energy (eV)")
            summary = res["summary"]
            worksheet.write(start_sum + 1, 0, "Arrhenius: ln(R) vs 1/T")
            worksheet.write(start_sum + 1, 1, summary.get("Slope_Arr", ""))
            worksheet.write(start_sum + 1, 2, summary.get("Intercept_Arr", ""))
            worksheet.write(start_sum + 1, 3, summary.get("R2_Arr", ""))
            worksheet.write(start_sum + 1, 4, summary.get("Ea_Arr", ""))
            worksheet.write(start_sum + 2, 0, "ln(R) vs (1/T)^0.5")
            worksheet.write(start_sum + 2, 1, summary.get("Slope_Half", ""))
            worksheet.write(start_sum + 2, 2, summary.get("Intercept_Half", ""))
            worksheet.write(start_sum + 2, 3, summary.get("R2_Half", ""))
            worksheet.write(start_sum + 2, 4, summary.get("Ea_Half", ""))
            worksheet.write(start_sum + 3, 0, "ln(R) vs (1/T)^0.25")
            worksheet.write(start_sum + 3, 1, summary.get("Slope_Quarter", ""))
            worksheet.write(start_sum + 3, 2, summary.get("Intercept_Quarter", ""))
            worksheet.write(start_sum + 3, 3, summary.get("R2_Quarter", ""))
            worksheet.write(start_sum + 3, 4, summary.get("Ea_Quarter", ""))

            n_reg_rows = len(res["df_reg"])
            arr_min_x = res["df_reg"]["1/T (1/K)"].min()
            arr_max_x = res["df_reg"]["1/T (1/K)"].max()
            arr_min_y = res["df_reg"]["ln(R)"].min()
            arr_max_y = res["df_reg"]["ln(R)"].max()
            half_min_x = res["df_reg"]["invT_half"].min()
            half_max_x = res["df_reg"]["invT_half"].max()
            quarter_min_x = res["df_reg"]["invT_quarter"].min()
            quarter_max_x = res["df_reg"]["invT_quarter"].max()
            t_min = res["df_reg"]["T, K"].min()
            t_max = res["df_reg"]["T, K"].max()
            composite_min_y = min(res["df_reg"]["Fit_Arrhenius"].min(),
                                  res["df_reg"]["Fit_Half"].min(),
                                  res["df_reg"]["Fit_Quarter"].min())
            composite_max_y = max(res["df_reg"]["Fit_Arrhenius"].max(),
                                  res["df_reg"]["Fit_Half"].max(),
                                  res["df_reg"]["Fit_Quarter"].max())

            try:
                chart_arr = writer.book.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                x_col = res["df_reg"].columns.get_loc("1/T (1/K)")
                y_col = res["df_reg"].columns.get_loc("ln(R)")
                chart_arr.add_series({
                    'name': f"Arrhenius (Ea: {summary['Ea_Arr']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, x_col, start_reg + n_reg_rows, x_col],
                    'values': [sheet_name, start_reg + 1, y_col, start_reg + n_reg_rows, y_col],
                    'marker': {'type': 'circle', 'size': 4},
                    'trendline': {'type': 'linear', 'display_equation': False, 'display_r_squared': False}
                })
                chart_arr.set_title({'name': "ln(R) vs 1/T"})
                chart_arr.set_x_axis({'name': "1/T (1/K)", 'min': arr_min_x, 'max': arr_max_x})
                chart_arr.set_y_axis({'name': "ln(R)", 'min': arr_min_y, 'max': arr_max_y})
                worksheet.insert_chart('H2', chart_arr)
            except Exception as e:
                print(f"[Excel Chart Error] Arrhenius: {e}")

            try:
                chart_half = writer.book.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                x_col_half = res["df_reg"].columns.get_loc("invT_half")
                chart_half.add_series({
                    'name': f"(1/T)^0.5 (Ea: {summary['Ea_Half']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, x_col_half, start_reg + n_reg_rows, x_col_half],
                    'values': [sheet_name, start_reg + 1, y_col, start_reg + n_reg_rows, y_col],
                    'marker': {'type': 'circle', 'size': 4},
                    'trendline': {'type': 'linear', 'display_equation': False, 'display_r_squared': False}
                })
                chart_half.set_title({'name': "ln(R) vs (1/T)^0.5"})
                chart_half.set_x_axis({'name': "(1/T)^0.5", 'min': half_min_x, 'max': half_max_x})
                chart_half.set_y_axis({'name': "ln(R)", 'min': arr_min_y, 'max': arr_max_y})
                worksheet.insert_chart('H20', chart_half)
            except Exception as e:
                print(f"[Excel Chart Error] (1/T)^0.5: {e}")

            try:
                chart_quarter = writer.book.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                x_col_quarter = res["df_reg"].columns.get_loc("invT_quarter")
                chart_quarter.add_series({
                    'name': f"(1/T)^0.25 (Ea: {summary['Ea_Quarter']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, x_col_quarter, start_reg + n_reg_rows, x_col_quarter],
                    'values': [sheet_name, start_reg + 1, y_col, start_reg + n_reg_rows, y_col],
                    'marker': {'type': 'circle', 'size': 4},
                    'trendline': {'type': 'linear', 'display_equation': False, 'display_r_squared': False}
                })
                chart_quarter.set_title({'name': "ln(R) vs (1/T)^0.25"})
                chart_quarter.set_x_axis({'name': "(1/T)^0.25", 'min': quarter_min_x, 'max': quarter_max_x})
                chart_quarter.set_y_axis({'name': "ln(R)", 'min': arr_min_y, 'max': arr_max_y})
                worksheet.insert_chart('H38', chart_quarter)
            except Exception as e:
                print(f"[Excel Chart Error] (1/T)^0.25: {e}")

            try:
                chart_all = writer.book.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                t_col = res["df_reg"].columns.get_loc("T, K")
                fit_arr_col = res["df_reg"].columns.get_loc("Fit_Arrhenius")
                fit_half_col = res["df_reg"].columns.get_loc("Fit_Half")
                fit_quarter_col = res["df_reg"].columns.get_loc("Fit_Quarter")
                chart_all.add_series({
                    'name': f"Arrhenius (Ea: {summary['Ea_Arr']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, t_col, start_reg + n_reg_rows, t_col],
                    'values': [sheet_name, start_reg + 1, fit_arr_col, start_reg + n_reg_rows, fit_arr_col],
                    'marker': {'type': 'none'},
                    'line': {'width': 2}
                })
                chart_all.add_series({
                    'name': f"(1/T)^0.5 (Ea: {summary['Ea_Half']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, t_col, start_reg + n_reg_rows, t_col],
                    'values': [sheet_name, start_reg + 1, fit_half_col, start_reg + n_reg_rows, fit_half_col],
                    'marker': {'type': 'none'},
                    'line': {'width': 2}
                })
                chart_all.add_series({
                    'name': f"(1/T)^0.25 (Ea: {summary['Ea_Quarter']:.3f} eV)",
                    'categories': [sheet_name, start_reg + 1, t_col, start_reg + n_reg_rows, t_col],
                    'values': [sheet_name, start_reg + 1, fit_quarter_col, start_reg + n_reg_rows, fit_quarter_col],
                    'marker': {'type': 'none'},
                    'line': {'width': 2}
                })
                chart_all.set_title({'name': "Composite Fit: All Models vs Temperature"})
                chart_all.set_x_axis({'name': "Temperature (K)", 'min': t_min, 'max': t_max})
                composite_min_y = min(res["df_reg"]["Fit_Arrhenius"].min(), res["df_reg"]["Fit_Half"].min(),
                                      res["df_reg"]["Fit_Quarter"].min())
                composite_max_y = max(res["df_reg"]["Fit_Arrhenius"].max(), res["df_reg"]["Fit_Half"].max(),
                                      res["df_reg"]["Fit_Quarter"].max())
                chart_all.set_y_axis({'name': "ln(R)", 'min': composite_min_y, 'max': composite_max_y})
                worksheet.insert_chart('H56', chart_all)
            except Exception as e:
                print(f"[Excel Chart Error] Composite: {e}")
        with open(txt_output, "w", encoding="utf-8") as f:
            f.write(df.to_csv(sep='\t', index=False))

    graph_window = Toplevel(root)
    graph_window.title("Graphical Results")
    fig = Figure(figsize=(8, 16), dpi=100)

    ax1 = fig.add_subplot(411)
    ax1.plot(df_reg["1/T (1/K)"], df_reg["ln(R)"], 'o', label="Data")
    ax1.plot(df_reg["1/T (1/K)"], intercept_arr + slope_arr * df_reg["1/T (1/K)"],
             '-', label=f"Fit (Ea={slope_arr * k_B:.3f} eV)")
    ax1.set_title("Arrhenius Model: ln(R) vs 1/T")
    ax1.set_xlabel("1/T (1/K)")
    ax1.set_ylabel("ln(R)")
    ax1.legend(title=results[0]["longname"])
    ax1.grid(True)

    ax2 = fig.add_subplot(412)
    ax2.plot(df_reg["invT_half"], df_reg["ln(R)"], 'o', label="Data")
    ax2.plot(df_reg["invT_half"], intercept_half + slope_half * df_reg["invT_half"],
             '-', label=f"Fit (Ea={slope_half * k_B:.3f} eV)")
    ax2.set_title("ln(R) vs (1/T)^0.5")
    ax2.set_xlabel("(1/T)^0.5")
    ax2.set_ylabel("ln(R)")
    ax2.legend(title=results[0]["longname"])
    ax2.grid(True)

    ax3 = fig.add_subplot(413)
    ax3.plot(df_reg["invT_quarter"], df_reg["ln(R)"], 'o', label="Data")
    ax3.plot(df_reg["invT_quarter"], intercept_quarter + slope_quarter * df_reg["invT_quarter"],
             '-', label=f"Fit (Ea={slope_quarter * k_B:.3f} eV)")
    ax3.set_title("ln(R) vs (1/T)^0.25")
    ax3.set_xlabel("(1/T)^0.25")
    ax3.set_ylabel("ln(R)")
    ax3.legend(title=results[0]["longname"])
    ax3.grid(True)

    ax4 = fig.add_subplot(414)
    ax4.plot(df_reg["T, K"], df_reg["Fit_Arrhenius"], '-', label=f"Arrhenius, Ea={slope_arr * k_B:.3f} eV")
    ax4.plot(df_reg["T, K"], df_reg["Fit_Half"], '-', label=f"(1/T)^0.5, Ea={slope_half * k_B:.3f} eV")
    ax4.plot(df_reg["T, K"], df_reg["Fit_Quarter"], '-', label=f"(1/T)^0.25, Ea={slope_quarter * k_B:.3f} eV")
    ax4.set_title("Composite Fit: All Models vs Temperature")
    ax4.set_xlabel("Temperature (K)")
    ax4.set_ylabel("ln(R)")
    ax4.legend()
    ax4.grid(True)

    canvas = FigureCanvasTkAgg(fig, master=graph_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)
    toolbar = NavigationToolbar2Tk(canvas, graph_window)
    toolbar.update()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    if ext in [".opj", ".opju"]:
        try:
            for i, col in enumerate(df.columns):
                wks.Columns.Add()
                wks.Columns(i).SetName(col)
                wks.SetColValues(i, df[col].tolist())
            origin.Execute("plotxy iy:=(1,2) plot:=200;")
        except Exception as e:
            print(f"Could not load data into Origin: {e}")

    showinfo("Success", "File processed successfully!")


root = Tk()
root.title("Thermal Data Analyzer")
root.geometry("300x180")

Label(root, text="Select a CSV, TXT or Origin file to analyze").pack(pady=10)
Button(root, text="Open File", command=process_file).pack(pady=5)
Button(root, text="Exit", command=root.quit).pack(pady=5)

root.mainloop()
