import tkinter as tk
from tkinter import filedialog
import pandas as pd
from PIL import Image, ImageTk
from xlsxwriter.utility import xl_range
from datetime import datetime

root = tk.Tk()
root.title("Tax Transaction History")
root.geometry("800x350")
image = Image.open("369.png")
background_image = ImageTk.PhotoImage(image)
background_label = tk.Label(root, image=background_image)
background_label.place(relx=1, rely=0, anchor='ne')
image1 = Image.open("326.png")
background_image1 = ImageTk.PhotoImage(image1)
background_label1 = tk.Label(root, image=background_image1)
background_label1.place(x=0, y=0, anchor="nw")

def process_csv_file(file_path,variable_value):
    global df
    df = pd.read_csv(file_path, sep=';', index_col=False)
    df = df.drop('#', axis=1)
    df['Transaction Time'] = pd.to_datetime(df['Transaction Time'])
    df['Transaction Time'] = df['Transaction Time'].apply(lambda x: datetime.strftime(x, '%d.%m.%Y %H:%M:%S'))

    def format_player_idnp(x):
        try:
            return '{:.0f}'.format(float(x))
        except (ValueError, TypeError):
            return str(x)

    df['Player IDNP'] = df['Player IDNP'].apply(format_player_idnp)
    pd.set_option('display.float_format', '{:.0f}'.format)
    df['Company Fiscal Code'] = df['Company Fiscal Code'].apply(lambda x: '{:.0f}'.format(x))
    df = df.groupby('Tax Session ID').apply(lambda x: pd.concat([x, pd.DataFrame([[]])], ignore_index=True))
    df = df.reset_index(drop=True)
    df_sum = df.groupby(['Tax Session ID'])[['Deposit Amount', 'Withdrawals']].sum().reset_index()


    for index, row in df_sum.iterrows():
        session_id = row['Tax Session ID']
        deposit_sum = row['Deposit Amount']
        withdrawals_sum = row['Withdrawals']
        row_index = df.loc[df['Tax Session ID'] == session_id].index[-1] + 1
        df.loc[row_index, 'Tax Session ID'] = session_id
        df.loc[row_index, 'Deposit Amount'] = deposit_sum
        df.loc[row_index, 'Withdrawals'] = withdrawals_sum
        dif = (withdrawals_sum - deposit_sum)/100
        df.loc[row_index, 'Paid Win'] = str(dif) + "MDL"
        if deposit_sum > withdrawals_sum:
            df.loc[row_index, 'Unused Deposit'] = "Diferenta -"
        else:
            df.loc[row_index, 'Unused Deposit'] = "Diferenta +"
        df.loc[row_index, 'Deposit Amount'] = str(deposit_sum / 100) + " MDL"
        df.loc[row_index, 'Withdrawals'] = str(withdrawals_sum / 100) + " MDL"
    total_deposit = df_sum['Deposit Amount'].sum() / 100
    total_withdrawals = df_sum['Withdrawals'].sum() / 100
    df = pd.concat([df, pd.DataFrame({'Player Id': '', 'Player Name': ''}, index=[len(df)]*5)], ignore_index=True)
    total_deposit = df_sum['Deposit Amount'].sum() / 100
    total_withdrawals = df_sum['Withdrawals'].sum() / 100
    bala = total_withdrawals - total_deposit
    bala1 = bala + float(variable_value)
    num_unique_tax_session_ids = df['Tax Session ID'].nunique()
    dif_plus_count = df['Unused Deposit'].value_counts()['Diferenta +']
    dif_minus_count = df['Unused Deposit'].value_counts()['Diferenta -']

    df.loc[len(df) - 4, 'Player Id'] = 'Castigul Jucatorului'

    df.loc[len(df)-3, 'Player Id'] = 'Total Depuneri'
    df.loc[len(df)-2, 'Player Id'] = 'Total Retrageri'
    df.loc[len(df)-1, 'Player Id'] = 'Total Extrageri inclusiv castigul'
    df.loc[len(df)-0, 'Player Id'] = 'Balanta'
    df.loc[len(df)+1, 'Player Id'] = 'Balanta inclusiv castigul'
    df.loc[len(df) + 2, 'Player Id'] = 'Sesiuni de joc'
    df.loc[len(df) + 3, 'Player Id'] = 'Sesiuni de joc Pozitive'
    df.loc[len(df) + 3, 'Player Id'] = 'Sesiuni de joc Negative'

    numar_formatat_0 = f"{float(variable_value):,.2f}"
    df.loc[len(df) - 9, 'Player Name'] = str(numar_formatat_0) + ' MDL'

    numar_formatat = f"{total_deposit:,.2f}"
    df.loc[len(df) - 8, 'Player Name'] = str(numar_formatat) + ' MDL'
    numar_formatat_1 = f"{total_withdrawals:,.2f}"
    df.loc[len(df) - 7, 'Player Name'] = str(numar_formatat_1) + ' MDL'
    tot = total_withdrawals + float(variable_value)
    numar_formatat_2 = f"{tot:,.2f}"
    df.loc[len(df) - 6, 'Player Name'] = str(numar_formatat_2) + ' MDL'
    bal = total_withdrawals - total_deposit
    numar_formatat_3 = f"{bal:,.2f}"
    df.loc[len(df) - 5, 'Player Name'] = str(numar_formatat_3) + ' MDL'
    bal_1 = bal + float(variable_value)
    numar_formatat_4 = f"{bal_1:,.2f}"
    df.loc[len(df) - 3, 'Player Name'] = str(numar_formatat_4) + ' MDL'
    df.loc[len(df) - 1, 'Player Name'] = num_unique_tax_session_ids
    df.loc[len(df) + 1, 'Player Name'] = dif_plus_count
    df.loc[len(df) + 2, 'Player Name'] = dif_minus_count
    player_name = df['Player Name'][0]
    save_excel_file(f"TaxTransactionHistory_{player_name}.xlsx")


def save_excel_file(file_name=None, player_name=None):
    if file_name is None:
        file_name = f"TaxTransactionHistory_{player_name}.xlsx"
    root = tk.Tk()
    root.withdraw()
    player_name = df['Player Name'][0]
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[('Excel Files', '*.xlsx')], initialfile=file_name)
    if file_path:
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        medium_11 = workbook.add_format({'font_name': 'Calibri', 'font_size': 11,
                                         'align': 'center', 'valign': 'vcenter'
                                        })
        yellow_bg = workbook.add_format({'bg_color': 'yellow', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': True})

        for row_index, amount in enumerate(df['Deposit Amount'], start=1):
            if isinstance(amount, float):
                amount_str = str(amount)
            else:
                amount_str = amount
            if 'MDL' in amount_str:
                worksheet.write(row_index, df.columns.get_loc('Deposit Amount'), amount_str, yellow_bg)

        for row_index, amount in enumerate(df['Withdrawals'], start=1):
            if isinstance(amount, float):
                amount_str = str(amount)
            else:
                amount_str = amount
            if 'MDL' in amount_str:
                worksheet.write(row_index, df.columns.get_loc('Withdrawals'), amount_str, yellow_bg)

        for row_index, amount in enumerate(df['Paid Win'], start=1):
            if isinstance(amount, float):
                amount_str = str(amount)
            else:
                amount_str = amount
            if 'MDL' in amount_str:
                worksheet.write(row_index, df.columns.get_loc('Paid Win'), amount_str, yellow_bg)

        for row_index, paid_win in enumerate(df['Unused Deposit'], start=1):
            if isinstance(paid_win, float):
                paid_win_str = str(paid_win)
            else:
                paid_win_str = paid_win
            if 'Diferenta +' in paid_win_str or 'Diferenta -' in paid_win_str:
                worksheet.write(row_index, df.columns.get_loc('Unused Deposit'), paid_win_str, yellow_bg)

        y = len(df.index)
        x = len(df.columns)
        cell_range = xl_range(0, 0, y - 10, x - 1)
        worksheet.add_table(cell_range, {'style': 'Table Style Medium 11', 'header_row': False})
        worksheet.set_column(1, len(df.columns) - 1, 18, medium_11)
        z = len(df.index) - 8
        a = len(df.columns) - 12
        z1 = len(df.index)
        a1 = len(df.columns) - 9
        cell_range1 = xl_range(z, a, z1, a1 - 1)
        format1 = workbook.add_format({'bg_color': 'FF8C00', 'bold': True, 'border': True})
        worksheet.conditional_format(cell_range1, {'type': 'no_blanks', 'format': format1})
        writer.close()

def exit_application():
    root.quit()

def on_entry_click(event=None):
    if variable_entry.get() == 'Introduceți castigul':
        variable_entry.delete(0, "end")
        variable_entry.insert(0, '')
        variable_entry.config(fg='black')

variable_entry = tk.Entry(root, font=("Helvetica", 14), fg='gray', insertwidth=1)
variable_entry.insert(0, 'Introduceți castigul')
variable_entry.bind('<FocusIn>', on_entry_click)
variable_entry.pack(side="top", padx=20, pady=10)


def get_variable_value():
    variable_value = variable_entry.get()
    file_path = filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=[("CSV files", "*.csv")])
    process_csv_file(file_path, variable_value)

    on_entry_click()
    variable_entry.delete(0, "end")
    variable_entry.insert(0, 'Introduceți castigul')
    variable_entry.config(fg='gray', state='normal')
    variable_entry.config(state='disabled')
    success_label.config(text="Raportul a fost creat cu succes")

variable_button = tk.Button(root, text="START", command=get_variable_value , height=3, width=20, font=("Helvetica", 14, "bold"), bg="#c8ffb0")
variable_button.pack(side="top", padx=10, pady=10)
exit_button = tk.Button(root, text="EXIT", fg="red", command=exit_application, height=3, width=20, font=("Helvetica", 14, "bold"), bg="#ffa07a")
exit_button.pack(side="top", padx=10, pady=10)
author_label = tk.Label(root, text=" By Spînu Ion 2023   ", fg="gray", font=("Helvetica", 9))
author_label.pack(side="bottom", anchor="se")
label_font = ("Helvetica", 10, "bold")
success_label = tk.Label(root, text="", font=label_font, fg="green")
success_label.pack(side="bottom", padx=20, pady=10)

root.mainloop()
