import tkinter as tk
from tkinter import ttk
import pandas as pd
from PIL import Image, ImageTk

chemin = "C:/Users/jules/OneDrive/Desktop/Projet Python/OCO_results.csv"

df = pd.read_csv(chemin, encoding="cp1252", sep=";")
produits = df["Produit"].unique().tolist()

#Recherche et copie de l'ISIN qui est présent dans l'EXCEL du chemin.csv

def isin(chemin, strike, maturite, produit, option):
    filtre = pd.read_csv(chemin, encoding="cp1252", sep=";")
    try:
        strike = float(strike)
    except ValueError:
        strike = strike
    filtrage = filtre[(filtre["Strike"] == strike) &
                      (filtre["Maturité"] == maturite) &
                      (filtre["Produit"] == produit) &
                      (filtre["Option"] == option)]
    return filtrage["ISIN"].values[0] if not filtrage.empty else None

def chercher_isin():
    try:
        code = isin(chemin, combo_strike.get(), combo_maturite.get(), combo_produit.get(), combo_option.get())
        if code:
            fenetre.clipboard_clear() 
            fenetre.clipboard_append(code) 
            fenetre.update()
            label_result.config(text=f"ISIN copié : {code}", fg="green")
        else:
            label_result.config(text="Aucun résultat trouvé", fg="orange")
    except Exception as e:
        label_result.config(text=f"Erreur : {str(e)}", fg="red")

# Màj maturités strikes etc.. selon le produit
def update_maturites(event=None):
    produit = combo_produit.get()
    if produit:
        df_filtre = df[df["Produit"] == produit]
        maturites = sorted(df_filtre["Maturité"].unique().tolist())
        combo_maturite['values'] = maturites
        combo_maturite.set('')
        combo_option.set('')
        combo_strike.set('')
        combo_option['values'] = []
        combo_strike['values'] = []
        if maturites:
            combo_maturite.current(0)
            update_options()

def update_options(event=None):
    produit = combo_produit.get()
    maturite = combo_maturite.get()
    if produit and maturite:
        df_filtre = df[(df["Produit"] == produit) &
                       (df["Maturité"] == maturite)]
        options = sorted(df_filtre["Option"].unique().tolist())
        combo_option['values'] = options
        combo_option.set('')
        combo_strike.set('')
        combo_strike['values'] = []
        if options:
            combo_option.current(0)
            update_strikes()
    else:
        combo_option.set('')
        combo_strike.set('')
        combo_option['values'] = []
        combo_strike['values'] = []

def update_strikes(event=None):
    produit = combo_produit.get()
    maturite = combo_maturite.get()
    option = combo_option.get()
    if produit and maturite and option:
        df_filtre = df[(df["Produit"] == produit) &
                       (df["Maturité"] == maturite) &
                       (df["Option"] == option)]
        strikes = sorted(df_filtre["Strike"].unique().tolist())
        combo_strike['values'] = strikes
        combo_strike.set('')
        if strikes:
            combo_strike.current(0)
    else:
        combo_strike.set('')
        combo_strike['values'] = []

#Création de l'interface

fenetre = tk.Tk()
fenetre.title("Recherche ISIN")
fenetre.geometry("500x450")
fenetre.eval('tk::PlaceWindow . center')

style = ttk.Style()
style.configure('TCombobox', padding=5, font=('Arial', 10))
style.configure('TButton', padding=10, font=('Arial', 10, 'bold'))

main_frame = ttk.Frame(fenetre, padding=10)
main_frame.pack(expand=True)

ttk.Label(main_frame, text="Produit").grid(row=0, column=0, pady=3, sticky="e")
combo_produit = ttk.Combobox(main_frame, values=produits, state="readonly", width=30)
combo_produit.grid(row=0, column=1, pady=3)
combo_produit.bind("<<ComboboxSelected>>", update_maturites)

ttk.Label(main_frame, text="Maturité").grid(row=1, column=0, pady=3, sticky="e")
combo_maturite = ttk.Combobox(main_frame, state="readonly", width=30)
combo_maturite.grid(row=1, column=1, pady=3)
combo_maturite.bind("<<ComboboxSelected>>", update_options)

ttk.Label(main_frame, text="Option").grid(row=2, column=0, pady=3, sticky="e")
combo_option = ttk.Combobox(main_frame, state="readonly", width=30)
combo_option.grid(row=2, column=1, pady=3)
combo_option.bind("<<ComboboxSelected>>", update_strikes)

ttk.Label(main_frame, text="Strike").grid(row=3, column=0, pady=3, sticky="e")
combo_strike = ttk.Combobox(main_frame, state="readonly", width=30)
combo_strike.grid(row=3, column=1, pady=3)

ttk.Button(main_frame, text="Valider", command=chercher_isin).grid(row=4, column=0, columnspan=2, pady=(12, 5))

label_result = tk.Label(main_frame, text="", font=('Arial', 11))
label_result.grid(row=5, column=0, columnspan=2, pady=(5, 0))

if produits:
    combo_produit.current(0)
    update_maturites()

fenetre.mainloop()
