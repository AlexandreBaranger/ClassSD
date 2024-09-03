import tkinter as tk
from tkinter import messagebox
import json

# Charger les termes depuis un fichier JSON structuré
with open('terms.json', 'r', encoding='utf-8') as file:
    categories = json.load(file)

# Créer une liste de correspondance français -> anglais
terms_fr_to_en = {}
for category, terms in categories.items():
    for en_term, fr_term in terms.items():
        terms_fr_to_en[fr_term] = en_term

# Fonction qui sera appelée lors de la validation
def on_validate():
    selected_terms = []
    for term, var in checkboxes.items():
        if var.get() == 1:
            selected_terms.append(terms_fr_to_en.get(term, "Inconnu"))

    if selected_terms:
        result = ", ".join(selected_terms)
        
        # Copier le résultat dans le presse-papiers
        root.clipboard_clear()  # Efface le contenu actuel du presse-papiers
        root.clipboard_append(result)  # Ajoute le résultat au presse-papiers

        messagebox.showinfo("Traductions", f"Termes en anglais : {result}\n\nLe résultat a été copié dans le presse-papiers.")
    else:
        messagebox.showwarning("Aucun terme sélectionné", "Veuillez sélectionner au moins un terme.")

# Créer la fenêtre principale
root = tk.Tk()
root.title("Sélectionnez les termes")

# Créer un dictionnaire pour stocker les variables des cases à cocher
checkboxes = {}

# Ajouter les cases à cocher pour chaque terme en français
for term in terms_fr_to_en.keys():
    var = tk.IntVar()
    checkbox = tk.Checkbutton(root, text=term, variable=var)
    checkbox.pack(anchor='w')
    checkboxes[term] = var

# Ajouter le bouton de validation
validate_button = tk.Button(root, text="Valider", command=on_validate)
validate_button.pack()

# Lancer l'interface utilisateur
root.mainloop()
