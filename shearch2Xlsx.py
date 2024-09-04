import sys
import json
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QCheckBox, QPushButton, QGridLayout, QWidget, QMessageBox, QScrollArea, QVBoxLayout, QGroupBox, QFileDialog

# Charger les termes depuis un fichier JSON structuré
with open('terms.json', 'r', encoding='utf-8') as file:
    categories = json.load(file)

# Créer une liste de correspondance français -> anglais
terms_fr_to_en = {}
for category, terms in categories.items():
    for en_term, fr_term in terms.items():
        terms_fr_to_en[fr_term] = en_term

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sélectionnez les termes")

        # Créer un widget central et un layout principal
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Créer un widget pour le contenu défilant
        self.scroll_area_widget = QWidget()
        self.scroll_area_layout = QVBoxLayout(self.scroll_area_widget)
        self.scroll_area_widget.setLayout(self.scroll_area_layout)

        # Créer un dictionnaire pour stocker les cases à cocher
        self.checkboxes = {}

        # Nombre de colonnes dans chaque groupe
        num_columns = 3

        # Liste des couleurs pour les QGroupBox
        colors = [
            "#D3D3D3",  # Gris Clair
            "#C0C0C0",  # Gris Argenté
            "#DCDCDC",  # Gris Ardoise
            "#FFFFE0",  # Jaune Clair
            "#FFFACD",  # Jaune Pâle
            "#F5F5DC",  # Beige Clair
            "#FFDAB9",  # Pêche Clair
            "#FFBF80",  # Abricot Clair
            "#FDFD96",  # Jaune Pâle
            "#FFFDD0"   # Crème
        ]

        # Créer les sections et ajouter les cases à cocher pour chaque terme en français
        for idx, (category, terms) in enumerate(categories.items()):
            # Créer un groupe pour chaque catégorie
            group_box = QGroupBox(category)
            group_box_layout = QGridLayout()
            group_box.setLayout(group_box_layout)

            # Appliquer la couleur de fond et le style du titre
            color = colors[idx % len(colors)]  # Utiliser une couleur différente pour chaque groupe
            group_box.setStyleSheet(f"""
                background-color: {color}; 
                border: 1px solid black;                
                font-size: 10pt;
                border-radius: 2px;
                padding: 5px;  
            """)

            row = 0
            column = 0

            # Ajouter les cases à cocher au layout de groupe
            for en_term, fr_term in terms.items():
                unique_id = f"{category}_{fr_term}"  # Générer un identifiant unique
                checkbox = QCheckBox(fr_term)
                checkbox.setObjectName(unique_id)  # Définir un nom d'objet unique pour la case à cocher
                group_box_layout.addWidget(checkbox, row, column)
                self.checkboxes[unique_id] = checkbox  # Associe l'identifiant unique avec le QCheckBox
                column += 1
                if column >= num_columns:
                    column = 0
                    row += 1

            # Ajouter le groupe au layout principal de la zone défilante
            self.scroll_area_layout.addWidget(group_box)

        # Ajouter le bouton de validation
        self.validate_button = QPushButton("Valider")
        self.validate_button.clicked.connect(self.on_validate)

        # Créer une zone de défilement
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.scroll_area_widget)

        # Ajouter la zone de défilement et le bouton de validation au layout principal
        self.main_layout.addWidget(self.scroll_area)
        self.main_layout.addWidget(self.validate_button)

    def on_validate(self):
        selected_terms = []

        for unique_id, checkbox in self.checkboxes.items():
            if checkbox.isChecked():
                fr_term = unique_id.split('_', 1)[1]
                english_term = terms_fr_to_en.get(fr_term, "Inconnu")
                selected_terms.append(english_term)

        if selected_terms:
            result = ", ".join(selected_terms)
            
            # Copier le résultat dans le presse-papiers
            clipboard = QApplication.clipboard()
            clipboard.setText(result)

            QMessageBox.information(self, "Traductions", f"Termes en anglais : {result}\n\nLe résultat a été copié dans le presse-papiers.")

            # Appel de la fonction search_and_save avec les termes sélectionnés
            search_and_save(selected_terms)

        else:
            QMessageBox.warning(self, "Aucun terme sélectionné", "Veuillez sélectionner au moins un terme.")

# Fonction pour parcourir les fichiers et rechercher les correspondances
def search_and_save(terms):
    folder_selected = QFileDialog.getExistingDirectory(None, "Sélectionnez le dossier contenant les fichiers .xlsx")
    
    if not folder_selected:
        return
    
    all_results = []

    for root, dirs, files in os.walk(folder_selected):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                try:
                    # Lire le fichier Excel sans considérer la première ligne comme un header
                    df = pd.read_excel(file_path, header=None)
                    
                    # Rechercher les termes sélectionnés dans toutes les colonnes
                    for term in terms:
                        result = df[df.apply(lambda row: row.astype(str).str.contains(term, case=False).any(), axis=1)]
                        if not result.empty:
                            all_results.append(result)
                
                except Exception as e:
                    print(f"Erreur lors de la lecture du fichier {file_path}: {e}")
    
    if all_results:
        # Concaténer tous les résultats en un seul DataFrame sans réinitialiser les index
        final_df = pd.concat(all_results, ignore_index=True)
        
        # Sauvegarder les résultats dans un nouveau fichier .xlsx
        save_path = QFileDialog.getSaveFileName(None, "Sauvegarder le fichier", "", "Excel files (*.xlsx)")[0]
        if save_path:
            # Sauvegarder sans header ni index
            final_df.to_excel(save_path, index=False, header=False)
            QMessageBox.information(None, "Succès", "Les résultats ont été sauvegardés avec succès.")
        else:
            QMessageBox.warning(None, "Annulé", "La sauvegarde du fichier a été annulée.")
    else:
        QMessageBox.information(None, "Aucun résultat", "Aucun terme correspondant n'a été trouvé.")


# Lancer l'application
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
