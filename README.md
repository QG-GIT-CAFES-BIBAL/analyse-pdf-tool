# Analyse PDF Tool

Outil interne pour analyser automatiquement les relevés **Touch N Pay** en PDF et générer un fichier **CSV** exploitable.

---

## 🚀 Fonctionnalités

* Extraction des données via **Poppler (pdftotext)**
* Fallback OCR via **Tesseract** si l’extraction échoue
* Résultats propres dans un CSV (`export_analyse_pdf.csv`)
* Interface console avec barre de progression colorée et résumé final

---

## 📂 Structure

```
AnalysePDF/
│── analyse_pdf.py          # Script principal
│── dist_bundle_ressources/ # Poppler + Tesseract intégrés
│── Mettre les PDF ICI/     # Dossier où déposer les fichiers à analyser
│── analyse_pdf.exe         # Exécutable pour les utilisateurs finaux
│── export_analyse_pdf.csv  # Résultat après analyse
```

---

## 🖥️ Utilisation

1. Copier vos fichiers **PDF** dans le dossier `Mettre les PDF ICI`
2. Lancer l’exécutable **analyse\_pdf.exe**
3. À la fin, un fichier **export\_analyse\_pdf.csv** est généré à la racine du dossier.

---

## 🔧 Développement

Pour regénérer l’exécutable après modification du script :

```bash
pyinstaller --onefile ^
  --add-data "dist_bundle_ressources;dist_bundle_ressources" ^
  analyse_pdf.py
```

---

## 📝 Notes

* Ne pas ajouter les binaires (`dist`, `build`, `dist_bundle_ressources`) dans Git : ils sont exclus via `.gitignore`.
* Le projet est pensé pour être **100% portable** sur n’importe quel poste.

---

## 📄 .gitignore recommandé

```gitignore
# Python cache
__pycache__/
*.pyc
*.pyo

# Build PyInstaller
build/
dist/
*.spec

# Executables et DLL
*.exe
*.dll
*.pyd

# Ressources locales lourdes
dist_bundle_ressources/

# Fichiers générés
export_analyse_pdf.csv
```
